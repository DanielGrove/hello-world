// Micha Religious School - consolidated attendance
//
// Reads every Microsoft Forms response workbook in one OneDrive/SharePoint folder,
// turns each class's Likert grid (one column per student) into one row per
// student per date, and stacks all classes into a single table:
//
//   Class | Date | Teacher | Student | Status | Submitted
//
// Paste into Excel: Data > Get Data > From Other Sources > Blank Query >
// Advanced Editor. Edit the SETTINGS block, then Close & Load.

let
    // ---------------- SETTINGS ----------------

    // Root URL of the OneDrive or SharePoint site that holds the response workbooks.
    //   OneDrive:   "https://micha-my.sharepoint.com/personal/office_micha_org/"
    //   SharePoint: "https://micha.sharepoint.com/sites/ReligiousSchool/"
    SiteUrl = "https://YOURTENANT-my.sharepoint.com/personal/YOUR_ACCOUNT/",

    // Only files whose folder path contains this text are read.
    FolderContains = "/Apps/Microsoft Forms/",

    // Only files whose name contains this text are read ("" = every .xlsx).
    // Also stripped from the class name, so "Kitah Gimel Attendance" -> "Kitah Gimel".
    FileNameContains = "Attendance",

    // Where the class name comes from: "File" (workbook name) or "Folder" (its folder).
    ClassFrom = "File",

    // The date and teacher questions are found by these words in their titles,
    // so they do not have to be worded identically on every form.
    DateKeyword = "date",
    TeacherKeyword = "teacher",

    // If the student columns come through as "<Likert title>.<Student>", put the
    // Likert title here and it is stripped (e.g. "Attendance"). Otherwise leave "".
    LikertTitle = "",

    // ------------------------------------------

    // Columns Forms adds to every response workbook; never treated as students.
    MetaColumns = {"id", "start time", "completion time", "email", "name", "last modified time"},

    LastFolder = (path as text) as text =>
        List.Last(List.Select(Text.Split(path, "/"), each _ <> "")),

    ClassFor = (fileName as text, folderPath as text) as text =>
        let
            Base = if ClassFrom = "Folder"
                then LastFolder(folderPath)
                else Text.BeforeDelimiter(fileName, ".", {0, RelativePosition.FromEnd}),
            // Downloaded copies are named like "Kitah Gimel Attendance(1-42)"
            NoRange = if Text.EndsWith(Base, ")") and Text.Contains(Base, "(")
                then Text.BeforeDelimiter(Base, "(", {0, RelativePosition.FromEnd})
                else Base,
            NoWord = if FileNameContains = "" then NoRange
                else Text.Replace(NoRange, FileNameContains, "")
        in
            Text.Trim(NoWord),

    ReadOne = (content as binary, className as text) as table =>
        let
            Book = Excel.Workbook(content, true),
            Tables = Table.SelectRows(Book, each [Kind] = "Table"),
            Raw = if Table.RowCount(Tables) > 0
                then Tables{0}[Data]
                else Table.SelectRows(Book, each [Kind] = "Sheet"){0}[Data],
            Cols = Table.ColumnNames(Raw),
            IsMeta = (c as text) => List.Contains(MetaColumns, Text.Lower(Text.Trim(c))),
            FindCol = (kw as text) => List.First(
                List.Select(Cols, each not IsMeta(_) and Text.Contains(_, kw, Comparer.OrdinalIgnoreCase)),
                null),
            DateCol = FindCol(DateKeyword),
            TeacherCol = FindCol(TeacherKeyword),
            CompletionCol = List.First(List.Select(Cols, each Text.Lower(Text.Trim(_)) = "completion time"), null),
            StudentCols = List.Select(Cols, each not IsMeta(_) and _ <> DateCol and _ <> TeacherCol),

            AddClass = Table.AddColumn(Raw, "__Class", each className),
            AddDate = Table.AddColumn(AddClass, "__Date", each if DateCol = null then null else Record.Field(_, DateCol)),
            AddTeacher = Table.AddColumn(AddDate, "__Teacher", each if TeacherCol = null then null else Record.Field(_, TeacherCol)),
            AddSubmitted = Table.AddColumn(AddTeacher, "__Submitted", each if CompletionCol = null then null else Record.Field(_, CompletionCol)),
            Fixed = {"__Class", "__Date", "__Teacher", "__Submitted"},
            Slim = Table.SelectColumns(AddSubmitted, Fixed & StudentCols),
            // Students left blank on a submission produce no row.
            Long = Table.UnpivotOtherColumns(Slim, Fixed, "__Student", "__Status")
        in
            Table.RenameColumns(Long, {
                {"__Class", "Class"}, {"__Date", "Date"}, {"__Teacher", "Teacher"},
                {"__Submitted", "Submitted"}, {"__Student", "Student"}, {"__Status", "Status"}}),

    Source = SharePoint.Files(SiteUrl, [ApiVersion = 15]),
    Files = Table.SelectRows(Source, each
        Text.Contains([Folder Path], FolderContains, Comparer.OrdinalIgnoreCase)
        and Text.EndsWith([Name], ".xlsx", Comparer.OrdinalIgnoreCase)
        and Text.Contains([Name], FileNameContains, Comparer.OrdinalIgnoreCase)
        and not Text.StartsWith([Name], "~$")),

    PerFile = Table.AddColumn(Files, "Rows", each ReadOne([Content], ClassFor([Name], [Folder Path]))),
    Combined = Table.Combine(PerFile[Rows],
        type table [Class = text, Date = any, Teacher = any, Submitted = any, Student = text, Status = any]),

    CleanStudent = Table.TransformColumns(Combined, {{"Student", each
        let s = Text.Trim(_) in
            if LikertTitle <> "" and Text.StartsWith(s, LikertTitle & ".")
            then Text.Trim(Text.AfterDelimiter(s, LikertTitle & "."))
            else s, type text}}),

    Typed = Table.TransformColumnTypes(CleanStudent, {
        {"Date", type date}, {"Teacher", type text}, {"Submitted", type datetime}, {"Status", type text}}, "en-US"),

    // If a teacher submits the same class/date twice, keep only the latest submission.
    Latest = Table.Buffer(Table.Sort(Typed, {{"Submitted", Order.Descending}})),
    Deduped = Table.Distinct(Latest, {"Class", "Date", "Student"}),

    Result = Table.Sort(Deduped, {{"Date", Order.Descending}, {"Class", Order.Ascending}, {"Student", Order.Ascending}})
in
    Result
