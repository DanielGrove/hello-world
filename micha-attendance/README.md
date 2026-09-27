# Micha Religious School attendance: one Excel table from many Forms

Teachers keep using their class's Microsoft Form exactly as they do now. No
Power Automate flows are needed. A master Excel workbook uses one Power Query
(`Attendance.pq`) that reads every form's response workbook and stacks them into
one table:

| Class | Date | Teacher | Student | Status | Submitted |
|---|---|---|---|---|---|
| Kitah Gimel | 2026-09-20 | Morah Rivka | Sarah Cohen | Present | 2026-09-20 10:12 |
| Kitah Gimel | 2026-09-20 | Morah Rivka | Eli Levy | Absent | 2026-09-20 10:12 |
| Kitah Dalet | 2026-09-20 | Mr. Katz | Noa Adler | Late | 2026-09-20 10:40 |

## Why this avoids "a million connections"

The Power Automate Forms trigger works on one form at a time, so a flow-based
setup needs one flow per class, and another flow whenever a class is added.
Microsoft Forms already writes each form's responses to its own Excel workbook
and keeps it in sync. This setup reads those workbooks directly:

```
Kitah Aleph form ─► Kitah Aleph Attendance.xlsx ─┐
Kitah Bet form   ─► Kitah Bet Attendance.xlsx   ─┼─► Power Query ─► Master "Attendance" table
Kitah Gimel form ─► Kitah Gimel Attendance.xlsx ─┘   (one query)
```

When a class is added, its workbook lands in the same folder and shows up on the
next refresh. When a student is added to a class's Likert roster, it becomes one
more column, and the query picks it up automatically.

## One-time setup (about 30 minutes)

### 1. Put all the forms under one owner

The query reads one folder, so all the class forms need to be owned by the same
account, ideally a shared school account (for example `religiousschool@…`) rather
than a teacher's personal account.

- If teachers created their own forms: in Forms, open each form, choose
  **… > Transfer ownership** (or **Move**) to hand it to a Microsoft 365 Group the
  school account belongs to, or have the school account recreate it with
  **Duplicate** from a shared link. The form link, questions and existing responses stay
  the same, so teachers don't notice.
- Name each form `<Class> Attendance`, for example `Kitah Gimel Attendance`. The
  class name in the master table comes from this.

### 2. Make sure every form has a live response workbook

For each form, open **Responses > Open results in Excel** once. For forms created
in the last couple of years, this opens a workbook in the owner's OneDrive, under
**Apps > Microsoft Forms**, that updates on its own as responses come in. (For
group forms, the workbook is in the group's SharePoint document library.)

If a form only offers a **download** (an older form with no live sync), make a
copy of the form: **… > Duplicate**. The copy gets a synced workbook. Send
teachers the new link.

Don't rename or move these workbooks. Moving one breaks the sync.

### 3. Check the question titles (no changes for teachers)

The query finds the columns by keyword, so the wording can vary between forms:

- The **date** question's title contains the word "date" (for example "Class date").
- The **teacher** free-text question's title contains the word "teacher".
- Every other question on the form is treated as a student (the Likert rows).

If a form has an extra question (for example "Notes"), rename it to include
"teacher" (for example "Teacher notes"), or add its title (lowercase) to
`MetaColumns` in the query so it is ignored.

### 4. Build the master workbook

1. Create `Attendance – Master.xlsx` in the school account's OneDrive (anywhere
   *outside* the Forms folder).
2. Open it in **desktop Excel** and choose **Data > Get Data > From Other Sources >
   Blank Query**.
3. Choose **Advanced Editor**, paste the contents of `Attendance.pq`, and edit the
   `SETTINGS` block:
   - `SiteUrl`: open the Forms folder in OneDrive in a browser and copy the address
     up to and including `/personal/<account>/`. For a group/SharePoint site, use
     `https://<tenant>.sharepoint.com/sites/<SiteName>/`.
   - `FolderContains`: leave `/Apps/Microsoft Forms/` for OneDrive. For group forms,
     set it to part of that library's path.
4. When prompted for credentials, choose **Microsoft account** / **Organizational
   account** and sign in as the school account.
5. Name the query `Attendance`, then choose **Close & Load**.
6. Look at the `Student` column. If names come through as
   `Attendance.Sarah Cohen`, set `LikertTitle = "Attendance"` (whatever the Likert
   question is titled) and refresh.

### 5. Refresh

- **Data > Refresh All** pulls in all new responses from every class.
- To refresh automatically: **Data > Queries & Connections**, right-click
  `Attendance` > **Properties** > check **Refresh data when opening the file**
  (and, if you want, **Refresh every 60 minutes**).
- In Excel for the web, the table can be viewed and filtered. If your version
  can't refresh it there, refresh from desktop Excel.

## Reports

Add a PivotTable on the `Attendance` table:

- **Weekly sheet**: Rows = Class, Student; Columns = Date; Values = Status (use
  conditional formatting to color Absent/Late).
- **Absence counts**: Rows = Student; Filter Status = Absent; Values = Count of Student.
- **Missing submissions**: Rows = Date; Columns = Class; Values = Count of Student.
  A blank cell means that class didn't submit that day.

## How it behaves

| Situation | Result |
|---|---|
| New class/form added | Appears on next refresh if its workbook is in the folder and its name contains "Attendance". |
| Student added to a roster | New column in that form's workbook; appears automatically. |
| Student renamed on a roster | Old and new names show as different students. Filter or fix old rows in a pivot. |
| Teacher submits the same class and date twice | Only the latest submission is kept. |
| Teacher leaves a student blank | That student gets no row for that day. |
| Date field left blank | The row has an empty Date; filter for blanks to catch it. |

## Settings reference

| Setting | Default | Meaning |
|---|---|---|
| `SiteUrl` | – | OneDrive or SharePoint site holding the response workbooks |
| `FolderContains` | `/Apps/Microsoft Forms/` | Only read files under folders containing this text |
| `FileNameContains` | `Attendance` | Only read workbooks with this in the name; stripped from the class name |
| `ClassFrom` | `File` | Take the class name from the workbook name (`File`) or its folder (`Folder`) |
| `DateKeyword` | `date` | Word that identifies the date question |
| `TeacherKeyword` | `teacher` | Word that identifies the teacher free-text question |
| `LikertTitle` | `""` | Prefix to strip from student column headers, if Forms adds one |

## If you ever need real-time push instead

If the master sheet must update the moment a form is submitted without anyone
refreshing, you do need Power Automate. Keep it to one real flow:

1. One **child flow** (manual trigger with inputs Class, Date, Teacher, and
   Responses JSON) that does all the Excel writing.
2. For each form, a tiny flow: **When a new response is submitted > Get response
   details > Run a child flow**. Build it once, then use **Save As** and change the
   form for each class.

This still means one small flow per class. The Power Query approach above avoids
that and is easier to maintain.
