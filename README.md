# Enterprise Employee Attendance Management System

Production-ready Google Apps Script attendance web app using Google Sheets as the database, manual Employee ID or employee name attendance capture, admin management, holidays, reporting, PDF print, and spreadsheet export.

## Files

- `Code.gs` - backend APIs, database setup, attendance rules, admin session handling, reports, cache, and locks.
- `Index.html` - clear employee mobile/manual attendance entry page.
- `Admin.html` - admin dashboard, employees, attendance editing, settings, and holidays.
- `Report.html` - monthly reports, filters, print-to-PDF, and spreadsheet export.
- `appsscript.json` - Apps Script manifest.

## Google Sheet Setup

1. Create a new Google Sheet.
2. Open `Extensions > Apps Script`.
3. Enable the manifest file:
   - Click the gear icon `Project Settings`.
   - Turn on `Show "appsscript.json" manifest file in editor`.
4. Add these project files with the exact names:
   - `Code.gs`
   - `Index.html`
   - `Admin.html`
   - `Report.html`
   - `appsscript.json`
5. Paste the contents from this folder into the matching Apps Script files.
6. In Apps Script, run `setupDatabase` once from the editor.
7. Approve the requested permissions.

If you do not see `appsscript.json`, it is only hidden. Apps Script projects support it, but the editor shows it only after that Project Settings option is enabled.

The script creates and formats these sheets:

- `Employees`
- `Master`
- `Config`

Default admin password is:

```text
admin123
```

Change it immediately in the `Config` sheet or from `Admin > Settings`.

## Required Sheet Structure

`Employees`

```text
Employee ID | Employee Name | Department | Phone | Email | Joining Date | Status
```

`Master`

```text
Date | Employee ID | Employee Name | Department | In Time | Out Time | Status | Late Minutes | Early Leave Minutes | OT Minutes | Working Minutes | Remarks | Edited By | Last Updated
```

`Config`

```text
Key | Value | Name/Notes
```

Default Config rows:

```text
Admin Password | admin123 |
Timezone | Asia/Karachi |
Work Start Time | 10:30 |
Work End Time | 20:00 |
Late Allowed Minutes | 5 |
Saturday Working | Yes |
Sunday Holiday | Yes |
OT Enabled | Yes |
Company Name | Attendance System |
```

Holidays are stored in `Config`, for example:

```text
HOLIDAY_01 | 2026-03-23 | Pakistan Day
HOLIDAY_02 | 2026-08-14 | Independence Day
```

## Deployment

1. In Apps Script, click `Deploy > New deployment`.
2. Select type `Web app`.
3. Set `Execute as` to `Me`.
4. Set `Who has access` based on your organization needs. For employee self-attendance, use `Anyone`.
5. Click `Deploy`.
6. Copy the Web App URL.

Routes:

- Employee app: the main Web App URL.
- Admin dashboard: add `?page=admin`.
- Reports: add `?page=report`.

## Employee Entry

Employees enter their Employee ID or their name, for example:

```text
EMP001
Yasir Abbas
```

If more than one employee has the same or similar name, the app asks for Employee ID so attendance is not marked for the wrong person.

## Attendance Rules

- First valid Employee ID/name entry of the day creates check-in.
- Second valid Employee ID/name entry of the day records check-out.
- Further entries for the same day are blocked.
- Inactive employees cannot mark attendance.
- Holidays and configured weekly holidays block attendance.
- Incomplete checkout remains `Present` with remark `Incomplete Checkout`.
- Late, early leave, overtime, and working time are stored in exact minutes.
- Duplicate rapid entries are blocked with `CacheService`.
- Race conditions are protected with `LockService`.

## Admin Features

- Password login from `Config`.
- Dashboard cards:
  - Total Employees
  - Active Employees
  - Today Present
  - Incomplete Records
  - Late Employees
  - Total OT Minutes
- Employee create/edit/search.
- Active/inactive employee status.
- Daily/monthly attendance filters.
- Partial search by employee name, ID, department, status, or remarks.
- Attendance record editing with automatic recalculation.
- Settings and holiday management.

## Reports

Reports support:

- Month filter
- Employee filter
- Department filter
- Status filter
- Partial search
- PDF export through browser print
- Excel export by creating a clean Google Spreadsheet

## Testing Checklist

1. Run `setupDatabase`.
2. Run `applyAttendanceSettings` once to force:
   - `Work Start Time = 10:30`
   - `Work End Time = 20:00`
   - `Late Allowed Minutes = 5`
3. If you already had wrong minute values, run `repairMasterAttendance` once.
4. Change the admin password.
5. Create one active employee, for example `EMP001`.
6. Open the employee app URL on a phone.
7. Enter `EMP001` or the employee name.
8. Confirm a row is created in `Master` with `In Time`.
9. Enter the same employee again after the duplicate cache window.
10. Confirm `Out Time`, `OT Minutes`, and `Working Minutes` are calculated.
11. Try a third entry and confirm it is blocked.
12. Mark the employee inactive and confirm attendance entry is blocked.
13. Add a holiday in Admin Settings and confirm attendance entry is blocked on that date.
14. Edit an attendance record and confirm minutes recalculate.
15. Open Reports, filter by month, print PDF, and export Excel.

## Notes

- The employee page is manual-entry only. It does not request camera permission.
- For 500+ employees, the code uses batch reads/writes, caching for employees/config, and single-row writes for attendance changes.
