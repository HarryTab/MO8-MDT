# MO8 MDT Supabase Migration

This migration moves the MDT away from Google Apps Script and Google Sheets as the live backend.

The target setup is:

- GitHub Pages continues hosting the frontend.
- Supabase stores the live MDT data.
- Supabase Auth handles logins.
- Google Sheets becomes optional backup/export only.

## Why This Should Be Faster

Google Apps Script is slow for this MDT because every action goes through a script execution, reads/writes spreadsheet rows, then often reloads related data.

Supabase gives the MDT:

- a real Postgres database
- indexed queries
- direct browser-to-database requests through Supabase APIs
- built-in authentication
- row-level security policies
- less full-table reloading after small changes

## Free Plan Notes

Supabase currently has a Free plan suitable for small internal systems and prototypes. The exact limits can change, so check the official Supabase billing page before committing operationally.

For this MDT, the likely limits to watch are:

- database size
- monthly active users
- bandwidth/egress
- project inactivity/pausing rules

The MO8 roleplay use case should normally sit comfortably inside the free allowance, unless the system grows into a very large community or starts storing large files directly in Supabase.

## Step 1: Create Supabase Project

1. Go to `https://supabase.com`.
2. Create a free account or sign in.
3. Create a new project.
4. Name it something like `mo8-mdt`.
5. Save the database password somewhere safe.
6. Wait for the project to finish provisioning.

## Step 2: Create the Database Tables

1. Open the Supabase project.
2. Go to `SQL Editor`.
3. Open `supabase/schema.sql` from this repository.
4. Paste the whole file into Supabase.
5. Run it.

This creates the tables, indexes, starter permissions, training options, and first-pass row-level security rules.

## Step 3: Get Frontend Connection Details

In Supabase:

1. Go to `Project Settings`.
2. Open `API`.
3. Copy:
   - Project URL
   - anon public key

Only the anon public key goes in the frontend. Do not put the service role key in GitHub Pages.

## Step 4: Auth Decision

The old MDT uses username/password stored in Google Sheets.

Supabase Auth is safer and faster. The cleanest approach is:

- officers sign in with email/password
- each auth user links to one `profiles` row
- `profiles.roblox_username` remains the display identity

If you want to avoid real emails, use roleplay/admin email aliases such as:

- `harry_ted@mo8.local`
- `officername@mo8.local`

Supabase requires email-shaped values for email/password auth, but the officer-facing MDT can still display Roblox usernames.

## Step 5: Data Migration

Export each Google Sheet tab as CSV, then import into the matching Supabase table.

The key mappings are:

| Google Sheet | Supabase Table |
| --- | --- |
| Users | profiles |
| Officers | officers |
| TrainingRecords | training_records |
| TrainingMatrix | training_matrix |
| TrainingOptions | training_options |
| TrainingCourses | training_courses |
| CourseBookings | course_bookings |
| DisciplinaryActions | disciplinary_actions |
| LOARequests | loa_requests |
| TransferRequests | transfer_requests |
| SupervisorRequests | supervisor_requests |
| SupervisorCheckins | supervisor_checkins |
| DevelopmentPlans | development_plans |
| Appeals | appeals |
| Documents | documents |
| DocumentAcknowledgements | document_acknowledgements |
| DashboardWidgets | dashboard_widgets |
| ShiftLogs | shift_logs |
| Announcements | announcements |
| Permissions | permissions |
| UserPermissions | user_permissions |
| AuditLog | audit_log |
| RankChanges | rank_changes |
| Notifications | notifications |

Some columns are renamed to snake_case in Supabase. For example:

- `UserID` becomes `user_id`
- `OfficerID` becomes `officer_id`
- `RobloxUsername` becomes `roblox_username`
- `CreatedAt` becomes `created_at`

## Step 6: Frontend Migration Plan

The current frontend calls:

```js
api('actionName', payload)
```

The migration should be done in stages:

1. Add Supabase client configuration.
2. Replace login/logout first.
3. Replace read-only views next: documents, announcements, officers, courses.
4. Replace simple writes: notifications read, document acknowledgement, course request.
5. Replace admin writes: officer edits, LOA approvals, discipline, training, supervisor tools.
6. Remove Apps Script once every active feature is working on Supabase.

This avoids breaking the live MDT all at once.

## Step 7: Recommended First Frontend Cutover

Start with these views:

- login/session
- documents
- announcements
- notifications
- my profile

Then move the heavier admin workflows:

- officers
- LOA/tasks
- training/courses
- shifts
- supervisor
- permissions

## Important Security Notes

- Never commit the Supabase service role key.
- Only commit the project URL and anon public key.
- Keep row-level security enabled.
- Test with a Constable account and a Command account before making the Supabase version live.

