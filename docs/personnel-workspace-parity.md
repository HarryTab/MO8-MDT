# Personnel Desktop Workspace parity checklist

Desktop Workspace is the standard Personnel interface. It remains an interface layer over the existing Personnel Hub modules: a module is not considered migrated unless its permissions, actions, task routes, notifications, mobile behavior and Supabase operations remain available.

## Home and personal work

- Dashboard and configurable widgets
- Personal inbox, notifications and priority actions
- Chat, formal messages and attachments
- Global search, saved searches and record previews
- Personal profile, requests, activity and disciplinary visibility
- Tasks, filters, assignment, completion and generated workflows
- Calendar month/week/day views and targeted events
- Individual/system timelines and data-quality links

## Personnel management

- Shift logging, retrospective requests, debriefs, reviews and audited voiding
- Supervisor dashboard, assignment, check-ins, development plans and workload
- Officer roster, profiles, tags, supervisor links, bulk actions and account linking
- Probation, performance reviews, restrictions and AIPs
- Rank history and promotions
- Discipline create/edit/delete and officer visibility
- LOA request, approval, denial, appeal, edit and deletion
- Command handovers and targeted recipients

## Learning and knowledge

- Courses, bookings, trainers, approvals, waitlists and outcomes
- Training matrix, qualifications, driving levels and configurable standards
- Document folders, uploads, permissions, tags and acknowledgements

## Management and assurance

- Notice board and targeted announcements
- Recruitment vacancies, forms, applications and reviewer access
- Reports and exports
- Users, linked officer accounts, password and account requests
- Role and individual permission editing
- Audit log, permission audit and data-quality centre
- Settings, access simulation and browser-local design controls

## Cross-system requirements

- Personnel/CAD record links and supervisor CAD review tasks
- Discord notifications and decision-specific formatting
- Rank, role, tag and individual permission enforcement
- Loading, caching, update detection and stale-record protection
- Desktop and mobile workflows
- Classic Personnel Hub remains available temporarily as an emergency browser-local fallback

## Release gate

Desktop Workspace is now the default shell. Before the Classic fallback is removed, every item above must be exercised using representative Constable, Sergeant, Inspector and Command accounts. No existing action may be removed merely because it has not yet received its final Desktop Workspace presentation.
