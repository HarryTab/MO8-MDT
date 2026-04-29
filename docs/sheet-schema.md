# Google Sheet Schema

Run `setupSpreadsheet()` in Apps Script to create these tabs automatically.

## Users

| Column | Purpose |
| --- | --- |
| UserID | Internal unique ID |
| RobloxUsername | Login username |
| DiscordID | Optional Discord ID |
| Rank | Community rank |
| Role | Permission role |
| PasswordHash | Salted password hash |
| Salt | Per-user salt |
| Status | Active / Suspended / Archived |
| LastLogin | ISO timestamp |
| CreatedAt | ISO timestamp |
| CreatedBy | UserID or system |

## Sessions

| Column | Purpose |
| --- | --- |
| SessionID | Internal unique ID |
| UserID | Linked user |
| TokenHash | Hashed session token |
| CreatedAt | ISO timestamp |
| ExpiresAt | ISO timestamp |
| RevokedAt | ISO timestamp if logged out/revoked |
| UserAgent | Browser user agent if supplied |

## Officers

| Column | Purpose |
| --- | --- |
| OfficerID | Internal unique ID |
| RobloxUsername | Roblox username |
| DiscordID | Optional Discord ID |
| Callsign | Unit callsign |
| Rank | Officer rank |
| Status | Active / LOA / Suspended / Archived |
| JoinDate | Date joined MO8 |
| Notes | General notes |
| CreatedAt | ISO timestamp |
| UpdatedAt | ISO timestamp |

## TrainingRecords

| Column | Purpose |
| --- | --- |
| TrainingID | Internal unique ID |
| OfficerID | Linked officer |
| Standard | Training standard name |
| Status | Not Started / In Progress / Passed / Failed |
| Assessor | Roblox username or UserID |
| DateCompleted | Date |
| ExpiryDate | Optional date |
| Notes | Training notes |
| UpdatedAt | ISO timestamp |

## DisciplinaryActions

| Column | Purpose |
| --- | --- |
| ActionID | Internal unique ID |
| OfficerID | Linked officer |
| Type | Warning / Suspension / Removal / Note |
| Summary | Short summary |
| Details | Details |
| IssuedBy | UserID |
| IssuedAt | ISO timestamp |
| Status | Active / Expired / Appealed / Removed |

## LOARequests

| Column | Purpose |
| --- | --- |
| RequestID | Internal unique ID |
| OfficerID | Linked officer |
| StartDate | Date |
| EndDate | Date |
| Reason | Roleplay/community reason |
| Status | Pending / Approved / Denied / Cancelled |
| ReviewedBy | UserID |
| ReviewedAt | ISO timestamp |
| CreatedAt | ISO timestamp |

## Documents

| Column | Purpose |
| --- | --- |
| DocumentID | Internal unique ID |
| Title | Document title |
| Category | Training / Policy / SOP / Form |
| DriveURL | Google Drive or Docs URL |
| RequiredRole | Minimum role |
| Status | Published / Draft / Archived |
| UpdatedBy | UserID |
| UpdatedAt | ISO timestamp |

## Permissions

| Column | Purpose |
| --- | --- |
| Role | Role name |
| Permission | Permission key |
| Allowed | TRUE / FALSE |

## AuditLog

| Column | Purpose |
| --- | --- |
| AuditID | Internal unique ID |
| Timestamp | ISO timestamp |
| ActorUserID | User who performed action |
| Action | Action key |
| TargetType | Entity type |
| TargetID | Entity ID |
| Details | JSON details |
