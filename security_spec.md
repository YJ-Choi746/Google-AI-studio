# Security Specification & Test Payloads

## 1. Data Invariants
- Preorder applications must contain valid applicant identity info (`name`, `email`, `phone`, `ageOver14: true`, `consent: true`).
- An application cannot be read or listed by unauthenticated public users; only authenticated admins (matching `yeonj445@gmail.com` or configured in `/admins/{adminId}`) have read/list/update/delete permissions.
- Public users may create an application if and only if all mandatory fields are valid and volumetric bounds are respected.
- Document IDs must conform to `^[a-zA-Z0-9_\-]+$` with length <= 128.
- Audit logs can only be created by admins and are append-only.

## 2. Dirty Dozen Payloads (Targeting Rejection)
1. Public unauthorized read of applicant PII (`GET /applications/app123` with no auth). Expected: PERMISSION_DENIED.
2. Public listing of applicants (`LIST /applications` with no auth). Expected: PERMISSION_DENIED.
3. Applicant payload with oversized name (`name.length > 40`). Expected: PERMISSION_DENIED.
4. Applicant payload with empty email or invalid email string. Expected: PERMISSION_DENIED.
5. Applicant payload with underage user (`ageOver14: false`). Expected: PERMISSION_DENIED.
6. Applicant payload missing privacy consent (`consent: false`). Expected: PERMISSION_DENIED.
7. Unauthenticated update to an application (`UPDATE /applications/app123`). Expected: PERMISSION_DENIED.
8. Unauthenticated deletion of an application (`DELETE /applications/app123`). Expected: PERMISSION_DENIED.
9. Malicious path injection with illegal characters in document ID (`/applications/../../badId`). Expected: PERMISSION_DENIED.
10. Non-admin authenticated user attempting to read logs or applications. Expected: PERMISSION_DENIED.
11. Creation of an application with arbitrary unexpected keys (Shadow Update/Injection). Expected: PERMISSION_DENIED.
12. Attempt to tamper with immutable admin records by a standard user. Expected: PERMISSION_DENIED.
