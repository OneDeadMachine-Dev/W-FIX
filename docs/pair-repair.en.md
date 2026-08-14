# Pair Repair Wizard v3.1

Pair Repair fixes a shared printer between two explicitly selected PCs. It is neither a network scanner nor a persistent agent: W-Fix checks only the named peer and removes its temporary listener after the PairRun.

In a domain, Host and Client can be entered in one window. W-Fix uses the current Kerberos identity and existing WinRM transport. Queue connection and test-page operations run in the signed-in Client user's session through a temporary Task Scheduler job instead of the administrator service account's HKCU.

## Live mode

1. Run W-Fix as administrator on both PCs.
2. On the PC physically hosting the printer, open **Pair Repair**, choose **Host**, and enter the client name and exact local printer name.
3. Create and save a `.wfixpair` invitation. It expires after 15 minutes and is accepted once.
4. Select **Wait for client**. On the other PC choose **Client** and open the invitation.
5. Compare the six-digit code in both windows and approve on both PCs. Mutations remain disabled until dual approval succeeds.
6. The client gathers both snapshots and displays evidence plus the shared repair plan. Expert actions are excluded by default.
7. Execute the plan. Snapshots precede every mutation, verification determines success, and reversible steps roll back on both endpoints after a failure.
8. Send a Windows test page and confirm the physical output.

The live session uses TLS 1.2/1.3, a temporary ECDSA P-256 certificate, and public-key pinning. Its temporary Firewall rule is restricted to the current executable, exact port, Private/Domain profiles, and `LocalSubnet`. A listener is never opened on a Public profile.

## Workgroup credentials

When Windows requires authentication, the client can save an existing `HOST\User` credential scoped to that host. The password goes directly to Windows Credential Manager and is excluded from PairRun, invitations, support bundles, and W-Fix logs.

W-Fix never synchronizes primary-user passwords. SMB1 cannot be enabled. Insecure guest, disabled SMB signing, and reduced RPC protection remain expert-only actions with separate warnings and rollback.

## Offline fallback

If the live port is blocked, export a signed snapshot on one side and import it on the other. Each PC runs only its local plan. The file contains no credentials, but cross-machine automatic rollback is unavailable in offline mode.

## Collected diagnostics

- Windows version, domain/workgroup, and network profile;
- DNS for the selected peer, SMB 445, and RPC 135;
- Function Discovery, Spooler, and narrow Firewall rules;
- SMB signing, guest restrictions, and conflicting SMB sessions;
- shared queue, ShareName, driver, Point and Print, and RPC policies;
- PrintService/SMBClient errors without document names or content.

PairRun reports are stored under `%ProgramData%\W-Fix\Runs\pair-<RunId>`.
