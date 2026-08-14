# W-Fix v3 architecture

W-Fix v3 separates presentation, remote transport, inventory, diagnosis, planning, execution, verification,
rollback, and reporting. `App.xaml.cs` is the WPF composition root; view models receive dependencies instead of
constructing WMI, Active Directory, backup, or repair services.

The Remote Center pipeline is:

1. Normalize manual or Active Directory computers into `TargetDescriptor` values.
2. Run capability-based preflight over WinRM. Ping is informational and never decides reachability.
3. Capture OS/build/KB, queues, jobs, ports, drivers, policies, Protected Print Mode, and PrintService events.
4. Produce evidence-backed `DiagnosticFinding` values and a capability-filtered `RepairPlan`.
5. Create targeted checkpoints, execute with bounded concurrency, verify through a fresh inventory, and roll back
   only the failed target.
6. Persist a JSON/HTML run report and optionally export an anonymized support bundle.

The twelve v2 fixers remain available through `LegacyFixerRepairAction`. This compatibility boundary preserves
working repair scripts while orchestration moves to `IRepairAction`, `IRepairPlanner`, and `IRepairExecutor`.

Alternate domain credentials are stored only in Windows Credential Manager. The password is materialized for a
short-lived `PSCredential`; it is never embedded in a PowerShell script, process arguments, configuration, or logs.

The known-issues catalog is declarative. A valid ECDSA signature does not grant code execution: catalog entries can
only reference action IDs compiled into W-Fix. Microsoft source URLs are restricted to `learn.microsoft.com` and
`support.microsoft.com`.
