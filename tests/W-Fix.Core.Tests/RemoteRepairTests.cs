using WFix.Core.Abstractions;
using WFix.Core.Catalog;
using WFix.Core.Diagnostics;
using WFix.Core.Models;
using WFix.Core.Repair;
using WFix.Core.Reporting;
using WFix.Core.Services;
using System.Security.Cryptography;
using System.Net;

namespace WFix.Core.Tests;

public class TargetParserTests
{
    [Fact]
    public void ParseManual_NormalizesAndRemovesDuplicates()
    {
        var targets = TargetParser.ParseManual("PC-01, pc-01\nprint-02.contoso.local invalid/name");

        Assert.Equal(2, targets.Count);
        Assert.Contains(targets, target => target.ComputerName == "PC-01");
        Assert.Contains(targets, target => target.Fqdn == "print-02.contoso.local");
    }
}

public class RemoteCredentialContextTests
{
    [Fact]
    public void Scope_IsTargetBoundAndRestoresPreviousValue()
    {
        using (RemoteCredentialContext.Push("PC1", new NetworkCredential("DOMAIN\\one", "secret-one")))
        {
            Assert.True(RemoteCredentialContext.TryResolve("pc1", out var user, out var password));
            Assert.Equal("DOMAIN\\one", user);
            Assert.Equal("secret-one", password);
            Assert.False(RemoteCredentialContext.TryResolve("PC2", out _, out _));

            using (RemoteCredentialContext.Push("PC1", new NetworkCredential("DOMAIN\\two", "secret-two")))
            {
                Assert.True(RemoteCredentialContext.TryResolve("PC1", out user, out password));
                Assert.Equal("DOMAIN\\two", user);
                Assert.Equal("secret-two", password);
            }

            Assert.True(RemoteCredentialContext.TryResolve("PC1", out user, out password));
            Assert.Equal("DOMAIN\\one", user);
            Assert.Equal("secret-one", password);
        }

        Assert.False(RemoteCredentialContext.TryResolve("PC1", out _, out _));
    }
}

public class DiagnosticRuleTests
{
    [Fact]
    public async Task SpoolerRule_ProducesExecutableFindingWhenServiceStopped()
    {
        var inventory = Inventory("Stopped");
        var findings = await new SpoolerDiagnosticRule().EvaluateAsync(inventory);

        var finding = Assert.Single(findings);
        Assert.Equal(FindingSeverity.Critical, finding.Severity);
        Assert.Contains("legacy:spooler", finding.RecommendedActionIds);
    }

    [Fact]
    public async Task ProtectedPrintRule_IsDiagnosticOnly()
    {
        var inventory = Inventory("Running") with
        {
            WindowsProtectedPrintModeEnabled = true,
            Drivers = [new PrinterDriverSnapshot { Name = "Vendor Universal Driver" }]
        };

        var finding = Assert.Single(await new ProtectedPrintModeDiagnosticRule().EvaluateAsync(inventory));
        Assert.Empty(finding.RecommendedActionIds);
        Assert.NotNull(finding.OfficialSource);
    }

    [Fact]
    public async Task KnownIssueRule_MatchesKbPortAndThirdPartyDriver()
    {
        var target = TargetDescriptor.Local();
        var catalog = new FakeKnownIssueCatalog(new KnownIssueEntry
        {
            Id = "issue",
            Title = "Known issue",
            Description = "Known issue",
            AffectedOperatingSystems = ["Windows 11"],
            RequiredKnowledgeBaseIds = ["KB5079473"],
            PortKinds = ["USB"],
            ThirdPartyDriverOnly = true,
            RecommendedActionIds = ["legacy:ipp"],
            OfficialSource = new Uri("https://learn.microsoft.com/test")
        });
        var inventory = new PrinterInventory
        {
            Target = target,
            OperatingSystem = "Microsoft Windows 11 Enterprise",
            InstalledKnowledgeBaseIds = ["KB5079473"],
            Printers = [new PrinterSnapshot { Name = "MFP", PortKind = "USB", DriverName = "Vendor MFP" }]
        };

        var finding = Assert.Single(await new UsbIppUpdateDiagnosticRule(catalog).EvaluateAsync(inventory));

        Assert.Equal("MFP", finding.PrinterName);
        Assert.Contains("legacy:ipp", finding.RecommendedActionIds);
    }

    private static PrinterInventory Inventory(string spooler) => new()
    {
        Target = TargetDescriptor.Local(),
        SpoolerStatus = spooler
    };
}

public class RepairPlannerTests
{
    [Fact]
    public void CreatePlan_ExcludesActionWhenInteractiveUserCapabilityMissing()
    {
        var action = new FakeAction("user-action", RepairRequirement.InteractiveUser);
        var registry = new FakeRegistry(action);
        var target = TargetDescriptor.Local();
        var finding = new DiagnosticFinding
        {
            RuleId = "test",
            Target = target,
            Title = "test",
            Description = "test",
            RecommendedActionIds = [action.Id]
        };
        var capabilities = new RemoteCapabilityReport
        {
            Target = target,
            WinRmAvailable = true,
            IsAdministrator = true,
            SpoolerAvailable = true
        };

        var plan = new RepairPlanner(registry).CreatePlan(target, [finding], capabilities);

        Assert.Empty(plan.Steps);
    }
}

public class RepairExecutorTests
{
    [Fact]
    public async Task ExecuteBatch_ContainsFailureAndRollsBackOnlyFailedTarget()
    {
        var action = new FakeAction("test-action", RepairRequirement.None, failTarget: "PC2");
        var executor = new RepairExecutor(
            new FakeRegistry(action),
            new FakeInventoryService(),
            [],
            new FakeReportService());
        var plans = Enumerable.Range(1, 4).Select(index => Plan($"PC{index}", action.Id)).ToArray();

        var runs = await executor.ExecuteBatchAsync(plans, new RepairBatchOptions { MaxConcurrency = 2 });

        Assert.Equal(4, runs.Count);
        Assert.Equal(3, runs.Count(run => run.Status == RepairRunStatus.Succeeded));
        Assert.Equal(RepairRunStatus.RolledBack, runs.Single(run => run.Target.ComputerName == "PC2").Status);
        Assert.Equal(1, action.RollbackCount);
        Assert.InRange(action.MaximumConcurrency, 1, 2);
    }

    [Fact]
    public async Task ExecuteBatch_InventoryFailureDoesNotFaultOtherTargets()
    {
        var action = new FakeAction("test-action", RepairRequirement.None);
        var executor = new RepairExecutor(
            new FakeRegistry(action),
            new ConditionalInventoryService("PC2"),
            [],
            new FakeReportService());
        var plans = Enumerable.Range(1, 3).Select(index => Plan($"PC{index}", action.Id)).ToArray();

        var runs = await executor.ExecuteBatchAsync(plans, new RepairBatchOptions { MaxConcurrency = 2 });

        Assert.Equal(3, runs.Count);
        Assert.Equal(RepairRunStatus.Failed, runs.Single(run => run.Target.ComputerName == "PC2").Status);
        Assert.Equal(2, runs.Count(run => run.Status == RepairRunStatus.Succeeded));
    }

    private static RepairPlan Plan(string computer, string actionId)
    {
        var target = new TargetDescriptor { Id = computer, ComputerName = computer, Source = TargetSource.Manual };
        return new RepairPlan
        {
            Id = Guid.NewGuid().ToString("N"),
            Target = target,
            Steps = [new RepairStep { Id = "step-01", ActionId = actionId, Title = "test" }]
        };
    }
}

public class RunReportServiceTests
{
    [Fact]
    public void BuildHtml_EscapesRemoteValues()
    {
        var run = new RepairRun
        {
            Id = "run",
            Target = new TargetDescriptor { Id = "id", ComputerName = "<script>alert(1)</script>", Source = TargetSource.Manual },
            StartedAt = DateTimeOffset.UtcNow,
            CompletedAt = DateTimeOffset.UtcNow,
            Status = RepairRunStatus.Failed,
            Warnings = ["<b>warning</b>"]
        };

        var html = RunReportService.BuildHtml(run);

        Assert.DoesNotContain("<script>", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<b>warning</b>", html, StringComparison.OrdinalIgnoreCase);
    }
}

public class SupportBundleServiceTests
{
    [Fact]
    public async Task ExportAsync_RemovesComputerPrinterCredentialAndSnapshotPath()
    {
        var output = Path.Combine(Path.GetTempPath(), $"wfix-{Guid.NewGuid():N}.zip");
        var target = new TargetDescriptor
        {
            Id = "SECRET-PC",
            ComputerName = "SECRET-PC",
            Source = TargetSource.Manual,
            Credential = new CredentialReference("W-Fix/domain/admin", "DOMAIN\\admin")
        };
        var run = new RepairRun
        {
            Id = "run",
            Target = target,
            StartedAt = DateTimeOffset.UtcNow,
            CompletedAt = DateTimeOffset.UtcNow,
            Status = RepairRunStatus.Failed,
            Findings = [new DiagnosticFinding
            {
                RuleId = "rule",
                Target = target,
                PrinterName = "SECRET-PRINTER",
                Title = "SECRET-PRINTER on SECRET-PC",
                Description = "test"
            }],
            Steps = [new RepairStepResult
            {
                StepId = "step",
                ActionId = "action",
                SnapshotPath = @"C:\ProgramData\W-Fix\Backups\secret"
            }]
        };

        try
        {
            await new SupportBundleService().ExportAsync([run], output);
            using var archive = System.IO.Compression.ZipFile.OpenRead(output);
            using var reader = new StreamReader(archive.GetEntry("support-bundle.json")!.Open());
            var json = await reader.ReadToEndAsync();

            Assert.DoesNotContain("SECRET-PC", json, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("SECRET-PRINTER", json, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("DOMAIN", json, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Backups", json, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(output);
        }
    }
}

public class KnownIssueCatalogTests
{
    [Fact]
    public void EcdsaVerifier_AcceptsMatchingSignatureAndRejectsChangedCatalog()
    {
        using var key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var catalog = "catalog"u8.ToArray();
        var signature = key.SignData(catalog, HashAlgorithmName.SHA256);
        using var verifier = new EcdsaCatalogSignatureVerifier(key.ExportSubjectPublicKeyInfoPem());

        Assert.True(verifier.Verify(catalog, signature));
        Assert.False(verifier.Verify("changed"u8, signature));
    }

    [Fact]
    public async Task LoadAsync_WhenNetworkUnavailable_UsesValidatedEmbeddedCatalog()
    {
        var cache = Path.Combine(Path.GetTempPath(), "wfix-tests", Guid.NewGuid().ToString("N"));
        using var client = new HttpClient(new FailingHttpHandler());
        using var service = new KnownIssueCatalogService(client, cacheDirectory: cache);

        var snapshot = await service.LoadAsync();

        Assert.True(snapshot.IsFallback);
        Assert.Equal("embedded", snapshot.Source);
        Assert.Contains(snapshot.Entries, entry => entry.Id == "windows11.kb5079473.usb-ipp");
    }
}

internal sealed class FakeAction : IRepairAction
{
    private readonly string? _failTarget;
    private int _concurrency;
    private int _maximumConcurrency;
    private int _rollbackCount;

    public FakeAction(string id, RepairRequirement requirements, string? failTarget = null)
    {
        Id = id;
        Requirements = requirements;
        _failTarget = failTarget;
    }

    public string Id { get; }
    public string Name => Id;
    public string Description => Id;
    public RepairRisk Risk => RepairRisk.Reversible;
    public RepairRequirement Requirements { get; }
    public bool IsIdempotent => true;
    public int RollbackCount => _rollbackCount;
    public int MaximumConcurrency => _maximumConcurrency;

    public Task<RepairActionCheckpoint> PrepareAsync(RepairActionContext context, CancellationToken cancellationToken = default) =>
        Task.FromResult(new RepairActionCheckpoint("fake", new object()));

    public async Task<RepairActionResult> ExecuteAsync(RepairActionContext context, CancellationToken cancellationToken = default)
    {
        var current = Interlocked.Increment(ref _concurrency);
        int observed;
        do
        {
            observed = _maximumConcurrency;
            if (observed >= current) break;
        } while (Interlocked.CompareExchange(ref _maximumConcurrency, current, observed) != observed);

        try
        {
            await Task.Delay(50, cancellationToken);
            var success = !context.Target.ComputerName.Equals(_failTarget, StringComparison.OrdinalIgnoreCase);
            return new RepairActionResult(success, success ? "ok" : "failed", []);
        }
        finally
        {
            Interlocked.Decrement(ref _concurrency);
        }
    }

    public Task<bool> VerifyAsync(RepairActionContext context, RepairActionResult result, CancellationToken cancellationToken = default) =>
        Task.FromResult(result.Success);

    public Task<RepairActionResult> RollbackAsync(RepairActionContext context, RepairActionCheckpoint checkpoint, CancellationToken cancellationToken = default)
    {
        Interlocked.Increment(ref _rollbackCount);
        return Task.FromResult(new RepairActionResult(true, "rolled back", []));
    }
}

internal sealed class FakeRegistry(params IRepairAction[] actions) : IRepairActionRegistry
{
    public IReadOnlyList<IRepairAction> GetAll() => actions;
    public IRepairAction? Get(string actionId) => actions.FirstOrDefault(action => action.Id == actionId);
}

internal sealed class FakeInventoryService : IPrinterInventoryService
{
    public Task<PrinterInventory> CaptureAsync(TargetDescriptor target, CancellationToken cancellationToken = default) =>
        Task.FromResult(new PrinterInventory { Target = target, SpoolerStatus = "Running" });
}

internal sealed class ConditionalInventoryService(string failedComputer) : IPrinterInventoryService
{
    public Task<PrinterInventory> CaptureAsync(TargetDescriptor target, CancellationToken cancellationToken = default)
    {
        if (target.ComputerName.Equals(failedComputer, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("inventory failed");
        return Task.FromResult(new PrinterInventory { Target = target, SpoolerStatus = "Running" });
    }
}

internal sealed class FakeReportService : IRunReportService
{
    public Task<string> WriteAsync(RepairRun run, CancellationToken cancellationToken = default) =>
        Task.FromResult("fake-report");
}

internal sealed class FakeKnownIssueCatalog(params KnownIssueEntry[] entries) : IKnownIssueCatalog
{
    public Task<KnownIssueCatalogSnapshot> LoadAsync(bool forceRefresh = false, CancellationToken cancellationToken = default) =>
        Task.FromResult(new KnownIssueCatalogSnapshot
        {
            Source = "test",
            ExpiresAt = DateTimeOffset.UtcNow.AddDays(1),
            Entries = entries
        });
}

internal sealed class FailingHttpHandler : HttpMessageHandler
{
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
        throw new HttpRequestException("offline");
}
