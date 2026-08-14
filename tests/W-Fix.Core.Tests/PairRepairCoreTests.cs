using System.Net;
using System.Text.Json.Nodes;
using WFix.Core.Abstractions;
using WFix.Core.Models;
using WFix.Core.Pairing;
using WFix.Core.Reporting;

namespace WFix.Core.Tests;

public sealed class PairInvitationValidatorTests
{
    [Fact]
    public void Validate_accepts_current_unexpired_invitation()
    {
        var now = DateTimeOffset.UtcNow;
        new PairInvitationValidator().Validate(CreateInvitation(now), now);
    }

    [Fact]
    public void Validate_rejects_expired_invitation()
    {
        var now = DateTimeOffset.UtcNow;
        var invitation = CreateInvitation(now - TimeSpan.FromMinutes(20));
        Assert.Throws<InvalidDataException>(() => new PairInvitationValidator().Validate(invitation, now));
    }

    [Fact]
    public void Validate_rejects_unknown_schema()
    {
        var now = DateTimeOffset.UtcNow;
        var invitation = CreateInvitation(now) with { SchemaVersion = 99 };
        Assert.Throws<InvalidDataException>(() => new PairInvitationValidator().Validate(invitation, now));
    }

    private static PairInvitation CreateInvitation(DateTimeOffset createdAt) => new()
    {
        SessionId = Guid.NewGuid().ToString("N"),
        HostComputerName = "PRINT-HOST",
        HostAddresses = ["192.0.2.10"],
        Port = 43123,
        CertificatePublicKeySha256 = new string('a', 64),
        ConfirmationCode = "123456",
        CreatedAt = createdAt,
        ExpiresAt = createdAt + TimeSpan.FromMinutes(15)
    };
}

public sealed class PairProtocolSerializerTests
{
    [Fact]
    public void Serializer_round_trips_allowlisted_message()
    {
        var serializer = new PairProtocolSerializer();
        var message = new PairHello(Guid.NewGuid().ToString("N"), "CLIENT-PC");
        var bytes = serializer.Serialize(PairMessageKind.Hello, message);
        Assert.Equal(message, serializer.Deserialize<PairHello>(PairMessageKind.Hello, bytes));
    }

    [Fact]
    public void Serializer_rejects_non_pair_action()
    {
        var serializer = new PairProtocolSerializer();
        var request = new PairActionRequest("r1", PairActionOperation.Execute, new PairRepairStep
        {
            Id = "s1",
            ActionId = "legacy:spooler",
            Endpoint = PairEndpointRole.Host,
            Title = "bad"
        });
        Assert.Throws<InvalidDataException>(() => serializer.Serialize(PairMessageKind.ActionRequest, request));
    }

    [Fact]
    public void Serializer_rejects_wrong_dto_for_message_kind()
    {
        var serializer = new PairProtocolSerializer();
        Assert.Throws<InvalidOperationException>(() => serializer.Serialize(PairMessageKind.Plan, new PairHello("id", "pc")));
    }
}

public sealed class PairFileServiceTests
{
    [Fact]
    public async Task Offline_snapshot_round_trip_verifies_signature()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".wfixpair");
        try
        {
            var service = new PairFileService(new PairInvitationValidator());
            var snapshot = new PairEndpointSnapshot
            {
                Endpoint = new PairEndpointDescriptor { Role = PairEndpointRole.Host, ComputerName = "HOST" },
                PrinterName = "Office Printer"
            };
            await service.WriteOfflineSnapshotAsync(path, snapshot);
            var restored = await service.ReadOfflineSnapshotAsync(path);
            Assert.Equal(snapshot.Endpoint.ComputerName, restored.Endpoint.ComputerName);
            Assert.Equal(snapshot.PrinterName, restored.PrinterName);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Offline_snapshot_rejects_tampering()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".wfixpair");
        try
        {
            var service = new PairFileService(new PairInvitationValidator());
            await service.WriteOfflineSnapshotAsync(path, new PairEndpointSnapshot
            {
                Endpoint = new PairEndpointDescriptor { Role = PairEndpointRole.Host, ComputerName = "HOST" }
            });
            var document = JsonNode.Parse(await File.ReadAllTextAsync(path))!.AsObject();
            var payload = document["snapshotPayloadBase64"]!.GetValue<string>();
            document["snapshotPayloadBase64"] = (payload[0] == 'A' ? "B" : "A") + payload[1..];
            await File.WriteAllTextAsync(path, document.ToJsonString());
            await Assert.ThrowsAsync<InvalidDataException>(() => service.ReadOfflineSnapshotAsync(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

public sealed class TlsPairSessionTransportTests
{
    [Fact]
    public async Task Loopback_session_requires_both_approvals_and_exchanges_typed_messages()
    {
        var transport = new TlsPairSessionTransport(new PairInvitationValidator());
        await using var host = await transport.StartHostAsync(new PairHostOptions
        {
            HostComputerName = "localhost",
            ListenAddresses = [IPAddress.Loopback],
            InvitationLifetime = TimeSpan.FromMinutes(2)
        });
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        var accept = host.AcceptAsync(timeout.Token);
        IPairSession client;
        try
        {
            client = await transport.JoinAsync(host.Invitation, timeout.Token);
        }
        catch (Exception clientError)
        {
            var serverError = await Assert.ThrowsAnyAsync<Exception>(async () => await accept);
            throw new AggregateException(clientError, serverError);
        }
        await using var clientScope = client;
        await using var server = await accept;

        var approvals = await Task.WhenAll(
            server.ApproveAsync(true, timeout.Token),
            clientScope.ApproveAsync(true, timeout.Token));
        Assert.All(approvals, Assert.True);
        Assert.Equal(PairSessionState.Approved, server.State);
        Assert.Equal(server.ConfirmationCode, clientScope.ConfirmationCode);

        var snapshot = Snapshot(PairEndpointRole.Client, "CLIENT");
        await clientScope.SendAsync(PairMessageKind.Snapshot, snapshot, timeout.Token);
        var received = await server.ReceiveAsync<PairEndpointSnapshot>(PairMessageKind.Snapshot, timeout.Token);
        Assert.Equal("CLIENT", received.Endpoint.ComputerName);
    }

    [Fact]
    public async Task Join_rejects_certificate_pin_mismatch()
    {
        var transport = new TlsPairSessionTransport(new PairInvitationValidator());
        await using var host = await transport.StartHostAsync(new PairHostOptions
        {
            HostComputerName = "localhost",
            ListenAddresses = [IPAddress.Loopback],
            InvitationLifetime = TimeSpan.FromMinutes(2)
        });
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        var accept = host.AcceptAsync(timeout.Token);
        var tampered = host.Invitation with { CertificatePublicKeySha256 = new string('0', 64) };
        await Assert.ThrowsAsync<IOException>(() => transport.JoinAsync(tampered, timeout.Token));
        await Assert.ThrowsAnyAsync<Exception>(async () => await accept);
    }

    private static PairEndpointSnapshot Snapshot(PairEndpointRole role, string name) => new()
    {
        Endpoint = new PairEndpointDescriptor { Role = role, ComputerName = name, IsLocalAgent = true }
    };
}

public sealed class PairDiagnosticsAndPlannerTests
{
    [Fact]
    public async Task Diagnostics_build_secure_host_and_client_repair_plan()
    {
        var host = Snapshot(PairEndpointRole.Host, "HOST") with
        {
            PrinterName = "USB Printer",
            PrinterShared = false,
            FileAndPrinterSharingFirewallEnabled = false,
            RpcEndpointMapperReachable = true,
            PeerNameResolved = true,
            NetworkDiscoveryFirewallEnabled = true,
            ServiceStates = RunningServices()
        };
        var client = Snapshot(PairEndpointRole.Client, "CLIENT") with
        {
            SmbPortReachable = false,
            RpcEndpointMapperReachable = false,
            PeerNameResolved = true,
            NetworkDiscoveryFirewallEnabled = true,
            ServiceStates = RunningServices()
        };
        var service = new PairDiagnosticService([
            new PairDiscoveryDiagnosticRule(), new PairSmbDiagnosticRule(), new PairRpcDiagnosticRule(),
            new PairPrinterShareDiagnosticRule(), new PairRpcCompatibilityDiagnosticRule()
        ]);
        var findings = await service.DiagnoseAsync(host, client);
        var plan = new PairRepairPlanner().CreatePlan(host, client, PairTransportMode.LiveLan, findings);

        Assert.Contains(plan.Steps, step => step.ActionId == "pair.firewall.file-print" && step.Endpoint == PairEndpointRole.Host);
        Assert.Contains(plan.Steps, step => step.ActionId == "pair.printer.share" && step.Endpoint == PairEndpointRole.Host);
        Assert.Contains(plan.Steps, step => step.ActionId == "pair.printer.connect" && step.Endpoint == PairEndpointRole.Client);
        var shareStep = Assert.Single(plan.Steps, step => step.ActionId == "pair.printer.share");
        var connectStep = Assert.Single(plan.Steps, step => step.ActionId == "pair.printer.connect");
        Assert.Equal(shareStep.Parameters["shareName"], connectStep.Parameters["shareName"]);
        Assert.DoesNotContain(plan.Steps, step => step.ExpertOnly);
        Assert.All(plan.Steps.Skip(1), step => Assert.Single(step.DependsOn));
    }

    [Fact]
    public void Planner_adds_named_pipes_only_in_expert_mode()
    {
        var host = Snapshot(PairEndpointRole.Host, "HOST");
        var client = Snapshot(PairEndpointRole.Client, "CLIENT");
        var finding = new PairDiagnosticFinding
        {
            RuleId = "pair.rpc.compatibility",
            Title = "RPC",
            Description = "RPC mismatch",
            AffectedEndpoints = [PairEndpointRole.Host, PairEndpointRole.Client],
            RecommendedActionIds = ["pair.rpc.named-pipes"],
            ExpertOnly = true
        };
        var planner = new PairRepairPlanner();
        Assert.Empty(planner.CreatePlan(host, client, PairTransportMode.LiveLan, [finding]).Steps);
        Assert.Equal(2, planner.CreatePlan(host, client, PairTransportMode.LiveLan, [finding], true).Steps.Count);
    }

    private static PairEndpointSnapshot Snapshot(PairEndpointRole role, string name) => new()
    {
        Endpoint = new PairEndpointDescriptor { Role = role, ComputerName = name, IsLocalAgent = true }
    };

    private static IReadOnlyDictionary<string, string> RunningServices() => new Dictionary<string, string>
    {
        ["FDResPub"] = "Running",
        ["fdPHost"] = "Running"
    };
}

public sealed class PairRepairExecutorTests
{
    [Fact]
    public async Task Verification_failure_rolls_back_both_endpoints_in_reverse_order()
    {
        var root = Path.Combine(Path.GetTempPath(), "wfix-pair-tests", Guid.NewGuid().ToString("N"));
        try
        {
            var dispatcher = new FakeDispatcher(failVerificationFor: "s2");
            var executor = new PairRepairExecutor(dispatcher, new PairRunReportService(root), root);
            var plan = Plan();
            var targets = new Dictionary<PairEndpointRole, TargetDescriptor>
            {
                [PairEndpointRole.Host] = TargetDescriptor.Local(),
                [PairEndpointRole.Client] = TargetDescriptor.Local()
            };
            var run = await executor.ExecuteAsync(plan, targets);
            Assert.Equal(PairRunStatus.RolledBack, run.Status);
            Assert.Equal(["s2", "s1"], dispatcher.RollbackOrder);
            Assert.Contains(run.Steps, step => step.StepId == "s1" && step.RolledBack);
            Assert.DoesNotContain("credential", await File.ReadAllTextAsync(Path.Combine(run.ReportDirectory!, "pair-run.json")), StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    private static PairRepairPlan Plan() => new()
    {
        Id = Guid.NewGuid().ToString("N"),
        Host = new PairEndpointDescriptor { Role = PairEndpointRole.Host, ComputerName = "HOST" },
        Client = new PairEndpointDescriptor { Role = PairEndpointRole.Client, ComputerName = "CLIENT" },
        TransportMode = PairTransportMode.LiveLan,
        Steps =
        [
            new PairRepairStep { Id = "s1", ActionId = "pair.test.one", Endpoint = PairEndpointRole.Host, Title = "one" },
            new PairRepairStep { Id = "s2", ActionId = "pair.test.two", Endpoint = PairEndpointRole.Client, Title = "two", DependsOn = ["s1"] }
        ]
    };

    private sealed class FakeDispatcher(string failVerificationFor) : IPairActionDispatcher
    {
        public List<string> RollbackOrder { get; } = [];

        public Task<PairActionCheckpoint> PrepareAsync(PairActionContext context, CancellationToken cancellationToken = default) =>
            Task.FromResult(new PairActionCheckpoint { ActionId = context.Step.ActionId, Endpoint = context.Step.Endpoint });

        public Task<PairActionResult> ExecuteAsync(PairActionContext context, CancellationToken cancellationToken = default) =>
            Task.FromResult(new PairActionResult { Success = true, Summary = "done" });

        public Task<bool> VerifyAsync(PairActionContext context, CancellationToken cancellationToken = default) =>
            Task.FromResult(context.Step.Id != failVerificationFor);

        public Task<PairActionResult> RollbackAsync(PairActionContext context, PairActionCheckpoint checkpoint, CancellationToken cancellationToken = default)
        {
            RollbackOrder.Add(context.Step.Id);
            return Task.FromResult(new PairActionResult { Success = true, Summary = "rolled back" });
        }
    }
}
