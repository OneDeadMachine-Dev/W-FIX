using System.Collections.Concurrent;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Pairing;

public sealed class PairSessionActionDispatcher(
    IPairSession session,
    PairEndpointRole remoteEndpoint,
    IPairActionDispatcher localDispatcher) : IPairActionDispatcher
{
    private readonly ConcurrentDictionary<string, string> _requestIds = new(StringComparer.Ordinal);

    public async Task<PairActionCheckpoint> PrepareAsync(PairActionContext context, CancellationToken cancellationToken = default)
    {
        if (context.Step.Endpoint != remoteEndpoint)
            return await localDispatcher.PrepareAsync(context, cancellationToken);
        var requestId = Guid.NewGuid().ToString("N");
        var response = await InvokeRemoteAsync(requestId, PairActionOperation.Prepare, context.Step, cancellationToken);
        if (!response.Result.Success)
            throw new InvalidOperationException(response.Result.Summary);
        _requestIds[context.Step.Id] = requestId;
        return new PairActionCheckpoint
        {
            ActionId = context.Step.ActionId,
            Endpoint = context.Step.Endpoint,
            State = new Dictionary<string, string?> { ["remoteRequestId"] = requestId }
        };
    }

    public async Task<PairActionResult> ExecuteAsync(PairActionContext context, CancellationToken cancellationToken = default)
    {
        if (context.Step.Endpoint != remoteEndpoint)
            return await localDispatcher.ExecuteAsync(context, cancellationToken);
        var response = await InvokeRemoteAsync(RequestId(context.Step.Id), PairActionOperation.Execute, context.Step, cancellationToken);
        return response.Result;
    }

    public async Task<bool> VerifyAsync(PairActionContext context, CancellationToken cancellationToken = default)
    {
        if (context.Step.Endpoint != remoteEndpoint)
            return await localDispatcher.VerifyAsync(context, cancellationToken);
        var response = await InvokeRemoteAsync(RequestId(context.Step.Id), PairActionOperation.Verify, context.Step, cancellationToken);
        return response.Result.Success && response.Result.Verified;
    }

    public async Task<PairActionResult> RollbackAsync(
        PairActionContext context,
        PairActionCheckpoint checkpoint,
        CancellationToken cancellationToken = default)
    {
        if (context.Step.Endpoint != remoteEndpoint)
            return await localDispatcher.RollbackAsync(context, checkpoint, cancellationToken);
        var requestId = checkpoint.State.GetValueOrDefault("remoteRequestId") ?? RequestId(context.Step.Id);
        var response = await InvokeRemoteAsync(requestId, PairActionOperation.Rollback, context.Step, cancellationToken);
        _requestIds.TryRemove(context.Step.Id, out _);
        return response.Result;
    }

    public async Task CompleteAsync(bool commit, CancellationToken cancellationToken = default)
    {
        if (!commit) return;
        var step = new PairRepairStep
        {
            Id = "pair-session-commit",
            ActionId = "pair.session.commit",
            Endpoint = remoteEndpoint,
            Title = "Commit PairRun"
        };
        var response = await InvokeRemoteAsync(Guid.NewGuid().ToString("N"), PairActionOperation.Commit, step, cancellationToken);
        if (!response.Result.Success) throw new InvalidOperationException(response.Result.Summary);
        _requestIds.Clear();
        await localDispatcher.CompleteAsync(true, cancellationToken);
    }

    private async Task<PairActionResponse> InvokeRemoteAsync(
        string requestId,
        PairActionOperation operation,
        PairRepairStep step,
        CancellationToken cancellationToken)
    {
        await session.SendAsync(PairMessageKind.ActionRequest, new PairActionRequest(requestId, operation, step), cancellationToken);
        var response = await session.ReceiveAsync<PairActionResponse>(PairMessageKind.ActionResult, cancellationToken);
        if (!string.Equals(response.RequestId, requestId, StringComparison.Ordinal))
            throw new InvalidDataException("Ответ pairing-агента не соответствует запросу.");
        return response;
    }

    private string RequestId(string stepId) =>
        _requestIds.TryGetValue(stepId, out var value)
            ? value
            : throw new InvalidOperationException($"Remote checkpoint для шага '{stepId}' не найден.");
}

public sealed class PairAgentCommandLoop(IPairActionDispatcher localDispatcher) : IPairAgentCommandLoop
{
    public async Task RunAsync(IPairSession session, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(session);
        if (session.LocalRole != PairEndpointRole.Host)
            throw new InvalidOperationException("Agent command loop должен выполняться на стороне Host.");
        var runDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "W-Fix", "Runs", "pair-" + session.Invitation.SessionId);
        Directory.CreateDirectory(runDirectory);
        var prepared = new Dictionary<string, PreparedRemoteAction>(StringComparer.Ordinal);
        var committed = false;
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                var request = await session.ReceiveAsync<PairActionRequest>(PairMessageKind.ActionRequest, cancellationToken);
                PairActionResult result;
                try
                {
                    if (request.Operation == PairActionOperation.Commit)
                    {
                        prepared.Clear();
                        committed = true;
                        result = new PairActionResult { Success = true, Verified = true, Summary = "PairRun committed." };
                        await session.SendAsync(PairMessageKind.ActionResult, new PairActionResponse(request.RequestId, result), cancellationToken);
                        return;
                    }
                    var context = new PairActionContext(TargetDescriptor.Local(), request.Step, runDirectory);
                    result = request.Operation switch
                    {
                        PairActionOperation.Prepare => await PrepareAsync(request, context, prepared, cancellationToken),
                        PairActionOperation.Execute => await ExecuteAsync(request, context, prepared, cancellationToken),
                        PairActionOperation.Verify => await VerifyAsync(request, context, prepared, cancellationToken),
                        PairActionOperation.Rollback => await RollbackAsync(request, context, prepared, cancellationToken),
                        _ => throw new InvalidOperationException("Неподдерживаемая операция pairing-агента.")
                    };
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    result = new PairActionResult { Success = false, Summary = ex.Message };
                }
                await session.SendAsync(PairMessageKind.ActionResult, new PairActionResponse(request.RequestId, result), cancellationToken);
            }
        }
        finally
        {
            if (!committed)
            {
                foreach (var item in prepared.Values.Reverse())
                {
                    try { await localDispatcher.RollbackAsync(item.Context, item.Checkpoint, CancellationToken.None); }
                    catch { /* Recovery journal and report preserve the remaining manual recovery requirement. */ }
                }
            }
        }
    }

    private async Task<PairActionResult> PrepareAsync(
        PairActionRequest request,
        PairActionContext context,
        Dictionary<string, PreparedRemoteAction> prepared,
        CancellationToken cancellationToken)
    {
        if (prepared.ContainsKey(request.RequestId))
            throw new InvalidOperationException("Pair action request уже подготовлен.");
        var checkpoint = await localDispatcher.PrepareAsync(context, cancellationToken);
        prepared.Add(request.RequestId, new PreparedRemoteAction(context, checkpoint));
        return new PairActionResult { Success = true, Summary = "Host snapshot created." };
    }

    private async Task<PairActionResult> ExecuteAsync(
        PairActionRequest request,
        PairActionContext context,
        IReadOnlyDictionary<string, PreparedRemoteAction> prepared,
        CancellationToken cancellationToken)
    {
        EnsurePrepared(request, context, prepared);
        return await localDispatcher.ExecuteAsync(context, cancellationToken);
    }

    private async Task<PairActionResult> VerifyAsync(
        PairActionRequest request,
        PairActionContext context,
        IReadOnlyDictionary<string, PreparedRemoteAction> prepared,
        CancellationToken cancellationToken)
    {
        EnsurePrepared(request, context, prepared);
        var verified = await localDispatcher.VerifyAsync(context, cancellationToken);
        return new PairActionResult { Success = verified, Verified = verified, Summary = verified ? "Host verification passed." : "Host verification failed." };
    }

    private async Task<PairActionResult> RollbackAsync(
        PairActionRequest request,
        PairActionContext context,
        Dictionary<string, PreparedRemoteAction> prepared,
        CancellationToken cancellationToken)
    {
        EnsurePrepared(request, context, prepared);
        var item = prepared[request.RequestId];
        var result = await localDispatcher.RollbackAsync(item.Context, item.Checkpoint, cancellationToken);
        if (result.Success) prepared.Remove(request.RequestId);
        return result;
    }

    private static void EnsurePrepared(
        PairActionRequest request,
        PairActionContext context,
        IReadOnlyDictionary<string, PreparedRemoteAction> prepared)
    {
        if (!prepared.TryGetValue(request.RequestId, out var item) ||
            !string.Equals(item.Context.Step.Id, context.Step.Id, StringComparison.Ordinal) ||
            !string.Equals(item.Context.Step.ActionId, context.Step.ActionId, StringComparison.Ordinal))
            throw new InvalidOperationException("Pair action request не имеет соответствующего checkpoint.");
    }

    private sealed record PreparedRemoteAction(PairActionContext Context, PairActionCheckpoint Checkpoint);
}
