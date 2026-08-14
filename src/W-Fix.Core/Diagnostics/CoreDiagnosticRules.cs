using WFix.Core.Abstractions;
using WFix.Core.Models;
using WFix.Core.Repair;

namespace WFix.Core.Diagnostics;

public sealed class SpoolerDiagnosticRule : IDiagnosticRule
{
    public string Id => "print.spooler.not-running";

    public Task<IReadOnlyList<DiagnosticFinding>> EvaluateAsync(
        PrinterInventory inventory,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (inventory.SpoolerStatus.Equals("Running", StringComparison.OrdinalIgnoreCase))
            return Task.FromResult<IReadOnlyList<DiagnosticFinding>>([]);

        return Task.FromResult<IReadOnlyList<DiagnosticFinding>>([new DiagnosticFinding
        {
            RuleId = Id,
            Target = inventory.Target,
            Title = "Диспетчер печати не работает",
            Description = "Служба Spooler остановлена или недоступна. Очереди не смогут обрабатывать задания.",
            Severity = FindingSeverity.Critical,
            Confidence = 1,
            Evidence = [$"Spooler: {inventory.SpoolerStatus}"],
            RecommendedActionIds = ["legacy:spooler"]
        }]);
    }
}

public sealed class PrinterStateDiagnosticRule(IRepairActionRegistry actions) : IDiagnosticRule
{
    private static readonly HashSet<PrinterStatus> BrokenStatuses =
    [
        PrinterStatus.Error,
        PrinterStatus.Stopped,
        PrinterStatus.Deleting,
        PrinterStatus.NotAvailable,
        PrinterStatus.UserIntervention
    ];

    public string Id => "print.queue.error-state";

    public Task<IReadOnlyList<DiagnosticFinding>> EvaluateAsync(
        PrinterInventory inventory,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var findings = new List<DiagnosticFinding>();
        var legacyActions = actions.GetAll().OfType<LegacyFixerRepairAction>().ToArray();

        foreach (var printer in inventory.Printers)
        {
            var matchedActions = legacyActions
                .Where(action => printer.ErrorCodes.Any(code => action.ErrorCodes.Any(targetCode =>
                    code.Contains(targetCode, StringComparison.OrdinalIgnoreCase))))
                .Select(action => action.Id)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (BrokenStatuses.Contains(printer.Status) && matchedActions.Count == 0)
                matchedActions.Add("legacy:spooler");

            if (matchedActions.Count == 0)
                continue;

            findings.Add(new DiagnosticFinding
            {
                RuleId = Id,
                Target = inventory.Target,
                PrinterName = printer.Name,
                Title = $"Очередь «{printer.Name}» сообщает об ошибке",
                Description = "Состояние очереди и обнаруженные коды указывают на неисправность подсистемы печати.",
                Severity = BrokenStatuses.Contains(printer.Status) ? FindingSeverity.Error : FindingSeverity.Warning,
                Confidence = printer.ErrorCodes.Count > 0 ? 0.95 : 0.75,
                Evidence = [$"Статус: {printer.Status}", $"Коды: {string.Join(", ", printer.ErrorCodes.DefaultIfEmpty("нет"))}"],
                RecommendedActionIds = matchedActions
            });
        }

        return Task.FromResult<IReadOnlyList<DiagnosticFinding>>(findings);
    }
}

public sealed class DefaultPrinterDiagnosticRule : IDiagnosticRule
{
    public string Id => "print.default.missing";

    public Task<IReadOnlyList<DiagnosticFinding>> EvaluateAsync(
        PrinterInventory inventory,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (inventory.Printers.Count == 0 || inventory.Printers.Any(printer => printer.IsDefault))
            return Task.FromResult<IReadOnlyList<DiagnosticFinding>>([]);

        var candidate = inventory.Printers.FirstOrDefault(printer => printer.Status == PrinterStatus.Ready)
                        ?? inventory.Printers[0];
        return Task.FromResult<IReadOnlyList<DiagnosticFinding>>([new DiagnosticFinding
        {
            RuleId = Id,
            Target = inventory.Target,
            PrinterName = candidate.Name,
            Title = "Не определён принтер по умолчанию",
            Description = "В интерактивном профиле пользователя нет выбранного принтера по умолчанию.",
            Severity = FindingSeverity.Warning,
            Confidence = 0.9,
            Evidence = [$"Доступно очередей: {inventory.Printers.Count}", "Default=True не найден"],
            RecommendedActionIds = ["legacy:defaultprinter"]
        }]);
    }
}

public sealed class ProtectedPrintModeDiagnosticRule : IDiagnosticRule
{
    public string Id => "windows11.protected-print.compatibility";

    public Task<IReadOnlyList<DiagnosticFinding>> EvaluateAsync(
        PrinterInventory inventory,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!inventory.WindowsProtectedPrintModeEnabled)
            return Task.FromResult<IReadOnlyList<DiagnosticFinding>>([]);

        var vendorDrivers = inventory.Drivers
            .Where(driver => !driver.Name.Contains("Microsoft IPP Class Driver", StringComparison.OrdinalIgnoreCase))
            .Select(driver => driver.Name)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (vendorDrivers.Length == 0)
            return Task.FromResult<IReadOnlyList<DiagnosticFinding>>([]);

        return Task.FromResult<IReadOnlyList<DiagnosticFinding>>([new DiagnosticFinding
        {
            RuleId = Id,
            Target = inventory.Target,
            Title = "Protected Print Mode и сторонние драйверы",
            Description = "Режим защищённой печати несовместим со сторонними драйверами. W-Fix не переключает его автоматически.",
            Severity = FindingSeverity.Warning,
            Confidence = 0.9,
            Evidence = vendorDrivers.Take(10).Select(name => $"Сторонний драйвер: {name}").ToArray(),
            OfficialSource = new Uri("https://learn.microsoft.com/en-us/windows/modern-print/windows-protected-print-mode/windows-protected-print-mode")
        }]);
    }
}

public sealed class UsbIppUpdateDiagnosticRule(IKnownIssueCatalog catalog) : IDiagnosticRule
{
    public string Id => "windows.known-print-issue";

    public async Task<IReadOnlyList<DiagnosticFinding>> EvaluateAsync(
        PrinterInventory inventory,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = await catalog.LoadAsync(cancellationToken: cancellationToken);
        var findings = new List<DiagnosticFinding>();
        foreach (var issue in snapshot.Entries)
        {
            if (!MatchesMachine(issue, inventory))
                continue;

            var printers = issue.PortKinds.Count == 0
                ? new PrinterSnapshot?[] { null }
                : inventory.Printers.Where(printer => MatchesPrinter(issue, printer)).Cast<PrinterSnapshot?>().ToArray();

            foreach (var printer in printers)
            {
                var evidence = new List<string>();
                evidence.AddRange(issue.RequiredKnowledgeBaseIds.Select(kb => $"Установлено {kb}"));
                if (printer is not null)
                {
                    evidence.Add($"Порт: {printer.PortKind}");
                    evidence.Add($"Драйвер: {printer.DriverName}");
                }
                if (!string.IsNullOrWhiteSpace(snapshot.Warning))
                    evidence.Add($"Каталог: {snapshot.Warning}");

                findings.Add(new DiagnosticFinding
                {
                    RuleId = Id,
                    Target = inventory.Target,
                    PrinterName = printer?.Name,
                    Title = issue.Title,
                    Description = issue.Description,
                    Severity = issue.Severity,
                    Confidence = issue.Confidence,
                    Evidence = evidence,
                    RecommendedActionIds = issue.RecommendedActionIds,
                    OfficialSource = issue.OfficialSource,
                    // Наличие KB/типа драйвера не исчезает после временного workaround.
                    VerifyResolution = false
                });
            }
        }
        return findings;
    }

    private static bool MatchesMachine(KnownIssueEntry issue, PrinterInventory inventory) =>
        (issue.AffectedOperatingSystems.Count == 0 || issue.AffectedOperatingSystems.Any(os =>
            inventory.OperatingSystem.Contains(os, StringComparison.OrdinalIgnoreCase))) &&
        (!issue.MinimumBuild.HasValue || inventory.BuildNumber >= issue.MinimumBuild) &&
        (!issue.MaximumBuild.HasValue || inventory.BuildNumber <= issue.MaximumBuild) &&
        (issue.RequiredKnowledgeBaseIds.Count == 0 || issue.RequiredKnowledgeBaseIds.Any(kb =>
            inventory.InstalledKnowledgeBaseIds.Contains(kb, StringComparer.OrdinalIgnoreCase)));

    private static bool MatchesPrinter(KnownIssueEntry issue, PrinterSnapshot printer) =>
        issue.PortKinds.Contains(printer.PortKind, StringComparer.OrdinalIgnoreCase) &&
        (!issue.ThirdPartyDriverOnly ||
         !printer.DriverName.Contains("Microsoft IPP Class Driver", StringComparison.OrdinalIgnoreCase));
}
