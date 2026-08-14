using System.IO.Compression;
using System.Text.Json;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Reporting;

public sealed class SupportBundleService : ISupportBundleService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public async Task<string> ExportAsync(
        IReadOnlyList<RepairRun> runs,
        string outputPath,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(runs);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var fullPath = Path.GetFullPath(outputPath);
        var parent = Path.GetDirectoryName(fullPath) ?? throw new InvalidOperationException("Не удалось определить каталог экспорта.");
        Directory.CreateDirectory(parent);

        var aliases = runs.Select(run => run.Target.Id)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select((id, index) => (id, alias: $"PC-{index + 1:000}"))
            .ToDictionary(pair => pair.id, pair => pair.alias, StringComparer.OrdinalIgnoreCase);
        var sanitized = runs.Select(run => Sanitize(run, aliases[run.Target.Id])).ToArray();

        await using var archiveStream = new FileStream(fullPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 81920, useAsync: true);
        using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Create, leaveOpen: true);
        var jsonEntry = archive.CreateEntry("support-bundle.json", CompressionLevel.Optimal);
        await using (var json = jsonEntry.Open())
            await JsonSerializer.SerializeAsync(json, sanitized, JsonOptions, cancellationToken);

        var readmeEntry = archive.CreateEntry("README.txt", CompressionLevel.Fastest);
        await using (var writer = new StreamWriter(readmeEntry.Open()))
            await writer.WriteAsync("W-Fix support bundle. Имена компьютеров и credential references удалены; содержимое документов печати не экспортируется.");
        return fullPath;
    }

    private static RepairRun Sanitize(RepairRun run, string alias)
    {
        var target = new TargetDescriptor
        {
            Id = alias,
            ComputerName = alias,
            Source = run.Target.Source
        };
        var printerAliases = run.Findings.Select(finding => finding.PrinterName)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select((name, index) => (name: name!, alias: $"PRINTER-{index + 1:000}"))
            .ToDictionary(pair => pair.name, pair => pair.alias, StringComparer.OrdinalIgnoreCase);

        string Scrub(string value)
        {
            value = value.Replace(run.Target.ConnectionName, alias, StringComparison.OrdinalIgnoreCase)
                .Replace(run.Target.ComputerName, alias, StringComparison.OrdinalIgnoreCase);
            foreach (var pair in printerAliases)
                value = value.Replace(pair.Key, pair.Value, StringComparison.OrdinalIgnoreCase);
            return value;
        }

        return run with
        {
            Target = target,
            ReportDirectory = null,
            Findings = run.Findings.Select(finding => finding with
            {
                Target = target,
                PrinterName = finding.PrinterName is not null && printerAliases.TryGetValue(finding.PrinterName, out var printerAlias)
                    ? printerAlias
                    : null,
                Title = Scrub(finding.Title),
                Description = Scrub(finding.Description),
                Evidence = finding.Evidence.Select(Scrub).ToArray()
            }).ToArray(),
            Steps = run.Steps.Select(step => step with
            {
                SnapshotPath = null,
                Summary = Scrub(step.Summary),
                Output = step.Output.Select(Scrub).ToArray()
            }).ToArray(),
            Warnings = run.Warnings.Select(Scrub).ToArray()
        };
    }
}
