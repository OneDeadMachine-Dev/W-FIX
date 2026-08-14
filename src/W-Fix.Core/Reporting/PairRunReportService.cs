using System.Text.Encodings.Web;
using System.Text.Json;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Reporting;

public sealed class PairRunReportService : IPairRunReportService
{
    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web) { WriteIndented = true };
    private readonly string _rootDirectory;

    public PairRunReportService(string? rootDirectory = null)
    {
        _rootDirectory = rootDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "W-Fix", "Runs");
    }

    public async Task<string> WriteAsync(PairRun run, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(run);
        if (!Guid.TryParseExact(run.Id, "N", out _))
            throw new InvalidDataException("Некорректный PairRun ID.");
        var directory = Path.Combine(_rootDirectory, "pair-" + run.Id);
        Directory.CreateDirectory(directory);
        var persisted = run with { ReportDirectory = directory };
        await using (var stream = new FileStream(Path.Combine(directory, "pair-run.json"), FileMode.Create, FileAccess.Write, FileShare.Read, 81920, true))
            await JsonSerializer.SerializeAsync(stream, persisted, Options, cancellationToken);
        await File.WriteAllTextAsync(Path.Combine(directory, "pair-report.html"), BuildHtml(persisted), cancellationToken);
        return directory;
    }

    internal static string BuildHtml(PairRun run)
    {
        var encoder = HtmlEncoder.Default;
        var steps = string.Join(Environment.NewLine, run.Steps.Select(step =>
            $"<tr><td>{encoder.Encode(step.Endpoint.ToString())}</td><td>{encoder.Encode(step.ActionId)}</td><td>{encoder.Encode(step.Summary)}</td><td>{(step.Verified ? "✓" : "✗")}</td><td>{(step.RolledBack ? "✓" : "—")}</td></tr>"));
        var findings = string.Join(Environment.NewLine, run.Findings.Select(finding =>
            $"<li><strong>{encoder.Encode(finding.Title)}</strong><br>{encoder.Encode(finding.Description)}</li>"));
        var warnings = string.Join(Environment.NewLine, run.Warnings.Select(warning => $"<li>{encoder.Encode(warning)}</li>"));
        return $$"""
            <!doctype html><html lang="ru"><head><meta charset="utf-8"><title>W-Fix Pair {{encoder.Encode(run.Id)}}</title>
            <style>body{font:14px Segoe UI,sans-serif;margin:32px;background:#0d1117;color:#e6edf3}section{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:18px;margin:14px 0}table{width:100%;border-collapse:collapse}th,td{padding:8px;border-bottom:1px solid #30363d;text-align:left}</style></head>
            <body><h1>W-Fix — отчёт парного ремонта</h1>
            <section><b>Host:</b> {{encoder.Encode(run.Host.ConnectionName)}}<br><b>Client:</b> {{encoder.Encode(run.Client.ConnectionName)}}<br><b>Transport:</b> {{run.TransportMode}}<br><b>Status:</b> {{run.Status}}</section>
            <section><h2>Диагностика</h2><ul>{{findings}}</ul></section>
            <section><h2>Действия</h2><table><thead><tr><th>ПК</th><th>Action</th><th>Результат</th><th>Проверено</th><th>Rollback</th></tr></thead><tbody>{{steps}}</tbody></table></section>
            <section><h2>Предупреждения</h2><ul>{{warnings}}</ul></section></body></html>
            """;
    }
}
