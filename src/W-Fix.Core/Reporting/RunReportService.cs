using System.Text.Encodings.Web;
using System.Text.Json;
using WFix.Core.Abstractions;
using WFix.Core.Models;

namespace WFix.Core.Reporting;

public sealed class RunReportService : IRunReportService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly string _rootDirectory;

    public RunReportService(string? rootDirectory = null)
    {
        _rootDirectory = rootDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "W-Fix",
            "Runs");
    }

    public async Task<string> WriteAsync(RepairRun run, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(run);
        var directory = Path.Combine(_rootDirectory, SanitizePathSegment(run.Id));
        Directory.CreateDirectory(directory);
        var publicTarget = run.Target with { Credential = null };
        var persisted = run with
        {
            Target = publicTarget,
            ReportDirectory = directory,
            Findings = run.Findings.Select(finding => finding with { Target = publicTarget }).ToArray()
        };

        await using (var json = File.Create(Path.Combine(directory, "run.json")))
            await JsonSerializer.SerializeAsync(json, persisted, JsonOptions, cancellationToken);

        await File.WriteAllTextAsync(
            Path.Combine(directory, "report.html"),
            BuildHtml(persisted),
            cancellationToken);
        return directory;
    }

    internal static string BuildHtml(RepairRun run)
    {
        var enc = HtmlEncoder.Default;
        var findings = string.Join(Environment.NewLine, run.Findings.Select(finding =>
            $"<li><strong>{enc.Encode(finding.Title)}</strong><br>{enc.Encode(finding.Description)}</li>"));
        var steps = string.Join(Environment.NewLine, run.Steps.Select(step =>
            $"<tr><td>{enc.Encode(step.ActionId)}</td><td>{enc.Encode(step.Summary)}</td><td>{(step.Verified ? "✓" : "✗")}</td><td>{(step.RolledBack ? "✓" : "—")}</td></tr>"));
        var warnings = string.Join(Environment.NewLine, run.Warnings.Select(warning => $"<li>{enc.Encode(warning)}</li>"));

        return $$"""
            <!doctype html>
            <html lang="ru"><head><meta charset="utf-8"><title>W-Fix {{enc.Encode(run.Id)}}</title>
            <style>body{font:14px Segoe UI,sans-serif;margin:32px;background:#0d1117;color:#e6edf3}section{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:18px;margin:14px 0}table{border-collapse:collapse;width:100%}th,td{border-bottom:1px solid #30363d;padding:8px;text-align:left}.ok{color:#3fb950}.warn{color:#d29922}</style></head>
            <body><h1>W-Fix — отчёт ремонта</h1>
            <section><b>Компьютер:</b> {{enc.Encode(run.Target.ConnectionName)}}<br><b>Статус:</b> {{enc.Encode(run.Status.ToString())}}<br><b>Начало:</b> {{run.StartedAt:u}}<br><b>Завершение:</b> {{run.CompletedAt:u}}</section>
            <section><h2>Диагностика</h2><ul>{{findings}}</ul></section>
            <section><h2>Выполненные шаги</h2><table><thead><tr><th>Действие</th><th>Результат</th><th>Проверено</th><th>Откат</th></tr></thead><tbody>{{steps}}</tbody></table></section>
            <section><h2>Предупреждения</h2><ul>{{warnings}}</ul></section>
            </body></html>
            """;
    }

    private static string SanitizePathSegment(string value)
    {
        foreach (var invalid in Path.GetInvalidFileNameChars())
            value = value.Replace(invalid, '_');
        return value;
    }
}
