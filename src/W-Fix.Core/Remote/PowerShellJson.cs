using System.Text.Json;
using WFix.Core.Abstractions;

namespace WFix.Core.Remote;

internal static class PowerShellJson
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public static T Deserialize<T>(RemoteCommandResult result)
    {
        if (!result.Success)
            throw new InvalidOperationException(result.Error ?? "Удалённая команда завершилась с ошибкой.");

        var json = result.Output.FirstOrDefault(line =>
            line.TrimStart().StartsWith('{') || line.TrimStart().StartsWith('['));
        if (string.IsNullOrWhiteSpace(json))
            throw new InvalidOperationException("Удалённая команда не вернула JSON.");

        return JsonSerializer.Deserialize<T>(json, Options)
               ?? throw new InvalidOperationException("Не удалось разобрать JSON удалённой команды.");
    }
}
