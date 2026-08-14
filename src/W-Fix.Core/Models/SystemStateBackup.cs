namespace WFix.Core.Models;

/// <summary>Значение реестра, которое необходимо сохранить перед изменением.</summary>
public sealed record RegistryValueBackupTarget(string Path, string Name);

/// <summary>Раздел реестра, который может быть удалён целиком.</summary>
public sealed record RegistryKeyBackupTarget(string ProviderPath, string NativePath);

/// <summary>ACL раздела реестра, который может быть изменён фиксером.</summary>
public sealed record RegistryAclBackupTarget(string Path);

/// <summary>Декларативный список состояния, изменяемого фиксером.</summary>
public sealed record SystemStateBackupPlan
{
    public IReadOnlyList<RegistryValueBackupTarget> RegistryValues { get; init; } = [];
    public IReadOnlyList<RegistryKeyBackupTarget> RegistryKeys { get; init; } = [];
    public IReadOnlyList<RegistryAclBackupTarget> RegistryAcls { get; init; } = [];

    public bool IsEmpty => RegistryValues.Count == 0 && RegistryKeys.Count == 0 && RegistryAcls.Count == 0;
}

/// <summary>Снимок состояния, созданный на локальной или удалённой машине.</summary>
public sealed record SystemStateBackupResult
{
    public bool Success { get; init; }
    public bool Skipped { get; init; }
    public string? BackupDirectory { get; init; }
    public string? RemoteMachine { get; init; }
    public IReadOnlyList<string> Output { get; init; } = [];
    public string? Error { get; init; }
}
