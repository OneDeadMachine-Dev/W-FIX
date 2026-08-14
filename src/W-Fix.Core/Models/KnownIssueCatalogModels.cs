namespace WFix.Core.Models;

public sealed record KnownIssueCatalogDocument
{
    public int SchemaVersion { get; init; }
    public DateTimeOffset GeneratedAt { get; init; }
    public DateTimeOffset ExpiresAt { get; init; }
    public IReadOnlyList<KnownIssueEntry> Entries { get; init; } = [];
}

public sealed record KnownIssueEntry
{
    public required string Id { get; init; }
    public required string Title { get; init; }
    public required string Description { get; init; }
    public FindingSeverity Severity { get; init; } = FindingSeverity.Warning;
    public double Confidence { get; init; } = 0.8;
    public IReadOnlyList<string> AffectedOperatingSystems { get; init; } = [];
    public int? MinimumBuild { get; init; }
    public int? MaximumBuild { get; init; }
    public IReadOnlyList<string> RequiredKnowledgeBaseIds { get; init; } = [];
    public IReadOnlyList<string> PortKinds { get; init; } = [];
    public bool ThirdPartyDriverOnly { get; init; }
    public IReadOnlyList<string> RecommendedActionIds { get; init; } = [];
    public required Uri OfficialSource { get; init; }
}

public sealed record KnownIssueCatalogSnapshot
{
    public required string Source { get; init; }
    public DateTimeOffset LoadedAt { get; init; } = DateTimeOffset.UtcNow;
    public DateTimeOffset ExpiresAt { get; init; }
    public bool IsFallback { get; init; }
    public string? Warning { get; init; }
    public IReadOnlyList<KnownIssueEntry> Entries { get; init; } = [];
}
