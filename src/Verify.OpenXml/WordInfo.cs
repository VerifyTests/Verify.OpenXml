class WordInfo
{
    public Dictionary<string, object?>? Properties { get; init; }
    public Dictionary<string, object?>? CustomProperties { get; init; }
    public List<string>? Fonts { get; init; }
    public List<string>? EmbeddedFonts { get; init; }
    public string? Text { get; init; }
}
