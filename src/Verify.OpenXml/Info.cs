class ExcelInfo
{
    public required IReadOnlyList<string> SheetNames { get; init; }
    public required int WorksheetCount { get; init; }
    public string? Title { get; init; }
    public string? Subject { get; init; }
    public string? Creator { get; init; }
    public string? Keywords { get; init; }
    public string? Description { get; init; }
    public string? Category { get; init; }
    public bool? Date1904 { get; init; }
    public string? CalculationMode { get; init; }
}