class ExcelInfo
{
    public required IReadOnlyList<SheetInfo> Sheets { get; init; }
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

class SheetInfo
{
    public required string Name { get; init; }
    public IReadOnlyList<ColumnInfo>? Columns { get; init; }
}

class ColumnInfo
{
    public required string Name { get; init; }
    public double? Width { get; init; }
    public bool ContainsHtml { get; init; }
}