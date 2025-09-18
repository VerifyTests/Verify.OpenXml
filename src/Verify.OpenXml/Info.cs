class Info
{
    public required IReadOnlyList<string> SheetNames { get; init; }
    public required IReadOnlyList<CellFormat>? CellFormats { get; init; }
}