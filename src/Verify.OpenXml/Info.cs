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
    public WorkbookProtectionInfo? Protection { get; init; }
}

class SheetInfo
{
    public required string Name { get; init; }
    public IReadOnlyList<ColumnInfo>? Columns { get; init; }
    public SheetProtectionInfo? Protection { get; init; }
}

class ColumnInfo
{
    public required string Name { get; init; }
    public double? Width { get; init; }
    public bool ContainsRichText { get; init; }
    public string? NumberFormat { get; init; }
    public bool? Locked { get; init; }
    public bool RequiredHighlight { get; init; }
    public ColumnValidationInfo? Validation { get; init; }
}

class ColumnValidationInfo
{
    public string? Type { get; init; }
    public string? Operator { get; init; }
    public IReadOnlyList<string>? AllowedValues { get; init; }
    public string? Min { get; init; }
    public string? Max { get; init; }
    public bool AllowBlank { get; init; }
    public string? InputTitle { get; init; }
    public string? InputMessage { get; init; }
    public string? ErrorTitle { get; init; }
    public string? ErrorMessage { get; init; }
    public string? Range { get; init; }
}

class WorkbookProtectionInfo
{
    public string? Password { get; init; }
    public bool LockStructure { get; init; }
    public bool LockWindows { get; init; }
    public bool LockRevision { get; init; }
}

class SheetProtectionInfo
{
    public string? Password { get; init; }
    public bool Sheet { get; init; }
    public bool Objects { get; init; }
    public bool Scenarios { get; init; }
    public bool FormatCells { get; init; }
    public bool FormatColumns { get; init; }
    public bool FormatRows { get; init; }
    public bool InsertColumns { get; init; }
    public bool InsertRows { get; init; }
    public bool InsertHyperlinks { get; init; }
    public bool DeleteColumns { get; init; }
    public bool DeleteRows { get; init; }
    public bool SelectLockedCells { get; init; }
    public bool SelectUnlockedCells { get; init; }
    public bool Sort { get; init; }
    public bool AutoFilter { get; init; }
    public bool PivotTables { get; init; }
}
