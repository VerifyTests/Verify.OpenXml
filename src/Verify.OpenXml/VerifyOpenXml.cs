namespace VerifyTests;

public static class VerifyOpenXml
{
    public static bool Initialized { get; private set; }

    public static void Initialize()
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;

        VerifierSettings.RegisterStreamConverter("xlsx", (_, target, settings) => Convert(target, settings));
        VerifierSettings.RegisterFileConverter<SpreadsheetDocument>(Convert);
    }

    static ConversionResult Convert(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        return Convert(document, settings);
    }

    static ConversionResult Convert(SpreadsheetDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var sheets = Convert(document).ToList();
        var info = new Info
        {
            SheetNames = sheets.Select(_ => _.Name!),
        };
        if (sheets.Count == 1)
        {
            var (csv, _) = sheets[0];
            return new(info, [new("csv", csv)]);
        }

        return new(
            info,
            sheets.Select(_ => new Target("csv", _.Csv, _.Name)));
    }

    static IEnumerable<(StringBuilder Csv, string? Name)> Convert(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        foreach (var sheet in workbookPart.Workbook.Sheets!.Elements<Sheet>())
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);

            // Get shared string table for text values
            var sharedStringPart = workbookPart.SharedStringTablePart;

            var builder = new StringBuilder();

            foreach (var row in worksheetPart.Worksheet.Descendants<Row>().OrderBy(r => r.RowIndex))
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellValue = GetCellValue(cell, sharedStringPart);
                    builder.Append(EscapeCsvValue(cellValue));
                    builder.Append(',');
                }

                builder.Length -= 1;
                builder.AppendLine();
            }

            yield return (builder, sheet.Name!.Value!);
        }
    }

    private static string GetCellValue(Cell cell, SharedStringTablePart? sharedStringPart)
    {
        var value = cell.InnerText;

        if (cell.DataType != null)
        {
            if (cell.DataType.Value == CellValues.SharedString)
            {
                // Handle shared strings
                if (sharedStringPart != null && int.TryParse(value, out var ssid))
                {
                    return sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ElementAt(ssid).InnerText;
                }
            }
            else if (cell.DataType.Value == CellValues.Boolean)
            {
                return value == "1" ? "TRUE" : "FALSE";
            }
            else if (cell.DataType.Value == CellValues.Date)
            {
                if (double.TryParse(value, out var oaDate))
                {
                    return DateTime.FromOADate(oaDate).ToString("yyyy-MM-dd");
                }
            }
        }

        return value;
    }

    static uint GetColumnIndex(string cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
            return 0;

        // Extract column letters from cell reference (e.g., "A1" -> "A")
        var columnName = new string(cellReference.Where(char.IsLetter).ToArray());
        return GetColumnIndex2(columnName);
    }

    private static uint GetColumnIndex2(string columnName)
    {
        uint columnIndex = 0;
        for (var i = 0; i < columnName.Length; i++)
        {
            columnIndex = columnIndex * 26 + (uint)(columnName[i] - 'A' + 1);
        }
        return columnIndex;
    }

    private static string GetColumnName(uint columnIndex)
    {
        var columnName = "";
        while (columnIndex > 0)
        {
            columnIndex--;
            columnName = (char)('A' + columnIndex % 26) + columnName;
            columnIndex /= 26;
        }
        return columnName;
    }

    private static string EscapeCsvValue(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "";

        // Escape CSV special characters
        if (value.Contains(',') ||
            value.Contains('"') ||
            value.Contains('\n') ||
            value.Contains('\r'))
        {
            // Escape quotes by doubling them
            value = value.Replace("\"", "\"\"");
            // Wrap in quotes
            value = "\"" + value + "\"";
        }

        return value;
    }
}