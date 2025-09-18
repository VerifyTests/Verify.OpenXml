using Argon;
using DocumentFormat.OpenXml;

namespace VerifyTests;

public static class VerifyOpenXml
{
    [ThreadStatic]
    static SpreadsheetDocument? currentDocument;

    internal static List<JsonConverter> converters =
    [
        new CellFormatConverter(),
    ];

    public static bool Initialized { get; private set; }

    internal static SpreadsheetDocument? CurrentDocument => currentDocument;

    public static void Initialize()
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;

        VerifierSettings.RegisterStreamConverter("xlsx", (_, target, settings) => Convert(target, settings));
        VerifierSettings.RegisterFileConverter<SpreadsheetDocument>(Convert);
        VerifierSettings.AddExtraSettings(_ => _.Converters.AddRange(converters));
    }

    static ConversionResult Convert(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var document = SpreadsheetDocument.Open(stream, false, new()
        {
            AutoSave = false
        });
        return Convert(document, settings);
    }

    static ConversionResult Convert(SpreadsheetDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var sheets = Convert(document).ToList();

        var info = new Info
        {
            SheetNames = sheets.Select(_ => _.Name!).ToList(),
            CellFormats = document.WorkbookPart!.WorkbookStylesPart?.Stylesheet.CellFormats?.Elements<CellFormat>().ToList()
        };

        //new("xlsx", CloneToStream(document))
        List<Target> targets = [];
        if (sheets.Count == 1)
        {
            var (csv, _) = sheets[0];
            targets.Add(new("csv", csv));
        }
        else
        {
            targets.AddRange(sheets.Select(_ => new Target("csv", _.Csv, _.Name)));
        }

        return new(info, targets, () =>
        {
            if (!document.AutoSave)
            {
                document.Dispose();
            }
            return Task.CompletedTask;
        });
    }

    static IEnumerable<(StringBuilder Csv, string? Name)> Convert(SpreadsheetDocument document)
    {
        currentDocument = document;
        var workbookPart = document.WorkbookPart!;

        foreach (var sheet in workbookPart.Workbook.Sheets!.Elements<Sheet>())
        {
            var worksheetPart = (WorksheetPart) workbookPart.GetPartById(sheet.Id!);

            var sharedStringItems = workbookPart.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>().ToList();
            var builder = new StringBuilder();

            foreach (var row in worksheetPart.Worksheet
                         .Descendants<Row>()
                         .OrderBy(r => r.RowIndex))
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellValue = GetCellValue(cell, workbookPart, sharedStringItems);
                    builder.Append(EscapeCsvValue(cellValue));
                    builder.Append(',');
                }

                builder.Length -= 1;
                builder.AppendLine();
            }

            yield return (builder, sheet.Name!.Value!);
        }
    }

    static string GetCellValue(Cell cell, WorkbookPart workbookPart, List<SharedStringItem>? sharedStringItems)
    {
        var value = cell.InnerText;

        if (cell.DataType != null)
        {
            if (cell.DataType.Value == CellValues.SharedString)
            {
                // Handle shared strings
                if (sharedStringItems != null &&
                    int.TryParse(value, out var ssid))
                {
                    return sharedStringItems.ElementAt(ssid).InnerText;
                }
            }
            else if (cell.DataType.Value == CellValues.Boolean)
            {
                return value == "1" ? "true" : "false";
            }
            else if (cell.DataType.Value == CellValues.Date)
            {
                if (double.TryParse(value, out var oaDate))
                {
                    var date = DateTime.FromOADate(oaDate);
                    return DateFormatter.Convert(date);
                }
            }
        }
        else if (!string.IsNullOrEmpty(value))
        {
            // Check if this is a date based on number format
            if (double.TryParse(value, out var numericValue))
            {
                if (IsCellDateFormatted(cell, workbookPart))
                {
                    try
                    {
                        var date = DateTime.FromOADate(numericValue);
                        return DateFormatter.Convert(date);
                    }
                    catch (ArgumentException)
                    {
                        // If conversion fails, return the original numeric value
                        return value;
                    }
                }
            }
        }

        return value;
    }

    static bool IsCellDateFormatted(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.StyleIndex == null)
        {
            return false;
        }

        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart?.Stylesheet.CellFormats == null)
        {
            return false;
        }

        var cellFormats = stylesPart.Stylesheet.CellFormats;
        var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int) cell.StyleIndex.Value);

        if (cellFormat?.NumberFormatId == null)
        {
            return false;
        }

        var numberFormatId = cellFormat.NumberFormatId.Value;

        // Built-in date formats (14-22, 176-180, 181-183)
        if (numberFormatId is
            >= 14 and <= 22 or
            >= 176 and <= 180 or
            >= 181 and <= 183)
        {
            return true;
        }

        // Check custom number formats
        var numberingFormats = stylesPart.Stylesheet.NumberingFormats;
        if (numberingFormats != null)
        {
            var numberFormat = numberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(nf => nf.NumberFormatId != null && nf.NumberFormatId == numberFormatId);

            if (numberFormat?.FormatCode != null)
            {
                var formatCode = numberFormat.FormatCode.Value!.ToLower();
                // Look for common date format indicators
                return formatCode.Contains("yyyy") || formatCode.Contains("mm") ||
                       formatCode.Contains("dd") || formatCode.Contains('h') ||
                       formatCode.Contains("m/d") || formatCode.Contains("d/m");
            }
        }

        return false;
    }

    static string EscapeCsvValue(string value)
    {
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


    public static MemoryStream CloneToStream(SpreadsheetDocument sourceDocument)
    {
        var memoryStream = new MemoryStream();
        using (var targetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        {
            // Clone the workbook part and its content
            var sourceWorkbookPart = sourceDocument.WorkbookPart;
            var targetWorkbookPart = targetDocument.AddWorkbookPart();

            // Copy the workbook
            targetWorkbookPart.Workbook = new Workbook();
            targetWorkbookPart.Workbook.InnerXml = sourceWorkbookPart!.Workbook.InnerXml;

            // Copy styles if they exist
            if (sourceWorkbookPart.WorkbookStylesPart != null)
            {
                var targetStylesPart = targetWorkbookPart.AddNewPart<WorkbookStylesPart>();
                targetStylesPart.Stylesheet = new Stylesheet();
                targetStylesPart.Stylesheet.InnerXml = sourceWorkbookPart.WorkbookStylesPart.Stylesheet.InnerXml;
            }

            // Copy shared strings if they exist
            if (sourceWorkbookPart.SharedStringTablePart != null)
            {
                var targetSharedStringsPart = targetWorkbookPart.AddNewPart<SharedStringTablePart>();
                targetSharedStringsPart.SharedStringTable = new SharedStringTable();
                targetSharedStringsPart.SharedStringTable.InnerXml = sourceWorkbookPart.SharedStringTablePart.SharedStringTable.InnerXml;
            }

            // Copy worksheets
            foreach (var sourceWorksheetPart in sourceWorkbookPart.WorksheetParts)
            {
                var targetWorksheetPart = targetWorkbookPart.AddNewPart<WorksheetPart>();
                targetWorksheetPart.Worksheet = new Worksheet();
                targetWorksheetPart.Worksheet.InnerXml = sourceWorksheetPart.Worksheet.InnerXml;
            }

            targetDocument.Save();
        }
        return memoryStream;
    }
}