namespace VerifyTests;

public static partial class VerifyOpenXml
{
    public static bool Initialized { get; private set; }

    public static void Initialize()
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;

        VerifierSettings.RegisterStreamConverter("xlsx", (_, target, settings) => ConvertExcel(target, settings));
        VerifierSettings.RegisterFileConverter<SpreadsheetDocument>(ConvertExcel);

        VerifierSettings.RegisterStreamConverter("docx", (_, target, settings) => ConvertWord(target, settings));
        VerifierSettings.RegisterFileConverter<WordprocessingDocument>(ConvertWord);
    }

    static ConversionResult ConvertExcel(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var document = SpreadsheetDocument.Open(stream, false, new()
        {
            AutoSave = false
        });
        return ConvertExcel(document, settings);
    }

    static ConversionResult ConvertExcel(SpreadsheetDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var sheets = Convert(document).ToList();
        var workbookPart = document.WorkbookPart!;

        // Extract document properties
        var packageProperties = document.PackageProperties;
        var workbookProperties = workbookPart.Workbook.WorkbookProperties;

        var info = new ExcelInfo
        {
            SheetNames = sheets.Select(_ => _.Name!).ToList(),
            WorksheetCount = sheets.Count,
            Title = packageProperties.Title,
            Subject = packageProperties.Subject,
            Creator = packageProperties.Creator,
            Keywords = packageProperties.Keywords,
            Description = packageProperties.Description,
            Category = packageProperties.Category,
            Date1904 = workbookProperties?.Date1904?.Value,
            CalculationMode = workbookPart.Workbook.CalculationProperties?.CalculationMode?.HasValue == true
                ? workbookPart.Workbook.CalculationProperties.CalculationMode.Value.ToString()
                : null
        };

        // Create deterministic XLSX output
        using var sourceStream = new MemoryStream();
        document.Clone(sourceStream);
        sourceStream.Position = 0;
        var resultStream = DeterministicPackage.Convert(sourceStream);

        List<Target> targets = [new("xlsx", resultStream)];
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
        var workbookPart = document.WorkbookPart!;
        var counter = Counter.Current;

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
                    var cellValue = GetCellValue(cell, workbookPart, sharedStringItems, counter);
                    builder.Append(EscapeCsvValue(cellValue));

                    // Add formula if present
                    if (cell.CellFormula != null && !string.IsNullOrEmpty(cell.CellFormula.Text))
                    {
                        builder.Append(" (");
                        builder.Append(EscapeCsvValue(cell.CellFormula.Text));
                        builder.Append(')');
                    }

                    builder.Append(',');
                }

                builder.Length -= 1;
                builder.AppendLine();
            }

            yield return (builder, sheet.Name!.Value!);
        }
    }

    static string GetCellValue(Cell cell, WorkbookPart workbookPart, List<SharedStringItem>? sharedStringItems, Counter counter)
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
                    var text = sharedStringItems.ElementAt(ssid).InnerText;
                    if (counter.TryConvert(text, out var result))
                    {
                        return result;
                    }
                    return text;
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
                    var dateString = DateFormatter.Convert(date);
                    if (counter.TryConvert(dateString, out var result))
                    {
                        return result;
                    }
                    return dateString;
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
                        var dateString = DateFormatter.Convert(date);
                        if (counter.TryConvert(dateString, out var result))
                        {
                            return result;
                        }
                        return dateString;
                    }
                    catch (ArgumentException)
                    {
                        // If conversion fails, return the original numeric value
                        return value;
                    }
                }
            }
        }

        // Try to scrub GUIDs and other text values
        if (counter.TryConvert(value, out var scrubbedValue))
        {
            return scrubbedValue;
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
            var numberFormat = numberingFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>()
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
}