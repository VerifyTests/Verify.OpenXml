using NumberingFormat = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat;

namespace VerifyTests;

public static partial class VerifyOpenXml
{
    static ConversionResult ConvertExcel(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = SpreadsheetDocument.Open(stream, false, new()
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
        var workbookProperties = workbookPart.Workbook?.WorkbookProperties;

        var sheetInfos = BuildSheetInfos(workbookPart);

        var info = new ExcelInfo
        {
            Sheets = sheetInfos,
            WorksheetCount = sheets.Count,
            Title = packageProperties.Title,
            Subject = packageProperties.Subject,
            Creator = packageProperties.Creator,
            Keywords = packageProperties.Keywords,
            Description = packageProperties.Description,
            Category = packageProperties.Category,
            Date1904 = workbookProperties?.Date1904?.Value,
            CalculationMode = workbookPart.Workbook?.CalculationProperties?.CalculationMode?.HasValue == true
                ? workbookPart.Workbook.CalculationProperties.CalculationMode.Value.ToString()
                : null,
            Protection = BuildWorkbookProtectionInfo(workbookPart)
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

        return new(info, targets);
    }

    internal static List<SheetInfo> BuildSheetInfos(WorkbookPart workbookPart)
    {
        var sheetInfos = new List<SheetInfo>();
        var propertyNames = ReadColumnPropertyNames(workbookPart);

        foreach (var sheetElement in workbookPart.Workbook!.Sheets!.Elements<Sheet>())
        {
            var worksheetPart = (WorksheetPart) workbookPart.GetPartById(sheetElement.Id!);
            var sheetName = sheetElement.Name!.Value!;
            var columns = GetColumnInfos(worksheetPart, workbookPart);
            if (columns != null && propertyNames.TryGetValue(sheetName, out var sheetProps))
            {
                ApplyPropertyNames(columns, sheetProps);
            }

            var sheetInfo = new SheetInfo
            {
                Name = sheetName,
                Columns = columns is { Count: > 0 } ? columns : null,
                Protection = BuildSheetProtectionInfo(worksheetPart)
            };
            sheetInfos.Add(sheetInfo);
        }

        return sheetInfos;
    }

    static void ApplyPropertyNames(List<ColumnInfo> columns, IReadOnlyDictionary<int, string> propertyNames)
    {
        for (var i = 0; i < columns.Count; i++)
        {
            if (propertyNames.TryGetValue(i + 1, out var name))
            {
                var existing = columns[i];
                columns[i] = new()
                {
                    Name = existing.Name,
                    PropertyName = name,
                    Width = existing.Width,
                    ContainsRichText = existing.ContainsRichText,
                    NumberFormat = existing.NumberFormat,
                    Locked = existing.Locked,
                    RequiredHighlight = existing.RequiredHighlight,
                    Validation = existing.Validation
                };
            }
        }
    }

    static Dictionary<string, IReadOnlyDictionary<int, string>> ReadColumnPropertyNames(WorkbookPart workbookPart)
    {
        var result = new Dictionary<string, IReadOnlyDictionary<int, string>>();
        const string ns = "http://schemas.simoncropp.com/excelsior/v1";

        foreach (var part in workbookPart.CustomXmlParts)
        {
            using var stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
            System.Xml.Linq.XDocument doc;
            try
            {
                doc = System.Xml.Linq.XDocument.Load(stream);
            }
            catch
            {
                continue;
            }

            var root = doc.Root;
            if (root == null || root.Name.NamespaceName != ns || root.Name.LocalName != "excelsior")
            {
                continue;
            }

            foreach (var sheetElement in root.Elements(System.Xml.Linq.XName.Get("sheet", ns)))
            {
                var sheetName = sheetElement.Attribute("name")?.Value;
                if (sheetName == null)
                {
                    continue;
                }

                var columnMap = new Dictionary<int, string>();
                foreach (var col in sheetElement.Elements(System.Xml.Linq.XName.Get("column", ns)))
                {
                    var indexAttr = col.Attribute("index")?.Value;
                    var prop = col.Attribute("property")?.Value;
                    if (indexAttr != null && prop != null && int.TryParse(indexAttr, out var index))
                    {
                        columnMap[index] = prop;
                    }
                }

                result[sheetName] = columnMap;
            }
        }

        return result;
    }

    static WorkbookProtectionInfo? BuildWorkbookProtectionInfo(WorkbookPart workbookPart)
    {
        var protection = workbookPart.Workbook?.GetFirstChild<WorkbookProtection>();
        if (protection == null)
        {
            return null;
        }

        return new()
        {
            Password = protection.WorkbookPassword?.Value,
            LockStructure = protection.LockStructure?.Value ?? false,
            LockWindows = protection.LockWindows?.Value ?? false,
            LockRevision = protection.LockRevision?.Value ?? false
        };
    }

    static SheetProtectionInfo? BuildSheetProtectionInfo(WorksheetPart worksheetPart)
    {
        var protection = worksheetPart.Worksheet?.GetFirstChild<SheetProtection>();
        if (protection == null)
        {
            return null;
        }

        return new()
        {
            Password = protection.Password?.Value,
            Sheet = protection.Sheet?.Value ?? false,
            Objects = protection.Objects?.Value ?? false,
            Scenarios = protection.Scenarios?.Value ?? false,
            FormatCells = protection.FormatCells?.Value ?? false,
            FormatColumns = protection.FormatColumns?.Value ?? false,
            FormatRows = protection.FormatRows?.Value ?? false,
            InsertColumns = protection.InsertColumns?.Value ?? false,
            InsertRows = protection.InsertRows?.Value ?? false,
            InsertHyperlinks = protection.InsertHyperlinks?.Value ?? false,
            DeleteColumns = protection.DeleteColumns?.Value ?? false,
            DeleteRows = protection.DeleteRows?.Value ?? false,
            SelectLockedCells = protection.SelectLockedCells?.Value ?? false,
            SelectUnlockedCells = protection.SelectUnlockedCells?.Value ?? false,
            Sort = protection.Sort?.Value ?? false,
            AutoFilter = protection.AutoFilter?.Value ?? false,
            PivotTables = protection.PivotTables?.Value ?? false
        };
    }

    internal static List<ColumnInfo>? GetColumnInfos(WorksheetPart worksheetPart, WorkbookPart workbookPart)
    {
        var sharedStringItems = workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToList();

        // Get the first row to extract column names
        var firstRow = worksheetPart.Worksheet!
            .Descendants<Row>()
            .OrderBy(_ => _.RowIndex)
            .FirstOrDefault();

        if (firstRow == null)
        {
            return null;
        }

        // Build a map of column index to custom width
        var columnWidths = new Dictionary<uint, double>();
        var columnLevelStyles = new Dictionary<uint, uint>();
        var columnsElement = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
        if (columnsElement != null)
        {
            foreach (var col in columnsElement.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>())
            {
                for (var i = col.Min!.Value; i <= col.Max!.Value; i++)
                {
                    if (col.CustomWidth?.Value == true &&
                        col.Width?.Value is { } width)
                    {
                        columnWidths[i] = width;
                    }

                    if (col.Style?.Value is { } styleIndex)
                    {
                        columnLevelStyles[i] = styleIndex;
                    }
                }
            }
        }

        var richTextColumns = FindRichTextColumns(worksheetPart, sharedStringItems, firstRow.RowIndex?.Value);
        var validationsByColumn = BuildValidationsByColumn(worksheetPart);
        var requiredColumns = FindRequiredHighlightColumns(worksheetPart);

        var result = new List<ColumnInfo>();
        uint colIndex = 1;

        foreach (var cell in firstRow.Elements<Cell>())
        {
            var name = GetHeaderCellValue(cell, sharedStringItems);
            double? width = columnWidths.TryGetValue(colIndex, out var w) ? Math.Round(w, 1) : null;

            string? numberFormat = null;
            bool? locked = null;
            if (columnLevelStyles.TryGetValue(colIndex, out var styleIndex))
            {
                (numberFormat, locked) = ReadColumnLevelStyle(workbookPart, styleIndex);
            }

            validationsByColumn.TryGetValue(colIndex, out var validation);

            result.Add(
                new()
                {
                    Name = name,
                    Width = width,
                    ContainsRichText = richTextColumns.Contains(colIndex),
                    NumberFormat = numberFormat,
                    Locked = locked,
                    RequiredHighlight = requiredColumns.Contains(colIndex),
                    Validation = validation
                });

            colIndex++;
        }

        return result;
    }

    static (string? NumberFormat, bool? Locked) ReadColumnLevelStyle(WorkbookPart workbookPart, uint styleIndex)
    {
        var stylesPart = workbookPart.WorkbookStylesPart;
        var cellFormats = stylesPart?.Stylesheet?.CellFormats;
        if (cellFormats == null)
        {
            return (null, null);
        }

        var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        if (cellFormat == null)
        {
            return (null, null);
        }

        string? numberFormat = null;
        if (cellFormat.NumberFormatId?.Value is { } nfId && nfId != 0)
        {
            numberFormat = ResolveNumberFormat(stylesPart!, nfId);
        }

        bool? locked = null;
        var protection = cellFormat.GetFirstChild<Protection>();
        if (protection?.Locked?.Value is { } lockedValue)
        {
            locked = lockedValue;
        }

        return (numberFormat, locked);
    }

    static string? ResolveNumberFormat(WorkbookStylesPart stylesPart, uint numberFormatId)
    {
        // Excelsior emits all formats as custom (id >= 164); built-ins are not stored.
        var numberingFormats = stylesPart.Stylesheet?.NumberingFormats;
        if (numberingFormats != null)
        {
            foreach (var format in numberingFormats.Elements<NumberingFormat>())
            {
                if (format.NumberFormatId?.Value == numberFormatId)
                {
                    return format.FormatCode?.Value;
                }
            }
        }

        return BuiltInNumberFormat(numberFormatId);
    }

    static string? BuiltInNumberFormat(uint id) =>
        id switch
        {
            14 => "m/d/yyyy",
            15 => "d-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "m/d/yyyy h:mm",
            _ => null
        };

    static Dictionary<uint, ColumnValidationInfo> BuildValidationsByColumn(WorksheetPart worksheetPart)
    {
        var result = new Dictionary<uint, ColumnValidationInfo>();
        var dataValidations = worksheetPart.Worksheet?.GetFirstChild<DataValidations>();
        if (dataValidations == null)
        {
            return result;
        }

        foreach (var dv in dataValidations.Elements<DataValidation>())
        {
            var sqref = dv.SequenceOfReferences?.InnerText;
            if (sqref == null)
            {
                continue;
            }

            var (firstColumn, lastColumn, range) = ParseSqref(sqref);
            if (firstColumn == null || lastColumn == null)
            {
                continue;
            }

            var info = BuildValidationInfo(dv, range);
            for (var col = firstColumn.Value; col <= lastColumn.Value; col++)
            {
                result[col] = info;
            }
        }

        return result;
    }

    static ColumnValidationInfo BuildValidationInfo(DataValidation dv, string? range)
    {
        var type = dv.Type?.InnerText;
        var op = dv.Operator?.InnerText;
        IReadOnlyList<string>? allowedValues = null;
        string? min = null;
        string? max = null;

        var f1 = dv.Formula1?.Text;
        var f2 = dv.Formula2?.Text;

        if (dv.Type?.Value == DataValidationValues.List && f1 != null)
        {
            allowedValues = ParseListFormula(f1);
        }
        else if (dv.Type?.Value == DataValidationValues.Date)
        {
            min = ParseDateFormula(f1);
            max = ParseDateFormula(f2);
        }
        else
        {
            min = f1;
            max = f2;
        }

        return new()
        {
            Type = type,
            Operator = op,
            AllowedValues = allowedValues,
            Min = min,
            Max = max,
            AllowBlank = dv.AllowBlank?.Value ?? false,
            ShowInputMessage = dv.ShowInputMessage?.Value ?? false,
            InputTitle = dv.PromptTitle?.Value,
            InputMessage = dv.Prompt?.Value,
            ShowErrorMessage = dv.ShowErrorMessage?.Value ?? false,
            ErrorStyle = dv.ErrorStyle?.InnerText,
            ErrorTitle = dv.ErrorTitle?.Value,
            ErrorMessage = dv.Error?.Value,
            Range = range
        };
    }

    static IReadOnlyList<string> ParseListFormula(string formula)
    {
        // Excel embeds list values as a single quoted, comma-separated string: "A,B,C"
        var trimmed = formula.Trim();
        if (trimmed is ['"', _, ..] &&
            trimmed[^1] == '"')
        {
            trimmed = trimmed.Substring(1, trimmed.Length - 2);
        }

        return trimmed.Split(',');
    }

    static string? ParseDateFormula(string? formula)
    {
        if (formula == null)
        {
            return null;
        }

        if (double.TryParse(formula, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var oa))
        {
            return DateTime.FromOADate(oa).ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
        }

        return formula;
    }

    static (uint? First, uint? Last, string? Range) ParseSqref(string sqref)
    {
        // sqref can contain multiple regions separated by spaces.
        // For attribution we use the union of column ranges and the row range from the first region.
        uint? first = null;
        uint? last = null;
        string? rowRange = null;

        foreach (var region in sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = region.Split(':');
            var (firstCol, firstRow) = ParseCellRef(parts[0]);
            var lastCol = firstCol;
            var lastRow = firstRow;
            if (parts.Length > 1)
            {
                (lastCol, lastRow) = ParseCellRef(parts[1]);
            }

            if (firstCol == null || lastCol == null)
            {
                continue;
            }

            first = first == null ? firstCol : Math.Min(first.Value, firstCol.Value);
            last = last == null ? lastCol : Math.Max(last.Value, lastCol.Value);
            if (rowRange == null && firstRow != null && lastRow != null)
            {
                rowRange = firstRow == lastRow ? firstRow.ToString() : $"{firstRow}-{lastRow}";
            }
        }

        return (first, last, rowRange);
    }

    static (uint? Column, uint? Row) ParseCellRef(string reference)
    {
        uint column = 0;
        var i = 0;
        while (i < reference.Length)
        {
            var c = reference[i];
            if (c is >= 'A' and <= 'Z')
            {
                column = column * 26 + (uint)(c - 'A' + 1);
                i++;
            }
            else if (c is >= 'a' and <= 'z')
            {
                column = column * 26 + (uint)(c - 'a' + 1);
                i++;
            }
            else
            {
                break;
            }
        }

        if (column == 0)
        {
            return (null, null);
        }

        if (i >= reference.Length)
        {
            return (column, null);
        }

        if (uint.TryParse(reference.AsSpan(i), out var row))
        {
            return (column, row);
        }

        return (column, null);
    }

    static HashSet<uint> FindRequiredHighlightColumns(WorksheetPart worksheetPart)
    {
        var result = new HashSet<uint>();
        if (worksheetPart.Worksheet == null)
        {
            return result;
        }

        foreach (var cf in worksheetPart.Worksheet.Elements<ConditionalFormatting>())
        {
            var hasBlankRule = cf.Elements<ConditionalFormattingRule>()
                .Any(_ => _.Type?.Value == ConditionalFormatValues.ContainsBlanks);
            if (!hasBlankRule)
            {
                continue;
            }

            var sqref = cf.SequenceOfReferences?.InnerText;
            if (sqref == null)
            {
                continue;
            }

            var (first, last, _) = ParseSqref(sqref);
            if (first == null || last == null)
            {
                continue;
            }

            for (var col = first.Value; col <= last.Value; col++)
            {
                result.Add(col);
            }
        }

        return result;
    }

    static HashSet<uint> FindRichTextColumns(WorksheetPart worksheetPart, List<SharedStringItem>? sharedStringItems, uint? headerRowIndex)
    {
        var richTextColumns = new HashSet<uint>();
        foreach (var row in worksheetPart.Worksheet!.Descendants<Row>())
        {
            if (row.RowIndex?.Value == headerRowIndex)
            {
                continue;
            }

            foreach (var cell in row.Elements<Cell>())
            {
                var colIndex = GetColumnIndex(cell);
                if (colIndex == null || richTextColumns.Contains(colIndex.Value))
                {
                    continue;
                }

                if (IsRichText(cell, sharedStringItems))
                {
                    richTextColumns.Add(colIndex.Value);
                }
            }
        }

        return richTextColumns;
    }

    static bool IsRichText(Cell cell, List<SharedStringItem>? sharedStringItems)
    {
        if (cell.DataType?.Value == CellValues.SharedString &&
            sharedStringItems != null &&
            int.TryParse(cell.InnerText, out var ssid) &&
            ssid >= 0 &&
            ssid < sharedStringItems.Count)
        {
            return sharedStringItems[ssid].Elements<DocumentFormat.OpenXml.Spreadsheet.Run>().Any();
        }

        if (cell.DataType?.Value == CellValues.InlineString &&
            cell.InlineString != null)
        {
            return cell.InlineString.Elements<DocumentFormat.OpenXml.Spreadsheet.Run>().Any();
        }

        return false;
    }

    static uint? GetColumnIndex(Cell cell)
    {
        var reference = cell.CellReference?.Value;
        if (reference == null)
        {
            return null;
        }

        uint index = 0;
        foreach (var c in reference)
        {
            if (c is >= 'A' and <= 'Z')
            {
                index = index * 26 + (uint) (c - 'A' + 1);
            }
            else if (c is >= 'a' and <= 'z')
            {
                index = index * 26 + (uint) (c - 'a' + 1);
            }
            else
            {
                break;
            }
        }

        return index == 0 ? null : index;
    }

    internal static string GetHeaderCellValue(Cell cell, List<SharedStringItem>? sharedStringItems)
    {
        var value = cell.InnerText;

        if (cell.DataType?.Value == CellValues.SharedString &&
            sharedStringItems != null &&
            int.TryParse(value, out var ssid))
        {
            return sharedStringItems.ElementAt(ssid).InnerText;
        }

        if (cell.DataType?.Value == CellValues.InlineString &&
            cell.InlineString != null)
        {
            return cell.InlineString.InnerText;
        }

        return value;
    }

    static IEnumerable<(StringBuilder Csv, string? Name)> Convert(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var counter = Counter.Current;

        foreach (var sheet in workbookPart.Workbook!.Sheets!.Elements<Sheet>())
        {
            var worksheetPart = (WorksheetPart) workbookPart.GetPartById(sheet.Id!);

            var sharedStringItems = workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToList();
            var builder = new StringBuilder();

            foreach (var row in worksheetPart.Worksheet!
                         .Descendants<Row>()
                         .OrderBy(_ => _.RowIndex))
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellValue = GetCellValue(cell, workbookPart, sharedStringItems, counter);
                    builder.Append(EscapeCsvValue(cellValue));

                    // Add formula if present
                    if (cell.CellFormula != null &&
                        !string.IsNullOrEmpty(cell.CellFormula.Text))
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

    internal static bool IsCellDateFormatted(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.StyleIndex == null)
        {
            return false;
        }

        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart?.Stylesheet?.CellFormats == null)
        {
            return false;
        }

        var cellFormats = stylesPart.Stylesheet.CellFormats;
        var cellFormat = cellFormats
            .Elements<CellFormat>()
            .ElementAtOrDefault((int) cell.StyleIndex.Value);

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
            NumberingFormat? numberFormat = null;
            foreach (var format in numberingFormats.Elements<NumberingFormat>())
            {
                if (format.NumberFormatId != null &&
                    format.NumberFormatId == numberFormatId)
                {
                    numberFormat = format;
                    break;
                }
            }

            if (numberFormat?.FormatCode != null)
            {
                var formatCode = numberFormat.FormatCode.Value!.ToLower();
                // Look for common date format indicators
                return formatCode.Contains("yyyy") ||
                       formatCode.Contains("mm") ||
                       formatCode.Contains("dd") ||
                       formatCode.Contains('h') ||
                       formatCode.Contains("m/d") ||
                       formatCode.Contains("d/m");
            }
        }

        return false;
    }

    internal static string EscapeCsvValue(string value)
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
