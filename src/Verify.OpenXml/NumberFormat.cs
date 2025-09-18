static class NumberFormatConverter
{
    static readonly Dictionary<uint, string> builtInFormats = new()
    {
        {
            0, "0: General"
        },
        {
            1, "1: 0"
        },
        {
            2, "2: 0.00"
        },
        {
            3, "3: #,##0"
        },
        {
            4, "4: #,##0.00"
        },
        {
            9, "9: 0%"
        },
        {
            10, "10: 0.00%"
        },
        {
            11, "11: 0.00E+00"
        },
        {
            12, "12: # ?/?"
        },
        {
            13, "13: # ??/??"
        },
        {
            14, "14: m/d/yyyy"
        },
        {
            15, "15: d-mmm-yy"
        },
        {
            16, "16: d-mmm"
        },
        {
            17, "17: mmm-yy"
        },
        {
            18, "18: h:mm AM/PM"
        },
        {
            19, "19: h:mm:ss AM/PM"
        },
        {
            20, "20: h:mm"
        },
        {
            21, "21: h:mm:ss"
        },
        {
            22, "22: m/d/yyyy h:mm"
        },
        {
            37, "37: #,##0 ;(#,##0)"
        },
        {
            38, "38: #,##0 ;[Red](#,##0)"
        },
        {
            39, "39: #,##0.00;(#,##0.00)"
        },
        {
            40, "40: #,##0.00;[Red](#,##0.00)"
        },
        {
            45, "45: mm:ss"
        },
        {
            46, "46: [h]:mm:ss"
        },
        {
            47, "47: mmss.0"
        },
        {
            48, "48: ##0.0E+0"
        },
        {
            49, "49: @"
        }
    };

    public static string GetReadableFormat(uint numberFormatId, WorkbookPart workbookPart)
    {
        if (builtInFormats.TryGetValue(numberFormatId, out var format))
        {
            return format;
        }

        var customFormat = GetCustomFormat(numberFormatId, workbookPart);
        if (!string.IsNullOrEmpty(customFormat))
        {
            return customFormat!;
        }

        return $"Unknown Format (ID: {numberFormatId})";
    }

    static string? GetCustomFormat(uint numberFormatId, WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart?.Stylesheet.NumberingFormats == null)
        {
            return null;
        }

        foreach (var numFmt in stylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>())
        {
            if (numFmt.NumberFormatId?.Value == numberFormatId)
            {
                return numFmt.FormatCode?.Value;
            }
        }

        return null;
    }

    public static string GetNumberFormat(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.StyleIndex == null)
        {
            return "General";
        }

        var stylesPart = workbookPart.WorkbookStylesPart;
        var cellFormats = stylesPart?.Stylesheet.CellFormats;
        var cellFormat = (CellFormat?) cellFormats?.ElementAt((int) cell.StyleIndex.Value);

        var numberFormatId = cellFormat?.NumberFormatId?.Value ?? 0;
        return GetReadableFormat(numberFormatId, workbookPart);
    }
}