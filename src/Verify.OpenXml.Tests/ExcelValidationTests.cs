// Exercises the data-validation, conditional-formatting and column-level-style
// metadata captured per ColumnInfo.
[TestFixture]
public class ExcelValidationTests
{
    [Test]
    public Task ValidationsAndRequiredAndLocked()
    {
        using var document = CreateDocument();
        return Verify(document);
    }

    static SpreadsheetDocument CreateDocument()
    {
        var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = document.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());

        // Build stylesheet:
        //   format 0 = default
        //   format 1 = locked default (no number format)
        //   format 2 = locked + custom number format (id 164 -> "yyyy-mm-dd")
        //   format 3 = unlocked
        //   format 4 = locked + custom number format ($)
        var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new(
            new NumberingFormats(
                new NumberingFormat
                {
                    NumberFormatId = 164,
                    FormatCode = "yyyy-mm-dd"
                },
                new NumberingFormat
                {
                    NumberFormatId = 165,
                    FormatCode = "$#,##0.00"
                })
            {
                Count = 2
            },
            new CellFormats(
                new CellFormat(),
                new CellFormat
                {
                    ApplyProtection = true
                },
                new CellFormat
                {
                    NumberFormatId = 164,
                    ApplyNumberFormat = true
                },
                new CellFormat(
                    new Protection
                    {
                        Locked = false
                    })
                {
                    ApplyProtection = true
                },
                new CellFormat
                {
                    NumberFormatId = 165,
                    ApplyNumberFormat = true
                })
            {
                Count = 5
            });

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var cols = new Columns(
            new Column
            {
                Min = 1,
                Max = 1,
                Width = 25,
                CustomWidth = true
            },
            new Column
            {
                Min = 2,
                Max = 2,
                Width = 14,
                CustomWidth = true
            },
            new Column
            {
                Min = 3,
                Max = 3,
                Width = 15,
                CustomWidth = true,
                Style = 2
            },
            new Column
            {
                Min = 4,
                Max = 4,
                Width = 12,
                CustomWidth = true,
                Style = 3
            },
            new Column
            {
                Min = 5,
                Max = 5,
                Width = 18,
                CustomWidth = true,
                Style = 4
            });

        var sheetData = new SheetData(
            new Row(
                Header("Name"),
                Header("Status"),
                Header("HireDate"),
                Header("Score"),
                Header("Salary"))
            {
                RowIndex = 1u
            });

        var conditionalFormatting = new ConditionalFormatting(
            new ConditionalFormattingRule(
                new Formula("LEN(TRIM(A2))=0"))
            {
                Type = ConditionalFormatValues.ContainsBlanks,
                Priority = 1,
                FormatId = 0
            })
        {
            SequenceOfReferences = new()
            {
                InnerText = "A2:A100"
            }
        };

        var dataValidations = new DataValidations(
            new DataValidation(
                new Formula1("\"Active,Inactive,Contract\""))
            {
                Type = DataValidationValues.List,
                AllowBlank = false,
                ShowInputMessage = true,
                Prompt = "Pick one",
                ShowErrorMessage = true,
                ErrorTitle = "Invalid",
                Error = "Pick a status from the list.",
                SequenceOfReferences = new()
                {
                    InnerText = "B2:B100"
                }
            },
            new DataValidation(
                new Formula1("43831"),
                new Formula2("47848"))
            {
                Type = DataValidationValues.Date,
                Operator = DataValidationOperatorValues.Between,
                AllowBlank = true,
                SequenceOfReferences = new()
                {
                    InnerText = "C2:C100"
                }
            },
            new DataValidation(
                new Formula1("0"),
                new Formula2("100"))
            {
                Type = DataValidationValues.Decimal,
                Operator = DataValidationOperatorValues.Between,
                AllowBlank = true,
                SequenceOfReferences = new()
                {
                    InnerText = "D2:D100"
                }
            })
        {
            Count = 3
        };

        wsPart.Worksheet = new(cols, sheetData, conditionalFormatting, dataValidations);

        var sheets = wbPart.Workbook.GetFirstChild<Sheets>()!;
        sheets.Append(new Sheet
        {
            Id = wbPart.GetIdOfPart(wsPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        return document;

        static Cell Header(string value) =>
            new()
            {
                DataType = CellValues.InlineString,
                InlineString = new(new Text(value))
            };
    }
}
