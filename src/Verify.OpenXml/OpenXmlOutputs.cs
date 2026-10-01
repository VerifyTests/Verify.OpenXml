namespace VerifyTests;

/// <summary>
/// The kinds of target a document is split into, chosen globally via <see cref="VerifyOpenXml.Initialize" />.
/// The info and the deterministic source package (docx/xlsx/pptx) are not controlled by this.
/// </summary>
[Flags]
public enum OpenXmlOutputs
{
    /// <summary>
    /// No outputs. Only the source document and info are emitted.
    /// </summary>
    None = 0,

    /// <summary>
    /// One png per rendered page. Only produced on <c>net10.0</c> with a Morph rendering backend referenced.
    /// </summary>
    Png = 1,

    /// <summary>
    /// The txt target holding the text of a Word document or Powerpoint presentation.
    /// </summary>
    Text = 2,

    /// <summary>
    /// One csv target per Excel worksheet.
    /// </summary>
    Csv = 4,

    /// <summary>
    /// All outputs. The default.
    /// </summary>
    All = Png | Text | Csv
}
