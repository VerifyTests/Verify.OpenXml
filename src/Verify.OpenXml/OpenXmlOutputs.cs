namespace VerifyTests;

/// <summary>
/// What <c>Initialize</c> used to be given to choose the outputs a document is split into. Settings
/// of Verify choose that now, and nothing reads this: it is here so that code still naming it is
/// told what to use in its place.
/// </summary>
[Obsolete(
    "OpenXmlOutputs and the outputs argument of Initialize are replaced by settings of Verify: VerifierSettings.ExcludeDerivedTargets(\"png\") to leave out the page images, VerifierSettings.PageText(PageTextPlacement.None) to leave out the text and VerifierSettings.ExcludeDerivedTargets(\"csv\") to leave out the csv files. See https://github.com/VerifyTests/Verify.OpenXml#migrating-from-1x",
    true)]
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
