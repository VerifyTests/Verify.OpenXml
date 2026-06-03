namespace VerifyTests;

public static partial class VerifyOpenXml
{
    // Surfaces user-defined custom file properties (docProps/custom.xml) as a name → typed-value
    // map, mapping each entry to its variant type where possible. Shared by the Excel and Word
    // converters.
    internal static Dictionary<string, object?>? ReadCustomProperties(CustomFilePropertiesPart? part)
    {
        if (part?.Properties == null)
        {
            return null;
        }

        var properties = new Dictionary<string, object?>();
        foreach (var property in part.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
        {
            var name = property.Name?.Value;
            if (name == null)
            {
                continue;
            }

            object? value = property.InnerText;

            // Try to get typed value
            if (property.VTBool != null)
            {
                value = property.VTBool.Text == "true";
            }
            else if (property.VTInt32 != null)
            {
                value = int.Parse(property.VTInt32.Text);
            }
            else if (property.VTFloat != null)
            {
                value = float.Parse(property.VTFloat.Text);
            }
            else if (property.VTDouble != null)
            {
                value = double.Parse(property.VTDouble.Text);
            }
            else if (property.VTDate != null)
            {
                value = property.VTDate.Text;
            }
            else if (property.VTLPWSTR != null)
            {
                value = property.VTLPWSTR.Text;
            }

            properties[name] = value;
        }

        return properties.Count > 0 ? properties : null;
    }
}
