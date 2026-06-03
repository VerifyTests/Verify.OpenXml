namespace VerifyTests;

public static partial class VerifyOpenXml
{
    // Reads core document properties (docProps/core.xml) for the dictionary-based Word and
    // PowerPoint infos. Creator, LastModifiedBy, Created and Modified are intentionally excluded:
    // DeterministicIoPackaging's CorePatcher strips them (they are user/time-specific), so capturing
    // them would make the info disagree with the deterministic document target that Verify re-reads
    // as a second snapshot.
    internal static Dictionary<string, object?>? GetCoreProperties(OpenXmlPackage document)
    {
        var packageProperties = document.PackageProperties;
        var properties = new Dictionary<string, object?>();

        AddPropertyIfNotEmpty(properties, "Title", packageProperties.Title);
        AddPropertyIfNotEmpty(properties, "Subject", packageProperties.Subject);
        AddPropertyIfNotEmpty(properties, "Keywords", packageProperties.Keywords);
        AddPropertyIfNotEmpty(properties, "Description", packageProperties.Description);
        AddPropertyIfNotEmpty(properties, "Category", packageProperties.Category);
        AddPropertyIfNotEmpty(properties, "ContentStatus", packageProperties.ContentStatus);
        AddPropertyIfNotEmpty(properties, "Revision", packageProperties.Revision);

        return properties.Count > 0 ? properties : null;
    }

    static void AddPropertyIfNotEmpty(Dictionary<string, object?> properties, string key, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            properties[key] = value;
        }
    }
}
