namespace JobKeeper.WinForms.Utilities;

/// <summary>
/// Helper class for input validation
/// </summary>
public static class ValidationHelper
{
    /// <summary>
    /// Validates if a string is a valid URL
    /// </summary>
    public static bool IsValidUrl(string url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return true; // Empty URLs are allowed

        return Uri.TryCreate(url, UriKind.Absolute, out Uri? uriResult)
            && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
    }

    /// <summary>
    /// Validates if a file exists
    /// </summary>
    public static bool IsValidFilePath(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return true; // Empty paths are allowed

        return File.Exists(filePath);
    }

    /// <summary>
    /// Gets just the filename from a full path for display
    /// </summary>
    public static string GetFileName(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return string.Empty;

        return Path.GetFileName(filePath);
    }
}
