using JobKeeper.WinForms.Utilities;

namespace JobKeeper.Tests.Utilities;

/// <summary>
/// Unit tests for ValidationHelper
/// </summary>
public class ValidationHelperTests
{
    #region IsValidUrl Tests

    [Theory]
    [InlineData("https://www.google.com")]
    [InlineData("http://example.com")]
    [InlineData("https://test.com/path/to/page")]
    [InlineData("http://subdomain.example.com")]
    [InlineData("https://example.com:8080")]
    [InlineData("https://example.com/path?query=value")]
    public void IsValidUrl_WithValidUrls_ShouldReturnTrue(string url)
    {
        // Act
        var result = ValidationHelper.IsValidUrl(url);

        // Assert
        Assert.True(result);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("   ")]
    public void IsValidUrl_WithEmptyUrl_ShouldReturnTrue(string url)
    {
        // Act
        var result = ValidationHelper.IsValidUrl(url);

        // Assert
        Assert.True(result); // Empty URLs are allowed per implementation
    }

    [Theory]
    [InlineData("not a url")]
    [InlineData("ftp://example.com")]
    [InlineData("file://path/to/file")]
    [InlineData("javascript:alert('test')")]
    [InlineData("www.example.com")]
    [InlineData("example.com")]
    public void IsValidUrl_WithInvalidUrls_ShouldReturnFalse(string url)
    {
        // Act
        var result = ValidationHelper.IsValidUrl(url);

        // Assert
        Assert.False(result);
    }

    #endregion

    #region IsValidFilePath Tests

    [Fact]
    public void IsValidFilePath_WithExistingFile_ShouldReturnTrue()
    {
        // Arrange
        var tempFile = Path.GetTempFileName();

        try
        {
            // Act
            var result = ValidationHelper.IsValidFilePath(tempFile);

            // Assert
            Assert.True(result);
        }
        finally
        {
            // Cleanup
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    [Fact]
    public void IsValidFilePath_WithNonExistentFile_ShouldReturnFalse()
    {
        // Arrange
        var nonExistentPath = Path.Combine(Path.GetTempPath(), $"nonexistent_{Guid.NewGuid()}.txt");

        // Act
        var result = ValidationHelper.IsValidFilePath(nonExistentPath);

        // Assert
        Assert.False(result);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("   ")]
    public void IsValidFilePath_WithEmptyPath_ShouldReturnTrue(string path)
    {
        // Act
        var result = ValidationHelper.IsValidFilePath(path);

        // Assert
        Assert.True(result); // Empty paths are allowed per implementation
    }

    #endregion

    #region GetFileName Tests

    [Theory]
    [InlineData("C:\\Users\\test\\document.pdf", "document.pdf")]
    [InlineData("C:\\path\\to\\file.txt", "file.txt")]
    [InlineData("/usr/local/bin/script.sh", "script.sh")]
    [InlineData("document.pdf", "document.pdf")]
    public void GetFileName_WithValidPath_ShouldReturnFileName(string path, string expectedFileName)
    {
        // Act
        var result = ValidationHelper.GetFileName(path);

        // Assert
        Assert.Equal(expectedFileName, result);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("   ")]
    public void GetFileName_WithEmptyPath_ShouldReturnEmpty(string path)
    {
        // Act
        var result = ValidationHelper.GetFileName(path);

        // Assert
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void GetFileName_WithPathEndingInSlash_ShouldReturnEmpty()
    {
        // Arrange
        var path = "C:\\Users\\test\\";

        // Act
        var result = ValidationHelper.GetFileName(path);

        // Assert
        Assert.Equal(string.Empty, result);
    }

    [Theory]
    [InlineData("C:\\Users\\test\\my document.pdf", "my document.pdf")]
    [InlineData("C:\\path with spaces\\file name.txt", "file name.txt")]
    public void GetFileName_WithPathContainingSpaces_ShouldReturnFileName(string path, string expectedFileName)
    {
        // Act
        var result = ValidationHelper.GetFileName(path);

        // Assert
        Assert.Equal(expectedFileName, result);
    }

    #endregion
}
