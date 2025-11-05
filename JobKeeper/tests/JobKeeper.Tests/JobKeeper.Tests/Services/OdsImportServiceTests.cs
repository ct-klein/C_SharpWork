using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using JobKeeper.WinForms.Models;
using JobKeeper.WinForms.Services;

namespace JobKeeper.Tests.Services;

/// <summary>
/// Unit tests for OdsImportService
/// </summary>
public class OdsImportServiceTests : IDisposable
{
    private readonly OdsImportService _service;
    private readonly string _testFilePath;

    public OdsImportServiceTests()
    {
        _service = new OdsImportService();
        _testFilePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.ods");
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }

    [Fact]
    public void ImportFromOds_WithValidFile_ShouldReturnApplications()
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company", "Website", "Job Title", "Submitted", "Status" },
            new[] { "Test Corp", "https://test.com", "Developer", "2024-01-15", "Submitted" },
            new[] { "Example Inc", "https://example.com", "Engineer", "2024-01-20", "Rejected" }
        });

        // Act
        var applications = _service.ImportFromOds(_testFilePath);

        // Assert
        Assert.NotNull(applications);
        Assert.Equal(2, applications.Count);
        Assert.Equal("Test Corp", applications[0].Company);
        Assert.Equal("https://test.com", applications[0].Website);
        Assert.Equal("Developer", applications[0].JobTitle);
        Assert.Equal(ApplicationStatus.Submitted, applications[0].Status);
    }

    [Fact]
    public void ImportFromOds_WithNonExistentFile_ShouldThrowException()
    {
        // Arrange
        var nonExistentPath = Path.Combine(Path.GetTempPath(), $"nonexistent_{Guid.NewGuid()}.ods");

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => _service.ImportFromOds(nonExistentPath));
    }

    [Fact]
    public void ImportFromOds_WithInvalidOdsFile_ShouldThrowException()
    {
        // Arrange
        File.WriteAllText(_testFilePath, "This is not a valid ODS file");

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => _service.ImportFromOds(_testFilePath));
    }

    [Fact]
    public void ImportFromOds_WithEmptyRows_ShouldSkipEmptyRows()
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company", "Website", "Job Title" },
            new[] { "Test Corp", "https://test.com", "Developer" },
            new[] { "", "", "" },
            new[] { "Example Inc", "https://example.com", "Engineer" }
        });

        // Act
        var applications = _service.ImportFromOds(_testFilePath);

        // Assert
        Assert.Equal(2, applications.Count);
        Assert.Equal("Test Corp", applications[0].Company);
        Assert.Equal("Example Inc", applications[1].Company);
    }

    [Fact]
    public void ImportFromOds_WithMissingCompany_ShouldSkipRow()
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company", "Website", "Job Title" },
            new[] { "", "https://test.com", "Developer" },
            new[] { "Example Inc", "https://example.com", "Engineer" }
        });

        // Act
        var applications = _service.ImportFromOds(_testFilePath);

        // Assert
        Assert.Single(applications);
        Assert.Equal("Example Inc", applications[0].Company);
    }

    [Theory]
    [InlineData("Submitted", ApplicationStatus.Submitted)]
    [InlineData("REJECTED", ApplicationStatus.Rejected)]
    [InlineData("Interview Changed", ApplicationStatus.InterviewChanged)]
    [InlineData("GHOSTED", ApplicationStatus.Ghosted)]
    [InlineData("", ApplicationStatus.Submitted)]
    public void ImportFromOds_WithDifferentStatuses_ShouldParseCorrectly(string statusText, ApplicationStatus expectedStatus)
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company", "Status" },
            new[] { "Test Corp", statusText }
        });

        // Act
        var applications = _service.ImportFromOds(_testFilePath);

        // Assert
        Assert.Single(applications);
        Assert.Equal(expectedStatus, applications[0].Status);
    }

    [Theory]
    [InlineData("01/15/2024")]
    [InlineData("2024-01-15")]
    [InlineData("15/01/2024")]
    [InlineData("2024/01/15")]
    public void ImportFromOds_WithDifferentDateFormats_ShouldParseCorrectly(string dateText)
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company", "Submitted" },
            new[] { "Test Corp", dateText }
        });

        // Act
        var applications = _service.ImportFromOds(_testFilePath);

        // Assert
        Assert.Single(applications);
        Assert.NotNull(applications[0].Submitted);
    }

    [Fact]
    public void ImportFromOds_WithAllFields_ShouldPopulateAllProperties()
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company", "Website", "Job Title", "Submitted", "Resume", "Cover Letter", "Status", "Interview 1", "Interview 2" },
            new[] { "Test Corp", "https://test.com", "Developer", "2024-01-15", "resume.pdf", "cover.pdf", "Interview Changed", "2024-02-01", "2024-02-15" }
        });

        // Act
        var applications = _service.ImportFromOds(_testFilePath);

        // Assert
        Assert.Single(applications);
        var app = applications[0];
        Assert.Equal("Test Corp", app.Company);
        Assert.Equal("https://test.com", app.Website);
        Assert.Equal("Developer", app.JobTitle);
        Assert.NotNull(app.Submitted);
        Assert.Equal("resume.pdf", app.Resume);
        Assert.Equal("cover.pdf", app.CoverLetter);
        Assert.Equal(ApplicationStatus.InterviewChanged, app.Status);
        Assert.NotNull(app.Interview1);
        Assert.NotNull(app.Interview2);
    }

    [Fact]
    public void ImportFromOds_WithAlternativeHeaderNames_ShouldMapCorrectly()
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company Name", "Website URL", "Position", "Applied Date" },
            new[] { "Test Corp", "https://test.com", "Developer", "2024-01-15" }
        });

        // Act
        var applications = _service.ImportFromOds(_testFilePath);

        // Assert
        Assert.Single(applications);
        Assert.Equal("Test Corp", applications[0].Company);
        Assert.Equal("https://test.com", applications[0].Website);
        Assert.Equal("Developer", applications[0].JobTitle);
        Assert.NotNull(applications[0].Submitted);
    }

    [Fact]
    public void ImportFromOds_WithOnlyHeaderRow_ShouldThrowException()
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company", "Website", "Job Title" }
        });

        // Act & Assert
        var exception = Assert.Throws<InvalidOperationException>(() => _service.ImportFromOds(_testFilePath));
        Assert.Contains("at least a header row and one data row", exception.Message);
    }

    [Fact]
    public void ImportFromOds_WithInvalidDate_ShouldSetDateToNull()
    {
        // Arrange
        CreateTestOdsFile(_testFilePath, new[]
        {
            new[] { "Company", "Submitted" },
            new[] { "Test Corp", "invalid-date" }
        });

        // Act
        var applications = _service.ImportFromOds(_testFilePath);

        // Assert
        Assert.Single(applications);
        Assert.Null(applications[0].Submitted);
    }

    /// <summary>
    /// Helper method to create a test ODS file
    /// </summary>
    private void CreateTestOdsFile(string filePath, string[][] data)
    {
        // ODS namespaces
        XNamespace officeNs = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace tableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        XNamespace textNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

        // Create content.xml
        var content = new XDocument(
            new XElement(officeNs + "document-content",
                new XAttribute(XNamespace.Xmlns + "office", officeNs),
                new XAttribute(XNamespace.Xmlns + "table", tableNs),
                new XAttribute(XNamespace.Xmlns + "text", textNs),
                new XElement(officeNs + "body",
                    new XElement(officeNs + "spreadsheet",
                        new XElement(tableNs + "table",
                            new XAttribute(tableNs + "name", "Sheet1"),
                            data.Select(row =>
                                new XElement(tableNs + "table-row",
                                    row.Select(cell =>
                                        new XElement(tableNs + "table-cell",
                                            new XElement(textNs + "p", cell)
                                        )
                                    )
                                )
                            )
                        )
                    )
                )
            )
        );

        // Create ODS file (ZIP archive)
        using (var archive = ZipFile.Open(filePath, ZipArchiveMode.Create))
        {
            var entry = archive.CreateEntry("content.xml");
            using var stream = entry.Open();
            content.Save(stream);
        }
    }
}
