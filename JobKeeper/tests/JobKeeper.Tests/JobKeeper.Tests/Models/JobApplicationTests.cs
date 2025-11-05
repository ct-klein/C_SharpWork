using JobKeeper.WinForms.Models;

namespace JobKeeper.Tests.Models;

/// <summary>
/// Unit tests for JobApplication model
/// </summary>
public class JobApplicationTests
{
    [Fact]
    public void Constructor_ShouldGenerateNewGuid()
    {
        // Arrange & Act
        var application1 = new JobApplication();
        var application2 = new JobApplication();

        // Assert
        Assert.NotEqual(Guid.Empty, application1.Id);
        Assert.NotEqual(Guid.Empty, application2.Id);
        Assert.NotEqual(application1.Id, application2.Id);
    }

    [Fact]
    public void Constructor_ShouldSetDefaultStatusToSubmitted()
    {
        // Arrange & Act
        var application = new JobApplication();

        // Assert
        Assert.Equal(ApplicationStatus.Submitted, application.Status);
    }

    [Fact]
    public void Constructor_ShouldInitializeStringsToEmpty()
    {
        // Arrange & Act
        var application = new JobApplication();

        // Assert
        Assert.NotNull(application.Company);
        Assert.NotNull(application.Website);
        Assert.NotNull(application.JobTitle);
        Assert.NotNull(application.Resume);
        Assert.NotNull(application.CoverLetter);
        Assert.Equal(string.Empty, application.Company);
        Assert.Equal(string.Empty, application.Website);
        Assert.Equal(string.Empty, application.JobTitle);
    }

    [Fact]
    public void Properties_ShouldBeSettable()
    {
        // Arrange
        var application = new JobApplication();
        var testId = Guid.NewGuid();
        var testDate = DateTime.Now;

        // Act
        application.Id = testId;
        application.Company = "Test Company";
        application.Website = "https://test.com";
        application.JobTitle = "Software Engineer";
        application.Submitted = testDate;
        application.Resume = "C:\\resume.pdf";
        application.CoverLetter = "C:\\cover.pdf";
        application.Status = ApplicationStatus.Rejected;
        application.Interview1 = testDate.AddDays(7);
        application.Interview2 = testDate.AddDays(14);

        // Assert
        Assert.Equal(testId, application.Id);
        Assert.Equal("Test Company", application.Company);
        Assert.Equal("https://test.com", application.Website);
        Assert.Equal("Software Engineer", application.JobTitle);
        Assert.Equal(testDate, application.Submitted);
        Assert.Equal("C:\\resume.pdf", application.Resume);
        Assert.Equal("C:\\cover.pdf", application.CoverLetter);
        Assert.Equal(ApplicationStatus.Rejected, application.Status);
        Assert.Equal(testDate.AddDays(7), application.Interview1);
        Assert.Equal(testDate.AddDays(14), application.Interview2);
    }

    [Theory]
    [InlineData(ApplicationStatus.Submitted)]
    [InlineData(ApplicationStatus.Rejected)]
    [InlineData(ApplicationStatus.InterviewChanged)]
    [InlineData(ApplicationStatus.Ghosted)]
    public void Status_ShouldAcceptAllValidEnumValues(ApplicationStatus status)
    {
        // Arrange
        var application = new JobApplication();

        // Act
        application.Status = status;

        // Assert
        Assert.Equal(status, application.Status);
    }

    [Fact]
    public void DateProperties_ShouldAcceptNullValues()
    {
        // Arrange & Act
        var application = new JobApplication
        {
            Submitted = null,
            Interview1 = null,
            Interview2 = null
        };

        // Assert
        Assert.Null(application.Submitted);
        Assert.Null(application.Interview1);
        Assert.Null(application.Interview2);
    }
}
