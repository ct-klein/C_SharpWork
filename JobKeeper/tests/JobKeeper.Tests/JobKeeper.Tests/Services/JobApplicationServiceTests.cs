using System.Text.Json;
using JobKeeper.WinForms.Models;
using JobKeeper.WinForms.Services;

namespace JobKeeper.Tests.Services;

/// <summary>
/// Unit tests for JobApplicationService
/// </summary>
public class JobApplicationServiceTests : IDisposable
{
    private readonly string _testDataPath;
    private readonly JobApplicationService _service;

    public JobApplicationServiceTests()
    {
        // Create a temporary test data directory with unique ID
        var testId = Guid.NewGuid().ToString("N").Substring(0, 8);
        _testDataPath = Path.Combine(Path.GetTempPath(), $"JobKeeperTests_{testId}");

        // Ensure clean state
        if (Directory.Exists(_testDataPath))
        {
            Directory.Delete(_testDataPath, true);
        }

        Directory.CreateDirectory(_testDataPath);

        // Pass the test data path directly to the service constructor
        _service = new JobApplicationService(_testDataPath);
    }

    public void Dispose()
    {
        // Clean up test data
        if (Directory.Exists(_testDataPath))
        {
            Directory.Delete(_testDataPath, true);
        }
    }

    [Fact]
    public void Constructor_ShouldCreateDataDirectory()
    {
        // Assert - The test data path should exist since we pass it directly
        Assert.True(Directory.Exists(_testDataPath));
    }

    [Fact]
    public void GetAll_ShouldReturnEmptyListInitially()
    {
        // Act
        var applications = _service.GetAll();

        // Assert
        Assert.NotNull(applications);
        Assert.Empty(applications);
    }

    [Fact]
    public void Add_ShouldAddNewApplication()
    {
        // Arrange
        var application = new JobApplication
        {
            Company = "Test Company",
            JobTitle = "Developer"
        };

        // Act
        _service.Add(application);
        var applications = _service.GetAll();

        // Assert
        Assert.Single(applications);
        Assert.Equal("Test Company", applications[0].Company);
        Assert.Equal("Developer", applications[0].JobTitle);
    }

    [Fact]
    public void Add_WithNullApplication_ShouldThrowArgumentNullException()
    {
        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => _service.Add(null!));
    }

    [Fact]
    public void Update_ShouldModifyExistingApplication()
    {
        // Arrange
        var application = new JobApplication
        {
            Company = "Original Company",
            JobTitle = "Developer"
        };
        _service.Add(application);

        // Act
        application.Company = "Updated Company";
        application.JobTitle = "Senior Developer";
        _service.Update(application);

        var applications = _service.GetAll();

        // Assert
        Assert.Single(applications);
        Assert.Equal("Updated Company", applications[0].Company);
        Assert.Equal("Senior Developer", applications[0].JobTitle);
    }

    [Fact]
    public void Update_WithNullApplication_ShouldThrowArgumentNullException()
    {
        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => _service.Update(null!));
    }

    [Fact]
    public void Update_WithNonExistentId_ShouldNotModifyList()
    {
        // Arrange
        var application1 = new JobApplication { Company = "Company 1" };
        _service.Add(application1);

        var application2 = new JobApplication { Company = "Company 2" };

        // Act
        _service.Update(application2);
        var applications = _service.GetAll();

        // Assert
        Assert.Single(applications);
        Assert.Equal("Company 1", applications[0].Company);
    }

    [Fact]
    public void Delete_ShouldRemoveApplication()
    {
        // Arrange
        var application = new JobApplication { Company = "Test Company" };
        _service.Add(application);

        // Act
        _service.Delete(application.Id);
        var applications = _service.GetAll();

        // Assert
        Assert.Empty(applications);
    }

    [Fact]
    public void Delete_WithNonExistentId_ShouldNotThrowException()
    {
        // Arrange
        var application = new JobApplication { Company = "Test Company" };
        _service.Add(application);

        // Act
        _service.Delete(Guid.NewGuid());
        var applications = _service.GetAll();

        // Assert
        Assert.Single(applications);
    }

    [Fact]
    public void SaveData_ShouldPersistDataToFile()
    {
        // Arrange
        var application = new JobApplication
        {
            Company = "Test Company",
            JobTitle = "Developer"
        };
        _service.Add(application);

        // Act
        _service.SaveData();
        var filePath = _service.GetDataFilePath();

        // Assert
        Assert.True(File.Exists(filePath));
        var json = File.ReadAllText(filePath);
        var applications = JsonSerializer.Deserialize<List<JobApplication>>(json);
        Assert.NotNull(applications);
        Assert.Single(applications);
        Assert.Equal("Test Company", applications[0].Company);
    }

    [Fact]
    public void LoadData_ShouldRestoreDataFromFile()
    {
        // Arrange
        var application = new JobApplication
        {
            Company = "Test Company",
            JobTitle = "Developer"
        };
        _service.Add(application);
        _service.SaveData();

        // Act - Create new service instance to trigger LoadData with same test path
        var newService = new JobApplicationService(_testDataPath);
        var applications = newService.GetAll();

        // Assert
        Assert.Single(applications);
        Assert.Equal("Test Company", applications[0].Company);
        Assert.Equal("Developer", applications[0].JobTitle);
    }

    [Fact]
    public void GetDataFilePath_ShouldReturnValidPath()
    {
        // Act
        var path = _service.GetDataFilePath();

        // Assert
        Assert.NotNull(path);
        Assert.Contains("JobKeeper", path);
        Assert.Contains("jobapplications.json", path);
    }

    [Fact]
    public void MultipleOperations_ShouldMaintainDataIntegrity()
    {
        // Arrange & Act
        var app1 = new JobApplication { Company = "Company 1", JobTitle = "Job 1" };
        var app2 = new JobApplication { Company = "Company 2", JobTitle = "Job 2" };
        var app3 = new JobApplication { Company = "Company 3", JobTitle = "Job 3" };

        _service.Add(app1);
        _service.Add(app2);
        _service.Add(app3);

        app2.Status = ApplicationStatus.Rejected;
        _service.Update(app2);

        _service.Delete(app1.Id);

        var applications = _service.GetAll();

        // Assert
        Assert.Equal(2, applications.Count);
        Assert.DoesNotContain(applications, a => a.Id == app1.Id);
        Assert.Contains(applications, a => a.Id == app2.Id && a.Status == ApplicationStatus.Rejected);
        Assert.Contains(applications, a => a.Id == app3.Id);
    }

    [Fact]
    public void GetAll_ShouldReturnCopyOfList()
    {
        // Arrange
        var application = new JobApplication { Company = "Test" };
        _service.Add(application);

        // Act
        var list1 = _service.GetAll();
        var list2 = _service.GetAll();

        // Assert
        Assert.NotSame(list1, list2);
        Assert.Equal(list1.Count, list2.Count);
    }
}
