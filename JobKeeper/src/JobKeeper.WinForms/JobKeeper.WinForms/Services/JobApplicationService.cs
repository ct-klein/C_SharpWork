using System.Text.Json;
using JobKeeper.WinForms.Models;

namespace JobKeeper.WinForms.Services;

/// <summary>
/// Service for managing job application data with JSON-based local storage
/// </summary>
public class JobApplicationService
{
    private readonly string _dataFilePath;
    private List<JobApplication> _applications;

    public JobApplicationService(string? customDataPath = null)
    {
        // Store data in user's AppData folder or use custom path for testing
        string appFolder;

        if (!string.IsNullOrEmpty(customDataPath))
        {
            appFolder = customDataPath;
        }
        else
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            appFolder = Path.Combine(appDataPath, "JobKeeper");
        }

        if (!Directory.Exists(appFolder))
        {
            Directory.CreateDirectory(appFolder);
        }

        _dataFilePath = Path.Combine(appFolder, "jobapplications.json");
        _applications = new List<JobApplication>();
        LoadData();
    }

    /// <summary>
    /// Gets all job applications
    /// </summary>
    public List<JobApplication> GetAll()
    {
        return new List<JobApplication>(_applications);
    }

    /// <summary>
    /// Adds a new job application
    /// </summary>
    public void Add(JobApplication application)
    {
        if (application == null)
            throw new ArgumentNullException(nameof(application));

        _applications.Add(application);
        SaveData();
    }

    /// <summary>
    /// Updates an existing job application
    /// </summary>
    public void Update(JobApplication application)
    {
        if (application == null)
            throw new ArgumentNullException(nameof(application));

        var index = _applications.FindIndex(a => a.Id == application.Id);
        if (index >= 0)
        {
            _applications[index] = application;
            SaveData();
        }
    }

    /// <summary>
    /// Deletes a job application
    /// </summary>
    public void Delete(Guid id)
    {
        var application = _applications.FirstOrDefault(a => a.Id == id);
        if (application != null)
        {
            _applications.Remove(application);
            SaveData();
        }
    }

    /// <summary>
    /// Saves all data to JSON file
    /// </summary>
    public void SaveData()
    {
        try
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            string jsonString = JsonSerializer.Serialize(_applications, options);
            File.WriteAllText(_dataFilePath, jsonString);
        }
        catch (Exception ex)
        {
            throw new IOException($"Failed to save data: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Loads data from JSON file
    /// </summary>
    private void LoadData()
    {
        try
        {
            if (File.Exists(_dataFilePath))
            {
                string jsonString = File.ReadAllText(_dataFilePath);
                var applications = JsonSerializer.Deserialize<List<JobApplication>>(jsonString);
                if (applications != null)
                {
                    _applications = applications;
                }
            }
        }
        catch (Exception ex)
        {
            // Log error but don't crash - start with empty list
            MessageBox.Show($"Warning: Could not load existing data: {ex.Message}",
                "Data Load Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    /// <summary>
    /// Gets the data file path for display purposes
    /// </summary>
    public string GetDataFilePath() => _dataFilePath;
}
