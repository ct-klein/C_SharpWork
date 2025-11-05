namespace JobKeeper.WinForms.Models;

/// <summary>
/// Represents a job application record
/// </summary>
public class JobApplication
{
    /// <summary>
    /// Unique identifier for the job application
    /// </summary>
    public Guid Id { get; set; }

    /// <summary>
    /// Company name
    /// </summary>
    public string Company { get; set; } = string.Empty;

    /// <summary>
    /// Company website URL
    /// </summary>
    public string Website { get; set; } = string.Empty;

    /// <summary>
    /// Job title
    /// </summary>
    public string JobTitle { get; set; } = string.Empty;

    /// <summary>
    /// Date the application was submitted
    /// </summary>
    public DateTime? Submitted { get; set; }

    /// <summary>
    /// Path to resume file
    /// </summary>
    public string Resume { get; set; } = string.Empty;

    /// <summary>
    /// Path to cover letter file
    /// </summary>
    public string CoverLetter { get; set; } = string.Empty;

    /// <summary>
    /// Application status
    /// </summary>
    public ApplicationStatus Status { get; set; }

    /// <summary>
    /// First interview date
    /// </summary>
    public DateTime? Interview1 { get; set; }

    /// <summary>
    /// Second interview date
    /// </summary>
    public DateTime? Interview2 { get; set; }

    public JobApplication()
    {
        Id = Guid.NewGuid();
        Status = ApplicationStatus.Submitted;
    }
}

/// <summary>
/// Application status enumeration
/// </summary>
public enum ApplicationStatus
{
    Submitted,
    Rejected,
    InterviewChanged,
    Ghosted
}
