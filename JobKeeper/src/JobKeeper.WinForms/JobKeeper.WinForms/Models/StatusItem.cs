namespace JobKeeper.WinForms.Models;

/// <summary>
/// Represents a status item with optional icon for ComboBox display
/// </summary>
public class StatusItem
{
    public string Text { get; set; } = string.Empty;
    public ApplicationStatus Status { get; set; }
    public Image? Icon { get; set; }

    public StatusItem(string text, ApplicationStatus status, Image? icon = null)
    {
        Text = text;
        Status = status;
        Icon = icon;
    }

    public override string ToString()
    {
        return Text;
    }
}
