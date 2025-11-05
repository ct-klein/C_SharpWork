using System.Drawing.Printing;
using JobKeeper.WinForms.Models;

namespace JobKeeper.WinForms.Forms;

/// <summary>
/// Print preview and printing form
/// </summary>
public partial class PrintPreviewForm : Form
{
    private readonly List<JobApplication> _applications;
    private int _currentPage = 0;
    private int _totalPages = 0;

    public PrintPreviewForm(List<JobApplication> applications)
    {
        InitializeComponent();
        _applications = applications;
    }

    private void btnPrint_Click(object sender, EventArgs e)
    {
        using var printDialog = new PrintDialog();
        printDialog.Document = printDocument1;

        if (printDialog.ShowDialog() == DialogResult.OK)
        {
            printDocument1.Print();
        }
    }

    private void btnPreview_Click(object sender, EventArgs e)
    {
        printPreviewControl1.Document = printDocument1;
        printPreviewControl1.Zoom = 1.0;
    }

    private void btnClose_Click(object sender, EventArgs e)
    {
        this.Close();
    }

    private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
    {
        var font = new Font("Arial", 10);
        var titleFont = new Font("Arial", 16, FontStyle.Bold);
        var headerFont = new Font("Arial", 12, FontStyle.Bold);
        float yPos = 100;
        float leftMargin = e.MarginBounds.Left;
        float topMargin = e.MarginBounds.Top;

        // Print title
        e.Graphics!.DrawString("JobKeeper - Job Application Report", titleFont, Brushes.Black, leftMargin, topMargin);
        yPos = topMargin + 50;

        // Print date
        e.Graphics.DrawString($"Report Date: {DateTime.Now:MM/dd/yyyy}", font, Brushes.Black, leftMargin, yPos);
        yPos += 30;

        e.Graphics.DrawString($"Total Applications: {_applications.Count}", font, Brushes.Black, leftMargin, yPos);
        yPos += 40;

        // Print applications
        int itemsPerPage = 5;
        int startIndex = _currentPage * itemsPerPage;
        int endIndex = Math.Min(startIndex + itemsPerPage, _applications.Count);

        for (int i = startIndex; i < endIndex; i++)
        {
            var app = _applications[i];

            // Draw separator line
            e.Graphics.DrawLine(Pens.Black, leftMargin, yPos, leftMargin + 700, yPos);
            yPos += 10;

            e.Graphics.DrawString($"Company: {app.Company}", headerFont, Brushes.Black, leftMargin, yPos);
            yPos += 25;
            e.Graphics.DrawString($"Job Title: {app.JobTitle}", font, Brushes.Black, leftMargin, yPos);
            yPos += 20;
            e.Graphics.DrawString($"Website: {app.Website}", font, Brushes.Black, leftMargin, yPos);
            yPos += 20;
            e.Graphics.DrawString($"Status: {app.Status}", font, Brushes.Black, leftMargin, yPos);
            yPos += 20;
            e.Graphics.DrawString($"Submitted: {(app.Submitted.HasValue ? app.Submitted.Value.ToString("MM/dd/yyyy") : "N/A")}", font, Brushes.Black, leftMargin, yPos);
            yPos += 20;
            e.Graphics.DrawString($"Interview 1: {(app.Interview1.HasValue ? app.Interview1.Value.ToString("MM/dd/yyyy") : "N/A")}", font, Brushes.Black, leftMargin, yPos);
            yPos += 20;
            e.Graphics.DrawString($"Interview 2: {(app.Interview2.HasValue ? app.Interview2.Value.ToString("MM/dd/yyyy") : "N/A")}", font, Brushes.Black, leftMargin, yPos);
            yPos += 30;
        }

        // Print page number
        e.Graphics.DrawString($"Page {_currentPage + 1} of {_totalPages}", font, Brushes.Black, leftMargin, e.MarginBounds.Bottom);

        _currentPage++;
        e.HasMorePages = (_currentPage * itemsPerPage < _applications.Count);

        if (!e.HasMorePages)
        {
            _currentPage = 0; // Reset for next print
        }
    }

    private void printDocument1_BeginPrint(object sender, PrintEventArgs e)
    {
        _currentPage = 0;
        _totalPages = (_applications.Count + 4) / 5; // 5 items per page
    }

    private void PrintPreviewForm_Load(object sender, EventArgs e)
    {
        printPreviewControl1.Document = printDocument1;
        _totalPages = (_applications.Count + 4) / 5;
    }
}
