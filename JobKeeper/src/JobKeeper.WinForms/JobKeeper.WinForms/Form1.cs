using JobKeeper.WinForms.Models;
using JobKeeper.WinForms.Services;
using JobKeeper.WinForms.Utilities;

namespace JobKeeper.WinForms;

/// <summary>
/// Main form for JobKeeper application
/// </summary>
public partial class Form1 : Form
{
    private readonly JobApplicationService _service;
    private JobApplication? _currentEditingApplication;
    private DateTime? _filterStartDate;
    private DateTime? _filterEndDate;

    public Form1()
    {
        InitializeComponent();
        _service = new JobApplicationService();
        InitializeStatusComboBox();
        InitializeDataGridView();
        InitializeDateFilters();
        LoadApplications();
    }

    private void InitializeDataGridView()
    {
        // Hook up the CellPainting event for custom status rendering
        dgvApplications.CellPainting += DgvApplications_CellPainting;

        // Enable sorting on all columns
        dgvApplications.AutoGenerateColumns = true;
        dgvApplications.AllowUserToOrderColumns = true;
    }

    private void InitializeDateFilters()
    {
        // Set up date filter checkboxes
        dtpFilterStart.Checked = false;
        dtpFilterEnd.Checked = false;
    }

    private void btnApplyFilter_Click(object sender, EventArgs e)
    {
        _filterStartDate = dtpFilterStart.Checked ? dtpFilterStart.Value.Date : null;
        _filterEndDate = dtpFilterEnd.Checked ? dtpFilterEnd.Value.Date : null;
        LoadApplications();
    }

    private void btnClearFilter_Click(object sender, EventArgs e)
    {
        dtpFilterStart.Checked = false;
        dtpFilterEnd.Checked = false;
        _filterStartDate = null;
        _filterEndDate = null;
        LoadApplications();
    }

    private void DgvApplications_CellPainting(object? sender, DataGridViewCellPaintingEventArgs e)
    {
        // Check if this is the Status column and not the header row
        if (e.ColumnIndex >= 0 && e.RowIndex >= 0 &&
            dgvApplications.Columns[e.ColumnIndex].Name == "Status")
        {
            e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);

            // Get the status value
            var cellValue = e.Value;
            if (cellValue != null && Enum.TryParse<ApplicationStatus>(cellValue.ToString(), out var status))
            {
                // Get the corresponding icon
                Image? icon = status switch
                {
                    ApplicationStatus.Submitted => IconHelper.CreateCheckIcon(),
                    ApplicationStatus.Rejected => IconHelper.CreateRejectIcon(),
                    ApplicationStatus.InterviewChanged => IconHelper.CreateCalendarIcon(),
                    ApplicationStatus.Ghosted => IconHelper.CreateGhostIcon(),
                    _ => null
                };

                // Draw the icon
                if (icon != null)
                {
                    var iconRect = new Rectangle(e.CellBounds.Left + 4, e.CellBounds.Top + (e.CellBounds.Height - 16) / 2, 16, 16);
                    e.Graphics.DrawImage(icon, iconRect);
                }

                // Draw the text
                var statusText = status switch
                {
                    ApplicationStatus.Submitted => "SUBMITTED",
                    ApplicationStatus.Rejected => "REJECTED",
                    ApplicationStatus.InterviewChanged => "INTERVIEW CHANGED",
                    ApplicationStatus.Ghosted => "GHOSTED",
                    _ => status.ToString()
                };

                var textRect = new Rectangle(e.CellBounds.Left + 24, e.CellBounds.Top,
                    e.CellBounds.Width - 24, e.CellBounds.Height);

                TextRenderer.DrawText(e.Graphics, statusText, e.CellStyle.Font, textRect,
                    e.CellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
            }

            e.Handled = true;
        }
    }

    private void InitializeStatusComboBox()
    {
        // Set ComboBox to owner-drawn mode for custom rendering
        cmbStatus.DrawMode = DrawMode.OwnerDrawFixed;
        cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbStatus.ItemHeight = 20;

        cmbStatus.Items.Clear();
        cmbStatus.Items.Add(new StatusItem("SUBMITTED", ApplicationStatus.Submitted, IconHelper.CreateCheckIcon()));
        cmbStatus.Items.Add(new StatusItem("REJECTED", ApplicationStatus.Rejected, IconHelper.CreateRejectIcon()));
        cmbStatus.Items.Add(new StatusItem("INTERVIEW CHANGED", ApplicationStatus.InterviewChanged, IconHelper.CreateCalendarIcon()));
        cmbStatus.Items.Add(new StatusItem("GHOSTED", ApplicationStatus.Ghosted, IconHelper.CreateGhostIcon()));
        cmbStatus.SelectedIndex = 0;

        // Hook up the DrawItem event
        cmbStatus.DrawItem += CmbStatus_DrawItem;
    }

    private void CmbStatus_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0) return;

        e.DrawBackground();

        var item = (StatusItem)cmbStatus.Items[e.Index];

        // Draw the icon if available
        if (item.Icon != null)
        {
            e.Graphics.DrawImage(item.Icon, e.Bounds.Left + 2, e.Bounds.Top + 2, 16, 16);
        }

        // Draw the text
        using (var brush = new SolidBrush(e.ForeColor))
        {
            var textBounds = new Rectangle(e.Bounds.Left + 22, e.Bounds.Top, e.Bounds.Width - 22, e.Bounds.Height);
            e.Graphics.DrawString(item.Text, e.Font!, brush, textBounds, StringFormat.GenericDefault);
        }

        e.DrawFocusRectangle();
    }

    private void LoadApplications()
    {
        var applications = _service.GetAll();

        // Auto-update GHOSTED status for applications submitted more than 2 weeks ago
        UpdateGhostedStatus(applications);

        // Apply date range filter if set
        if (_filterStartDate.HasValue || _filterEndDate.HasValue)
        {
            applications = applications.Where(app =>
            {
                if (!app.Submitted.HasValue)
                    return false;

                if (_filterStartDate.HasValue && app.Submitted.Value.Date < _filterStartDate.Value)
                    return false;

                if (_filterEndDate.HasValue && app.Submitted.Value.Date > _filterEndDate.Value)
                    return false;

                return true;
            }).ToList();
        }

        dgvApplications.DataSource = null;
        dgvApplications.DataSource = applications;

        // Customize grid appearance
        if (dgvApplications.Columns.Count > 0 && dgvApplications.Columns.Contains("Status"))
        {
            dgvApplications.Columns["Id"].Visible = false;

            // Set column headers and enable sorting
            dgvApplications.Columns["Company"].HeaderText = "COMPANY";
            dgvApplications.Columns["Company"].SortMode = DataGridViewColumnSortMode.Automatic;

            dgvApplications.Columns["Website"].HeaderText = "WEBSITE";
            dgvApplications.Columns["Website"].SortMode = DataGridViewColumnSortMode.Automatic;

            dgvApplications.Columns["JobTitle"].HeaderText = "JOB TITLE";
            dgvApplications.Columns["JobTitle"].SortMode = DataGridViewColumnSortMode.Automatic;

            dgvApplications.Columns["Submitted"].HeaderText = "SUBMITTED";
            dgvApplications.Columns["Submitted"].SortMode = DataGridViewColumnSortMode.Automatic;

            dgvApplications.Columns["Resume"].HeaderText = "RESUME";
            dgvApplications.Columns["Resume"].SortMode = DataGridViewColumnSortMode.Automatic;

            dgvApplications.Columns["CoverLetter"].HeaderText = "COVER";
            dgvApplications.Columns["CoverLetter"].SortMode = DataGridViewColumnSortMode.Automatic;

            dgvApplications.Columns["Status"].HeaderText = "STATUS";
            dgvApplications.Columns["Status"].SortMode = DataGridViewColumnSortMode.Automatic;

            // Set fixed width for Status column to accommodate icon + text
            dgvApplications.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvApplications.Columns["Status"].Width = 180;

            dgvApplications.Columns["Interview1"].HeaderText = "INTERVIEW 1";
            dgvApplications.Columns["Interview1"].SortMode = DataGridViewColumnSortMode.Automatic;

            dgvApplications.Columns["Interview2"].HeaderText = "INTERVIEW 2";
            dgvApplications.Columns["Interview2"].SortMode = DataGridViewColumnSortMode.Automatic;
        }

        // Increase row height to better display icons
        dgvApplications.RowTemplate.Height = 28;
    }

    private void UpdateGhostedStatus(List<JobApplication> applications)
    {
        var twoWeeksAgo = DateTime.Now.AddDays(-14);
        bool hasChanges = false;

        foreach (var app in applications)
        {
            // Only auto-update if status is still SUBMITTED and submitted date is more than 2 weeks ago
            if (app.Status == ApplicationStatus.Submitted &&
                app.Submitted.HasValue &&
                app.Submitted.Value < twoWeeksAgo)
            {
                app.Status = ApplicationStatus.Ghosted;
                _service.Update(app);
                hasChanges = true;
            }
        }

        if (hasChanges)
        {
            _service.SaveData();
        }
    }

    private void btnAdd_Click(object sender, EventArgs e)
    {
        if (!ValidateInputs())
            return;

        var application = new JobApplication
        {
            Company = txtCompany.Text.Trim(),
            Website = txtWebsite.Text.Trim(),
            JobTitle = txtJobTitle.Text.Trim(),
            Submitted = dtpSubmitted.Checked ? dtpSubmitted.Value : null,
            Resume = txtResume.Text.Trim(),
            CoverLetter = txtCoverLetter.Text.Trim(),
            Status = GetSelectedStatus(),
            Interview1 = dtpInterview1.Checked ? dtpInterview1.Value : null,
            Interview2 = dtpInterview2.Checked ? dtpInterview2.Value : null
        };

        _service.Add(application);
        LoadApplications();
        ClearInputs();
        MessageBox.Show("Job application added successfully!", "Success",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void btnUpdate_Click(object sender, EventArgs e)
    {
        if (_currentEditingApplication == null)
        {
            MessageBox.Show("Please select a record from the grid to update.", "No Selection",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!ValidateInputs())
            return;

        _currentEditingApplication.Company = txtCompany.Text.Trim();
        _currentEditingApplication.Website = txtWebsite.Text.Trim();
        _currentEditingApplication.JobTitle = txtJobTitle.Text.Trim();
        _currentEditingApplication.Submitted = dtpSubmitted.Checked ? dtpSubmitted.Value : null;
        _currentEditingApplication.Resume = txtResume.Text.Trim();
        _currentEditingApplication.CoverLetter = txtCoverLetter.Text.Trim();
        _currentEditingApplication.Status = GetSelectedStatus();
        _currentEditingApplication.Interview1 = dtpInterview1.Checked ? dtpInterview1.Value : null;
        _currentEditingApplication.Interview2 = dtpInterview2.Checked ? dtpInterview2.Value : null;

        _service.Update(_currentEditingApplication);
        LoadApplications();
        ClearInputs();
        _currentEditingApplication = null;
        MessageBox.Show("Job application updated successfully!", "Success",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void btnDelete_Click(object sender, EventArgs e)
    {
        if (dgvApplications.SelectedRows.Count == 0)
        {
            MessageBox.Show("Please select a record to delete.", "No Selection",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = MessageBox.Show("Are you sure you want to delete this job application?",
            "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            var application = (JobApplication)dgvApplications.SelectedRows[0].DataBoundItem;
            _service.Delete(application.Id);
            LoadApplications();
            ClearInputs();
            _currentEditingApplication = null;
            MessageBox.Show("Job application deleted successfully!", "Success",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void btnClear_Click(object sender, EventArgs e)
    {
        ClearInputs();
        _currentEditingApplication = null;
    }

    private void dgvApplications_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex >= 0)
        {
            var application = (JobApplication)dgvApplications.Rows[e.RowIndex].DataBoundItem;
            LoadApplicationToForm(application);
        }
    }

    private void LoadApplicationToForm(JobApplication application)
    {
        _currentEditingApplication = application;
        txtCompany.Text = application.Company;
        txtWebsite.Text = application.Website;
        txtJobTitle.Text = application.JobTitle;

        if (application.Submitted.HasValue)
        {
            dtpSubmitted.Value = application.Submitted.Value;
            dtpSubmitted.Checked = true;
        }
        else
        {
            dtpSubmitted.Checked = false;
        }

        txtResume.Text = application.Resume;
        txtCoverLetter.Text = application.CoverLetter;

        // Find and select the matching StatusItem
        for (int i = 0; i < cmbStatus.Items.Count; i++)
        {
            if (cmbStatus.Items[i] is StatusItem item && item.Status == application.Status)
            {
                cmbStatus.SelectedIndex = i;
                break;
            }
        }

        if (application.Interview1.HasValue)
        {
            dtpInterview1.Value = application.Interview1.Value;
            dtpInterview1.Checked = true;
        }
        else
        {
            dtpInterview1.Checked = false;
        }

        if (application.Interview2.HasValue)
        {
            dtpInterview2.Value = application.Interview2.Value;
            dtpInterview2.Checked = true;
        }
        else
        {
            dtpInterview2.Checked = false;
        }
    }

    private void btnBrowseResume_Click(object sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "PDF Files (*.pdf)|*.pdf|Word Documents (*.doc;*.docx)|*.doc;*.docx|All Files (*.*)|*.*",
            Title = "Select Resume File"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtResume.Text = dialog.FileName;
        }
    }

    private void btnBrowseCoverLetter_Click(object sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "PDF Files (*.pdf)|*.pdf|Word Documents (*.doc;*.docx)|*.doc;*.docx|All Files (*.*)|*.*",
            Title = "Select Cover Letter File"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtCoverLetter.Text = dialog.FileName;
        }
    }

    private void saveToolStripMenuItem_Click(object sender, EventArgs e)
    {
        _service.SaveData();
        MessageBox.Show("All data has been saved successfully!", "Save Complete",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void printToolStripMenuItem_Click(object sender, EventArgs e)
    {
        var printForm = new Forms.PrintPreviewForm(_service.GetAll());
        printForm.ShowDialog();
    }

    private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
    {
        var aboutForm = new Forms.AboutForm();
        aboutForm.ShowDialog();
    }

    private bool ValidateInputs()
    {
        if (string.IsNullOrWhiteSpace(txtCompany.Text))
        {
            MessageBox.Show("Please enter a company name.", "Validation Error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtCompany.Focus();
            return false;
        }

        if (!string.IsNullOrWhiteSpace(txtWebsite.Text) && !ValidationHelper.IsValidUrl(txtWebsite.Text))
        {
            MessageBox.Show("Please enter a valid URL (must start with http:// or https://)",
                "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtWebsite.Focus();
            return false;
        }

        if (string.IsNullOrWhiteSpace(txtJobTitle.Text))
        {
            MessageBox.Show("Please enter a job title.", "Validation Error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtJobTitle.Focus();
            return false;
        }

        return true;
    }

    private void ClearInputs()
    {
        txtCompany.Clear();
        txtWebsite.Clear();
        txtJobTitle.Clear();
        dtpSubmitted.Checked = false;
        txtResume.Clear();
        txtCoverLetter.Clear();
        cmbStatus.SelectedIndex = 0;
        dtpInterview1.Checked = false;
        dtpInterview2.Checked = false;
    }

    private ApplicationStatus GetSelectedStatus()
    {
        if (cmbStatus.SelectedItem is StatusItem item)
        {
            return item.Status;
        }
        return ApplicationStatus.Submitted;
    }
}
