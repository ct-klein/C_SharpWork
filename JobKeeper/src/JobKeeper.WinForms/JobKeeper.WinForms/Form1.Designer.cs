namespace JobKeeper.WinForms;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        this.menuStrip1 = new System.Windows.Forms.MenuStrip();
        this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
        this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
        this.printToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
        this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
        this.groupBox1 = new System.Windows.Forms.GroupBox();
        this.groupBox2 = new System.Windows.Forms.GroupBox();
        this.dtpFilterStart = new System.Windows.Forms.DateTimePicker();
        this.dtpFilterEnd = new System.Windows.Forms.DateTimePicker();
        this.btnApplyFilter = new System.Windows.Forms.Button();
        this.btnClearFilter = new System.Windows.Forms.Button();
        this.label2 = new System.Windows.Forms.Label();
        this.label12 = new System.Windows.Forms.Label();
        this.btnBrowseCoverLetter = new System.Windows.Forms.Button();
        this.btnBrowseResume = new System.Windows.Forms.Button();
        this.dtpInterview2 = new System.Windows.Forms.DateTimePicker();
        this.dtpInterview1 = new System.Windows.Forms.DateTimePicker();
        this.cmbStatus = new System.Windows.Forms.ComboBox();
        this.txtCoverLetter = new System.Windows.Forms.TextBox();
        this.txtResume = new System.Windows.Forms.TextBox();
        this.dtpSubmitted = new System.Windows.Forms.DateTimePicker();
        this.txtJobTitle = new System.Windows.Forms.TextBox();
        this.txtWebsite = new System.Windows.Forms.TextBox();
        this.txtCompany = new System.Windows.Forms.TextBox();
        this.label11 = new System.Windows.Forms.Label();
        this.label10 = new System.Windows.Forms.Label();
        this.label9 = new System.Windows.Forms.Label();
        this.label8 = new System.Windows.Forms.Label();
        this.label7 = new System.Windows.Forms.Label();
        this.label6 = new System.Windows.Forms.Label();
        this.label5 = new System.Windows.Forms.Label();
        this.label4 = new System.Windows.Forms.Label();
        this.label3 = new System.Windows.Forms.Label();
        this.btnAdd = new System.Windows.Forms.Button();
        this.btnUpdate = new System.Windows.Forms.Button();
        this.btnDelete = new System.Windows.Forms.Button();
        this.btnClear = new System.Windows.Forms.Button();
        this.dgvApplications = new System.Windows.Forms.DataGridView();
        this.label1 = new System.Windows.Forms.Label();
        this.menuStrip1.SuspendLayout();
        this.groupBox1.SuspendLayout();
        this.groupBox2.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dgvApplications)).BeginInit();
        this.SuspendLayout();
        //
        // menuStrip1
        //
        this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.aboutToolStripMenuItem});
        this.menuStrip1.Location = new System.Drawing.Point(0, 0);
        this.menuStrip1.Name = "menuStrip1";
        this.menuStrip1.Size = new System.Drawing.Size(1200, 24);
        this.menuStrip1.TabIndex = 0;
        this.menuStrip1.Text = "menuStrip1";
        //
        // fileToolStripMenuItem
        //
        this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolStripMenuItem,
            this.printToolStripMenuItem});
        this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
        this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
        this.fileToolStripMenuItem.Text = "File";
        //
        // saveToolStripMenuItem
        //
        this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
        this.saveToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
        this.saveToolStripMenuItem.Text = "Save";
        this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
        //
        // printToolStripMenuItem
        //
        this.printToolStripMenuItem.Name = "printToolStripMenuItem";
        this.printToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
        this.printToolStripMenuItem.Text = "Print";
        this.printToolStripMenuItem.Click += new System.EventHandler(this.printToolStripMenuItem_Click);
        //
        // aboutToolStripMenuItem
        //
        this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
        this.aboutToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
        this.aboutToolStripMenuItem.Text = "About";
        this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
        //
        // groupBox1
        //
        this.groupBox1.Controls.Add(this.btnBrowseCoverLetter);
        this.groupBox1.Controls.Add(this.btnBrowseResume);
        this.groupBox1.Controls.Add(this.dtpInterview2);
        this.groupBox1.Controls.Add(this.dtpInterview1);
        this.groupBox1.Controls.Add(this.cmbStatus);
        this.groupBox1.Controls.Add(this.txtCoverLetter);
        this.groupBox1.Controls.Add(this.txtResume);
        this.groupBox1.Controls.Add(this.dtpSubmitted);
        this.groupBox1.Controls.Add(this.txtJobTitle);
        this.groupBox1.Controls.Add(this.txtWebsite);
        this.groupBox1.Controls.Add(this.txtCompany);
        this.groupBox1.Controls.Add(this.label11);
        this.groupBox1.Controls.Add(this.label10);
        this.groupBox1.Controls.Add(this.label9);
        this.groupBox1.Controls.Add(this.label8);
        this.groupBox1.Controls.Add(this.label7);
        this.groupBox1.Controls.Add(this.label6);
        this.groupBox1.Controls.Add(this.label5);
        this.groupBox1.Controls.Add(this.label4);
        this.groupBox1.Controls.Add(this.label3);
        this.groupBox1.Location = new System.Drawing.Point(12, 40);
        this.groupBox1.Name = "groupBox1";
        this.groupBox1.Size = new System.Drawing.Size(1176, 200);
        this.groupBox1.TabIndex = 1;
        this.groupBox1.TabStop = false;
        this.groupBox1.Text = "Job Application Details";
        //
        // btnBrowseCoverLetter
        //
        this.btnBrowseCoverLetter.Location = new System.Drawing.Point(1080, 110);
        this.btnBrowseCoverLetter.Name = "btnBrowseCoverLetter";
        this.btnBrowseCoverLetter.Size = new System.Drawing.Size(75, 23);
        this.btnBrowseCoverLetter.TabIndex = 19;
        this.btnBrowseCoverLetter.Text = "Browse...";
        this.btnBrowseCoverLetter.UseVisualStyleBackColor = true;
        this.btnBrowseCoverLetter.Click += new System.EventHandler(this.btnBrowseCoverLetter_Click);
        //
        // btnBrowseResume
        //
        this.btnBrowseResume.Location = new System.Drawing.Point(1080, 80);
        this.btnBrowseResume.Name = "btnBrowseResume";
        this.btnBrowseResume.Size = new System.Drawing.Size(75, 23);
        this.btnBrowseResume.TabIndex = 18;
        this.btnBrowseResume.Text = "Browse...";
        this.btnBrowseResume.UseVisualStyleBackColor = true;
        this.btnBrowseResume.Click += new System.EventHandler(this.btnBrowseResume_Click);
        //
        // dtpInterview2
        //
        this.dtpInterview2.Checked = false;
        this.dtpInterview2.CustomFormat = "MM/dd/yyyy";
        this.dtpInterview2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
        this.dtpInterview2.Location = new System.Drawing.Point(910, 170);
        this.dtpInterview2.Name = "dtpInterview2";
        this.dtpInterview2.ShowCheckBox = true;
        this.dtpInterview2.Size = new System.Drawing.Size(250, 23);
        this.dtpInterview2.TabIndex = 17;
        //
        // dtpInterview1
        //
        this.dtpInterview1.Checked = false;
        this.dtpInterview1.CustomFormat = "MM/dd/yyyy";
        this.dtpInterview1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
        this.dtpInterview1.Location = new System.Drawing.Point(910, 140);
        this.dtpInterview1.Name = "dtpInterview1";
        this.dtpInterview1.ShowCheckBox = true;
        this.dtpInterview1.Size = new System.Drawing.Size(250, 23);
        this.dtpInterview1.TabIndex = 16;
        //
        // cmbStatus
        //
        this.cmbStatus.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
        this.cmbStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        this.cmbStatus.FormattingEnabled = true;
        this.cmbStatus.ItemHeight = 20;
        this.cmbStatus.Location = new System.Drawing.Point(515, 140);
        this.cmbStatus.Name = "cmbStatus";
        this.cmbStatus.Size = new System.Drawing.Size(250, 26);
        this.cmbStatus.TabIndex = 15;
        //
        // txtCoverLetter
        //
        this.txtCoverLetter.Location = new System.Drawing.Point(515, 110);
        this.txtCoverLetter.Name = "txtCoverLetter";
        this.txtCoverLetter.ReadOnly = true;
        this.txtCoverLetter.Size = new System.Drawing.Size(559, 23);
        this.txtCoverLetter.TabIndex = 14;
        //
        // txtResume
        //
        this.txtResume.Location = new System.Drawing.Point(515, 80);
        this.txtResume.Name = "txtResume";
        this.txtResume.ReadOnly = true;
        this.txtResume.Size = new System.Drawing.Size(559, 23);
        this.txtResume.TabIndex = 13;
        //
        // dtpSubmitted
        //
        this.dtpSubmitted.Checked = false;
        this.dtpSubmitted.CustomFormat = "MM/dd/yyyy";
        this.dtpSubmitted.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
        this.dtpSubmitted.Location = new System.Drawing.Point(515, 50);
        this.dtpSubmitted.Name = "dtpSubmitted";
        this.dtpSubmitted.ShowCheckBox = true;
        this.dtpSubmitted.Size = new System.Drawing.Size(250, 23);
        this.dtpSubmitted.TabIndex = 12;
        //
        // txtJobTitle
        //
        this.txtJobTitle.Location = new System.Drawing.Point(120, 110);
        this.txtJobTitle.Name = "txtJobTitle";
        this.txtJobTitle.Size = new System.Drawing.Size(250, 23);
        this.txtJobTitle.TabIndex = 11;
        //
        // txtWebsite
        //
        this.txtWebsite.Location = new System.Drawing.Point(120, 80);
        this.txtWebsite.Name = "txtWebsite";
        this.txtWebsite.PlaceholderText = "https://example.com";
        this.txtWebsite.Size = new System.Drawing.Size(250, 23);
        this.txtWebsite.TabIndex = 10;
        //
        // txtCompany
        //
        this.txtCompany.Location = new System.Drawing.Point(120, 50);
        this.txtCompany.Name = "txtCompany";
        this.txtCompany.Size = new System.Drawing.Size(250, 23);
        this.txtCompany.TabIndex = 9;
        //
        // label11
        //
        this.label11.AutoSize = true;
        this.label11.Location = new System.Drawing.Point(800, 173);
        this.label11.Name = "label11";
        this.label11.Size = new System.Drawing.Size(92, 15);
        this.label11.TabIndex = 8;
        this.label11.Text = "INTERVIEW 2:";
        //
        // label10
        //
        this.label10.AutoSize = true;
        this.label10.Location = new System.Drawing.Point(800, 143);
        this.label10.Name = "label10";
        this.label10.Size = new System.Drawing.Size(92, 15);
        this.label10.TabIndex = 7;
        this.label10.Text = "INTERVIEW 1:";
        //
        // label9
        //
        this.label9.AutoSize = true;
        this.label9.Location = new System.Drawing.Point(440, 143);
        this.label9.Name = "label9";
        this.label9.Size = new System.Drawing.Size(55, 15);
        this.label9.TabIndex = 6;
        this.label9.Text = "STATUS:";
        //
        // label8
        //
        this.label8.AutoSize = true;
        this.label8.Location = new System.Drawing.Point(440, 113);
        this.label8.Name = "label8";
        this.label8.Size = new System.Drawing.Size(53, 15);
        this.label8.TabIndex = 5;
        this.label8.Text = "COVER:";
        //
        // label7
        //
        this.label7.AutoSize = true;
        this.label7.Location = new System.Drawing.Point(440, 83);
        this.label7.Name = "label7";
        this.label7.Size = new System.Drawing.Size(59, 15);
        this.label7.TabIndex = 4;
        this.label7.Text = "RESUME:";
        //
        // label6
        //
        this.label6.AutoSize = true;
        this.label6.Location = new System.Drawing.Point(440, 53);
        this.label6.Name = "label6";
        this.label6.Size = new System.Drawing.Size(75, 15);
        this.label6.TabIndex = 3;
        this.label6.Text = "SUBMITTED:";
        //
        // label5
        //
        this.label5.AutoSize = true;
        this.label5.Location = new System.Drawing.Point(20, 113);
        this.label5.Name = "label5";
        this.label5.Size = new System.Drawing.Size(65, 15);
        this.label5.TabIndex = 2;
        this.label5.Text = "JOB TITLE:";
        //
        // label4
        //
        this.label4.AutoSize = true;
        this.label4.Location = new System.Drawing.Point(20, 83);
        this.label4.Name = "label4";
        this.label4.Size = new System.Drawing.Size(60, 15);
        this.label4.TabIndex = 1;
        this.label4.Text = "WEBSITE:";
        //
        // label3
        //
        this.label3.AutoSize = true;
        this.label3.Location = new System.Drawing.Point(20, 53);
        this.label3.Name = "label3";
        this.label3.Size = new System.Drawing.Size(73, 15);
        this.label3.TabIndex = 0;
        this.label3.Text = "COMPANY:";
        //
        // btnAdd
        //
        this.btnAdd.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
        this.btnAdd.Location = new System.Drawing.Point(12, 250);
        this.btnAdd.Name = "btnAdd";
        this.btnAdd.Size = new System.Drawing.Size(100, 30);
        this.btnAdd.TabIndex = 2;
        this.btnAdd.Text = "Add";
        this.btnAdd.UseVisualStyleBackColor = true;
        this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
        //
        // btnUpdate
        //
        this.btnUpdate.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
        this.btnUpdate.Location = new System.Drawing.Point(118, 250);
        this.btnUpdate.Name = "btnUpdate";
        this.btnUpdate.Size = new System.Drawing.Size(100, 30);
        this.btnUpdate.TabIndex = 3;
        this.btnUpdate.Text = "Update";
        this.btnUpdate.UseVisualStyleBackColor = true;
        this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
        //
        // btnDelete
        //
        this.btnDelete.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
        this.btnDelete.Location = new System.Drawing.Point(224, 250);
        this.btnDelete.Name = "btnDelete";
        this.btnDelete.Size = new System.Drawing.Size(100, 30);
        this.btnDelete.TabIndex = 4;
        this.btnDelete.Text = "Delete";
        this.btnDelete.UseVisualStyleBackColor = true;
        this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
        //
        // btnClear
        //
        this.btnClear.Location = new System.Drawing.Point(330, 250);
        this.btnClear.Name = "btnClear";
        this.btnClear.Size = new System.Drawing.Size(100, 30);
        this.btnClear.TabIndex = 5;
        this.btnClear.Text = "Clear";
        this.btnClear.UseVisualStyleBackColor = true;
        this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
        //
        // dgvApplications
        //
        this.dgvApplications.AllowUserToAddRows = false;
        this.dgvApplications.AllowUserToDeleteRows = false;
        this.dgvApplications.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
        this.dgvApplications.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dgvApplications.Location = new System.Drawing.Point(12, 320);
        this.dgvApplications.MultiSelect = false;
        this.dgvApplications.Name = "dgvApplications";
        this.dgvApplications.ReadOnly = true;
        this.dgvApplications.RowTemplate.Height = 25;
        this.dgvApplications.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
        this.dgvApplications.Size = new System.Drawing.Size(1176, 300);
        this.dgvApplications.TabIndex = 6;
        this.dgvApplications.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvApplications_CellDoubleClick);
        //
        // groupBox2
        //
        this.groupBox2.Controls.Add(this.label12);
        this.groupBox2.Controls.Add(this.label2);
        this.groupBox2.Controls.Add(this.btnClearFilter);
        this.groupBox2.Controls.Add(this.btnApplyFilter);
        this.groupBox2.Controls.Add(this.dtpFilterEnd);
        this.groupBox2.Controls.Add(this.dtpFilterStart);
        this.groupBox2.Location = new System.Drawing.Point(440, 250);
        this.groupBox2.Name = "groupBox2";
        this.groupBox2.Size = new System.Drawing.Size(748, 60);
        this.groupBox2.TabIndex = 8;
        this.groupBox2.TabStop = false;
        this.groupBox2.Text = "Filter by Submitted Date";
        //
        // dtpFilterStart
        //
        this.dtpFilterStart.Checked = false;
        this.dtpFilterStart.CustomFormat = "MM/dd/yyyy";
        this.dtpFilterStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
        this.dtpFilterStart.Location = new System.Drawing.Point(70, 25);
        this.dtpFilterStart.Name = "dtpFilterStart";
        this.dtpFilterStart.ShowCheckBox = true;
        this.dtpFilterStart.Size = new System.Drawing.Size(150, 23);
        this.dtpFilterStart.TabIndex = 0;
        //
        // dtpFilterEnd
        //
        this.dtpFilterEnd.Checked = false;
        this.dtpFilterEnd.CustomFormat = "MM/dd/yyyy";
        this.dtpFilterEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
        this.dtpFilterEnd.Location = new System.Drawing.Point(310, 25);
        this.dtpFilterEnd.Name = "dtpFilterEnd";
        this.dtpFilterEnd.ShowCheckBox = true;
        this.dtpFilterEnd.Size = new System.Drawing.Size(150, 23);
        this.dtpFilterEnd.TabIndex = 1;
        //
        // btnApplyFilter
        //
        this.btnApplyFilter.Location = new System.Drawing.Point(480, 23);
        this.btnApplyFilter.Name = "btnApplyFilter";
        this.btnApplyFilter.Size = new System.Drawing.Size(120, 27);
        this.btnApplyFilter.TabIndex = 2;
        this.btnApplyFilter.Text = "Apply Filter";
        this.btnApplyFilter.UseVisualStyleBackColor = true;
        this.btnApplyFilter.Click += new System.EventHandler(this.btnApplyFilter_Click);
        //
        // btnClearFilter
        //
        this.btnClearFilter.Location = new System.Drawing.Point(610, 23);
        this.btnClearFilter.Name = "btnClearFilter";
        this.btnClearFilter.Size = new System.Drawing.Size(120, 27);
        this.btnClearFilter.TabIndex = 3;
        this.btnClearFilter.Text = "Clear Filter";
        this.btnClearFilter.UseVisualStyleBackColor = true;
        this.btnClearFilter.Click += new System.EventHandler(this.btnClearFilter_Click);
        //
        // label2
        //
        this.label2.AutoSize = true;
        this.label2.Location = new System.Drawing.Point(15, 28);
        this.label2.Name = "label2";
        this.label2.Size = new System.Drawing.Size(38, 15);
        this.label2.TabIndex = 4;
        this.label2.Text = "From:";
        //
        // label12
        //
        this.label12.AutoSize = true;
        this.label12.Location = new System.Drawing.Point(240, 28);
        this.label12.Name = "label12";
        this.label12.Size = new System.Drawing.Size(22, 15);
        this.label12.TabIndex = 5;
        this.label12.Text = "To:";
        //
        // label1
        //
        this.label1.AutoSize = true;
        this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
        this.label1.Location = new System.Drawing.Point(12, 295);
        this.label1.Name = "label1";
        this.label1.Size = new System.Drawing.Size(340, 15);
        this.label1.TabIndex = 7;
        this.label1.Text = "Job Applications (Double-click a row to edit):";
        //
        // Form1
        //
        this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(1200, 632);
        this.Controls.Add(this.groupBox2);
        this.Controls.Add(this.label1);
        this.Controls.Add(this.dgvApplications);
        this.Controls.Add(this.btnClear);
        this.Controls.Add(this.btnDelete);
        this.Controls.Add(this.btnUpdate);
        this.Controls.Add(this.btnAdd);
        this.Controls.Add(this.groupBox1);
        this.Controls.Add(this.menuStrip1);
        this.MainMenuStrip = this.menuStrip1;
        this.Name = "Form1";
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        this.Text = "JobKeeper - Job Application Tracker";
        this.menuStrip1.ResumeLayout(false);
        this.menuStrip1.PerformLayout();
        this.groupBox1.ResumeLayout(false);
        this.groupBox1.PerformLayout();
        this.groupBox2.ResumeLayout(false);
        this.groupBox2.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dgvApplications)).EndInit();
        this.ResumeLayout(false);
        this.PerformLayout();
    }

    #endregion

    private MenuStrip menuStrip1;
    private ToolStripMenuItem fileToolStripMenuItem;
    private ToolStripMenuItem saveToolStripMenuItem;
    private ToolStripMenuItem printToolStripMenuItem;
    private ToolStripMenuItem aboutToolStripMenuItem;
    private GroupBox groupBox1;
    private TextBox txtCompany;
    private Label label3;
    private Label label4;
    private Label label5;
    private Label label6;
    private Label label7;
    private Label label8;
    private Label label9;
    private Label label10;
    private Label label11;
    private TextBox txtWebsite;
    private TextBox txtJobTitle;
    private DateTimePicker dtpSubmitted;
    private TextBox txtResume;
    private TextBox txtCoverLetter;
    private ComboBox cmbStatus;
    private DateTimePicker dtpInterview1;
    private DateTimePicker dtpInterview2;
    private Button btnBrowseResume;
    private Button btnBrowseCoverLetter;
    private Button btnAdd;
    private Button btnUpdate;
    private Button btnDelete;
    private Button btnClear;
    private DataGridView dgvApplications;
    private Label label1;
    private GroupBox groupBox2;
    private DateTimePicker dtpFilterStart;
    private DateTimePicker dtpFilterEnd;
    private Button btnApplyFilter;
    private Button btnClearFilter;
    private Label label2;
    private Label label12;
}
