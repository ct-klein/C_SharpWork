namespace JobKeeper.WinForms.Forms;

partial class AboutForm
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        this.lblTitle = new System.Windows.Forms.Label();
        this.lblVersion = new System.Windows.Forms.Label();
        this.lblDescription = new System.Windows.Forms.Label();
        this.lblCopyright = new System.Windows.Forms.Label();
        this.btnOK = new System.Windows.Forms.Button();
        this.SuspendLayout();
        //
        // lblTitle
        //
        this.lblTitle.AutoSize = true;
        this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold);
        this.lblTitle.Location = new System.Drawing.Point(80, 30);
        this.lblTitle.Name = "lblTitle";
        this.lblTitle.Size = new System.Drawing.Size(113, 30);
        this.lblTitle.TabIndex = 0;
        this.lblTitle.Text = "JobKeeper";
        //
        // lblVersion
        //
        this.lblVersion.AutoSize = true;
        this.lblVersion.Location = new System.Drawing.Point(80, 70);
        this.lblVersion.Name = "lblVersion";
        this.lblVersion.Size = new System.Drawing.Size(69, 15);
        this.lblVersion.TabIndex = 1;
        this.lblVersion.Text = "Version 1.0.0";
        //
        // lblDescription
        //
        this.lblDescription.Location = new System.Drawing.Point(80, 100);
        this.lblDescription.Name = "lblDescription";
        this.lblDescription.Size = new System.Drawing.Size(320, 40);
        this.lblDescription.TabIndex = 2;
        this.lblDescription.Text = "A simple and efficient job application tracking system to help manage your job search.";
        //
        // lblCopyright
        //
        this.lblCopyright.AutoSize = true;
        this.lblCopyright.Location = new System.Drawing.Point(80, 150);
        this.lblCopyright.Name = "lblCopyright";
        this.lblCopyright.Size = new System.Drawing.Size(178, 15);
        this.lblCopyright.TabIndex = 3;
        this.lblCopyright.Text = "© 2025 JobKeeper. All rights reserved.";
        //
        // btnOK
        //
        this.btnOK.Location = new System.Drawing.Point(175, 190);
        this.btnOK.Name = "btnOK";
        this.btnOK.Size = new System.Drawing.Size(100, 30);
        this.btnOK.TabIndex = 4;
        this.btnOK.Text = "OK";
        this.btnOK.UseVisualStyleBackColor = true;
        this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
        //
        // AboutForm
        //
        this.AcceptButton = this.btnOK;
        this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(450, 250);
        this.Controls.Add(this.btnOK);
        this.Controls.Add(this.lblCopyright);
        this.Controls.Add(this.lblDescription);
        this.Controls.Add(this.lblVersion);
        this.Controls.Add(this.lblTitle);
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.Name = "AboutForm";
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
        this.Text = "About JobKeeper";
        this.ResumeLayout(false);
        this.PerformLayout();
    }

    private Label lblTitle;
    private Label lblVersion;
    private Label lblDescription;
    private Label lblCopyright;
    private Button btnOK;
}
