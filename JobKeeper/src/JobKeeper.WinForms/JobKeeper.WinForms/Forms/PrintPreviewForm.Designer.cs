using System.Drawing.Printing;

namespace JobKeeper.WinForms.Forms;

partial class PrintPreviewForm
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
        this.printPreviewControl1 = new System.Windows.Forms.PrintPreviewControl();
        this.btnPrint = new System.Windows.Forms.Button();
        this.btnClose = new System.Windows.Forms.Button();
        this.printDocument1 = new System.Drawing.Printing.PrintDocument();
        this.SuspendLayout();
        //
        // printPreviewControl1
        //
        this.printPreviewControl1.Location = new System.Drawing.Point(12, 12);
        this.printPreviewControl1.Name = "printPreviewControl1";
        this.printPreviewControl1.Size = new System.Drawing.Size(860, 500);
        this.printPreviewControl1.TabIndex = 0;
        this.printPreviewControl1.Zoom = 1D;
        //
        // btnPrint
        //
        this.btnPrint.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
        this.btnPrint.Location = new System.Drawing.Point(670, 530);
        this.btnPrint.Name = "btnPrint";
        this.btnPrint.Size = new System.Drawing.Size(100, 30);
        this.btnPrint.TabIndex = 1;
        this.btnPrint.Text = "Print";
        this.btnPrint.UseVisualStyleBackColor = true;
        this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
        //
        // btnClose
        //
        this.btnClose.Location = new System.Drawing.Point(776, 530);
        this.btnClose.Name = "btnClose";
        this.btnClose.Size = new System.Drawing.Size(100, 30);
        this.btnClose.TabIndex = 2;
        this.btnClose.Text = "Close";
        this.btnClose.UseVisualStyleBackColor = true;
        this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
        //
        // printDocument1
        //
        this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
        this.printDocument1.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.printDocument1_BeginPrint);
        //
        // PrintPreviewForm
        //
        this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(884, 572);
        this.Controls.Add(this.btnClose);
        this.Controls.Add(this.btnPrint);
        this.Controls.Add(this.printPreviewControl1);
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.Name = "PrintPreviewForm";
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
        this.Text = "Print Preview - JobKeeper";
        this.Load += new System.EventHandler(this.PrintPreviewForm_Load);
        this.ResumeLayout(false);
    }

    private PrintPreviewControl printPreviewControl1;
    private Button btnPrint;
    private Button btnClose;
    private PrintDocument printDocument1;
}
