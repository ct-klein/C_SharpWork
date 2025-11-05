namespace JobKeeper.WinForms.Forms;

/// <summary>
/// About dialog for the application
/// </summary>
public partial class AboutForm : Form
{
    public AboutForm()
    {
        InitializeComponent();
    }

    private void btnOK_Click(object sender, EventArgs e)
    {
        this.Close();
    }
}
