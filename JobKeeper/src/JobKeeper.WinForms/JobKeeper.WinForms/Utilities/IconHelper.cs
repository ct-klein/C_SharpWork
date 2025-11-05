namespace JobKeeper.WinForms.Utilities;

/// <summary>
/// Helper class for creating status icons
/// </summary>
public static class IconHelper
{
    /// <summary>
    /// Creates a ghost icon for the GHOSTED status
    /// </summary>
    public static Image CreateGhostIcon(int width = 16, int height = 16)
    {
        var bitmap = new Bitmap(width, height);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // Clear background
            g.Clear(Color.Transparent);

            // Draw ghost body (rounded shape)
            using (var brush = new SolidBrush(Color.FromArgb(200, 220, 220, 220)))
            {
                // Head (circle)
                g.FillEllipse(brush, 2, 1, 12, 12);

                // Body (rectangle with wavy bottom)
                g.FillRectangle(brush, 2, 7, 12, 7);

                // Wavy bottom (three small rectangles)
                g.FillRectangle(brush, 2, 14, 3, 2);
                g.FillRectangle(brush, 6, 14, 4, 2);
                g.FillRectangle(brush, 11, 14, 3, 2);
            }

            // Draw eyes (dark circles)
            using (var pen = new Pen(Color.FromArgb(100, 100, 100), 1.5f))
            {
                g.FillEllipse(Brushes.DarkGray, 5, 5, 2, 3);
                g.FillEllipse(Brushes.DarkGray, 9, 5, 2, 3);
            }

            // Add outline for better visibility
            using (var pen = new Pen(Color.FromArgb(150, 180, 180, 180), 1))
            {
                g.DrawEllipse(pen, 2, 1, 12, 12);
            }
        }

        return bitmap;
    }

    /// <summary>
    /// Creates a checkmark icon for SUBMITTED status
    /// </summary>
    public static Image CreateCheckIcon(int width = 16, int height = 16)
    {
        var bitmap = new Bitmap(width, height);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);

            using (var pen = new Pen(Color.Green, 2))
            {
                g.DrawLine(pen, 4, 8, 7, 11);
                g.DrawLine(pen, 7, 11, 12, 4);
            }
        }
        return bitmap;
    }

    /// <summary>
    /// Creates an X icon for REJECTED status
    /// </summary>
    public static Image CreateRejectIcon(int width = 16, int height = 16)
    {
        var bitmap = new Bitmap(width, height);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);

            using (var pen = new Pen(Color.Red, 2))
            {
                g.DrawLine(pen, 4, 4, 12, 12);
                g.DrawLine(pen, 12, 4, 4, 12);
            }
        }
        return bitmap;
    }

    /// <summary>
    /// Creates a calendar/clock icon for INTERVIEW CHANGED status
    /// </summary>
    public static Image CreateCalendarIcon(int width = 16, int height = 16)
    {
        var bitmap = new Bitmap(width, height);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);

            using (var pen = new Pen(Color.Blue, 1.5f))
            {
                // Calendar outline
                g.DrawRectangle(pen, 3, 4, 10, 9);
                // Top bar
                g.DrawLine(pen, 3, 6, 13, 6);
                // Binding rings
                g.DrawLine(pen, 5, 3, 5, 5);
                g.DrawLine(pen, 11, 3, 11, 5);
            }
        }
        return bitmap;
    }
}
