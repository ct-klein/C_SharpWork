using System.Drawing;
using JobKeeper.WinForms.Utilities;

namespace JobKeeper.Tests.Utilities;

/// <summary>
/// Unit tests for IconHelper
/// </summary>
public class IconHelperTests
{
    [Fact]
    public void CreateGhostIcon_ShouldReturnValidImage()
    {
        // Act
        var icon = IconHelper.CreateGhostIcon();

        // Assert
        Assert.NotNull(icon);
        Assert.Equal(16, icon.Width);
        Assert.Equal(16, icon.Height);
    }

    [Theory]
    [InlineData(16, 16)]
    [InlineData(32, 32)]
    [InlineData(24, 24)]
    public void CreateGhostIcon_WithCustomSize_ShouldReturnCorrectSize(int width, int height)
    {
        // Act
        var icon = IconHelper.CreateGhostIcon(width, height);

        // Assert
        Assert.NotNull(icon);
        Assert.Equal(width, icon.Width);
        Assert.Equal(height, icon.Height);
    }

    [Fact]
    public void CreateGhostIcon_ShouldCreateBitmap()
    {
        // Act
        var icon = IconHelper.CreateGhostIcon();

        // Assert
        Assert.IsType<Bitmap>(icon);
    }

    [Fact]
    public void CreateCheckIcon_ShouldReturnValidImage()
    {
        // Act
        var icon = IconHelper.CreateCheckIcon();

        // Assert
        Assert.NotNull(icon);
        Assert.Equal(16, icon.Width);
        Assert.Equal(16, icon.Height);
    }

    [Theory]
    [InlineData(16, 16)]
    [InlineData(32, 32)]
    [InlineData(24, 24)]
    public void CreateCheckIcon_WithCustomSize_ShouldReturnCorrectSize(int width, int height)
    {
        // Act
        var icon = IconHelper.CreateCheckIcon(width, height);

        // Assert
        Assert.NotNull(icon);
        Assert.Equal(width, icon.Width);
        Assert.Equal(height, icon.Height);
    }

    [Fact]
    public void CreateCheckIcon_ShouldCreateBitmap()
    {
        // Act
        var icon = IconHelper.CreateCheckIcon();

        // Assert
        Assert.IsType<Bitmap>(icon);
    }

    [Fact]
    public void CreateRejectIcon_ShouldReturnValidImage()
    {
        // Act
        var icon = IconHelper.CreateRejectIcon();

        // Assert
        Assert.NotNull(icon);
        Assert.Equal(16, icon.Width);
        Assert.Equal(16, icon.Height);
    }

    [Theory]
    [InlineData(16, 16)]
    [InlineData(32, 32)]
    [InlineData(24, 24)]
    public void CreateRejectIcon_WithCustomSize_ShouldReturnCorrectSize(int width, int height)
    {
        // Act
        var icon = IconHelper.CreateRejectIcon(width, height);

        // Assert
        Assert.NotNull(icon);
        Assert.Equal(width, icon.Width);
        Assert.Equal(height, icon.Height);
    }

    [Fact]
    public void CreateRejectIcon_ShouldCreateBitmap()
    {
        // Act
        var icon = IconHelper.CreateRejectIcon();

        // Assert
        Assert.IsType<Bitmap>(icon);
    }

    [Fact]
    public void CreateCalendarIcon_ShouldReturnValidImage()
    {
        // Act
        var icon = IconHelper.CreateCalendarIcon();

        // Assert
        Assert.NotNull(icon);
        Assert.Equal(16, icon.Width);
        Assert.Equal(16, icon.Height);
    }

    [Theory]
    [InlineData(16, 16)]
    [InlineData(32, 32)]
    [InlineData(24, 24)]
    public void CreateCalendarIcon_WithCustomSize_ShouldReturnCorrectSize(int width, int height)
    {
        // Act
        var icon = IconHelper.CreateCalendarIcon(width, height);

        // Assert
        Assert.NotNull(icon);
        Assert.Equal(width, icon.Width);
        Assert.Equal(height, icon.Height);
    }

    [Fact]
    public void CreateCalendarIcon_ShouldCreateBitmap()
    {
        // Act
        var icon = IconHelper.CreateCalendarIcon();

        // Assert
        Assert.IsType<Bitmap>(icon);
    }

    [Fact]
    public void AllIcons_ShouldBeDisposable()
    {
        // Arrange & Act
        var ghostIcon = IconHelper.CreateGhostIcon();
        var checkIcon = IconHelper.CreateCheckIcon();
        var rejectIcon = IconHelper.CreateRejectIcon();
        var calendarIcon = IconHelper.CreateCalendarIcon();

        // Assert - Should not throw
        ghostIcon.Dispose();
        checkIcon.Dispose();
        rejectIcon.Dispose();
        calendarIcon.Dispose();
    }

    [Fact]
    public void AllIcons_ShouldHaveNonZeroPixels()
    {
        // Arrange
        var icons = new[]
        {
            IconHelper.CreateGhostIcon(),
            IconHelper.CreateCheckIcon(),
            IconHelper.CreateRejectIcon(),
            IconHelper.CreateCalendarIcon()
        };

        // Act & Assert
        foreach (var icon in icons)
        {
            var bitmap = (Bitmap)icon;
            bool hasNonTransparentPixel = false;

            for (int x = 0; x < bitmap.Width && !hasNonTransparentPixel; x++)
            {
                for (int y = 0; y < bitmap.Height && !hasNonTransparentPixel; y++)
                {
                    var pixel = bitmap.GetPixel(x, y);
                    if (pixel.A > 0) // Check alpha channel
                    {
                        hasNonTransparentPixel = true;
                    }
                }
            }

            Assert.True(hasNonTransparentPixel, "Icon should have at least one non-transparent pixel");
            icon.Dispose();
        }
    }
}
