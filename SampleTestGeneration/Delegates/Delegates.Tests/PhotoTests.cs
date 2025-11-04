using Xunit;
using Delegates;

namespace Delegates.Tests
{
    public class PhotoTests
    {
        [Fact]
        public void Load_WithValidPath_ReturnsPhotoInstance()
        {
            // Arrange
            string path = "test.jpg";

            // Act
            var photo = Photo.Load(path);

            // Assert
            Assert.NotNull(photo);
            Assert.IsType<Photo>(photo);
        }

        [Fact]
        public void Load_WithEmptyPath_ReturnsPhotoInstance()
        {
            // Arrange
            string path = string.Empty;

            // Act
            var photo = Photo.Load(path);

            // Assert
            Assert.NotNull(photo);
            Assert.IsType<Photo>(photo);
        }

        [Fact]
        public void Load_WithNullPath_ReturnsPhotoInstance()
        {
            // Arrange
            string? path = null;

            // Act
            var photo = Photo.Load(path!);

            // Assert
            Assert.NotNull(photo);
            Assert.IsType<Photo>(photo);
        }

        [Fact]
        public void Load_WithDifferentExtensions_ReturnsPhotoInstance()
        {
            // Arrange
            string[] paths = { "photo.jpg", "photo.png", "photo.gif", "photo.bmp" };

            foreach (var path in paths)
            {
                // Act
                var photo = Photo.Load(path);

                // Assert
                Assert.NotNull(photo);
                Assert.IsType<Photo>(photo);
            }
        }

        [Fact]
        public void Save_OnPhotoInstance_DoesNotThrowException()
        {
            // Arrange
            var photo = Photo.Load("test.jpg");

            // Act & Assert
            var exception = Record.Exception(() => photo.Save());
            Assert.Null(exception);
        }

        [Fact]
        public void Save_OnMultiplePhotos_DoesNotThrowException()
        {
            // Arrange
            var photo1 = Photo.Load("test1.jpg");
            var photo2 = Photo.Load("test2.jpg");

            // Act & Assert
            var exception1 = Record.Exception(() => photo1.Save());
            var exception2 = Record.Exception(() => photo2.Save());

            Assert.Null(exception1);
            Assert.Null(exception2);
        }

        [Fact]
        public void Photo_CreatedByLoad_IsNewInstance()
        {
            // Arrange & Act
            var photo1 = Photo.Load("test.jpg");
            var photo2 = Photo.Load("test.jpg");

            // Assert
            Assert.NotSame(photo1, photo2);
        }
    }
}
