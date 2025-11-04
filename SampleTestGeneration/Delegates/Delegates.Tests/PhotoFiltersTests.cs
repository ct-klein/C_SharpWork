using Xunit;
using Delegates;
using System;
using System.IO;

namespace Delegates.Tests
{
    [Collection("Console Tests")]
    public class PhotoFiltersTests
    {
        [Fact]
        public void ApplyBrightness_WithValidPhoto_WritesToConsole()
        {
            // Arrange
            var filters = new PhotoFilters();
            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filters.ApplyBrightness(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void ApplyContrast_WithValidPhoto_WritesToConsole()
        {
            // Arrange
            var filters = new PhotoFilters();
            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filters.ApplyContrast(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply contrast", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void Resize_WithValidPhoto_WritesToConsole()
        {
            // Arrange
            var filters = new PhotoFilters();
            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filters.Resize(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Resize photo", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void ApplyBrightness_WithNullPhoto_DoesNotThrowException()
        {
            // Arrange
            var filters = new PhotoFilters();
            Photo? photo = null;

            // Act & Assert
            var exception = Record.Exception(() => filters.ApplyBrightness(photo!));
            Assert.Null(exception);
        }

        [Fact]
        public void ApplyContrast_WithNullPhoto_DoesNotThrowException()
        {
            // Arrange
            var filters = new PhotoFilters();
            Photo? photo = null;

            // Act & Assert
            var exception = Record.Exception(() => filters.ApplyContrast(photo!));
            Assert.Null(exception);
        }

        [Fact]
        public void Resize_WithNullPhoto_DoesNotThrowException()
        {
            // Arrange
            var filters = new PhotoFilters();
            Photo? photo = null;

            // Act & Assert
            var exception = Record.Exception(() => filters.Resize(photo!));
            Assert.Null(exception);
        }

        [Fact]
        public void ApplyMultipleFilters_InSequence_AllExecuteSuccessfully()
        {
            // Arrange
            var filters = new PhotoFilters();
            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filters.ApplyBrightness(photo);
            filters.ApplyContrast(photo);
            filters.Resize(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);
            Assert.Contains("Resize photo", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void PhotoFilters_CanBeInstantiatedMultipleTimes()
        {
            // Arrange & Act
            var filters1 = new PhotoFilters();
            var filters2 = new PhotoFilters();

            // Assert
            Assert.NotNull(filters1);
            Assert.NotNull(filters2);
            Assert.NotSame(filters1, filters2);
        }

        [Fact]
        public void ApplyBrightness_CalledMultipleTimes_WritesCorrectOutput()
        {
            // Arrange
            var filters = new PhotoFilters();
            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            var originalOut = Console.Out;
            Console.SetOut(stringWriter);

            // Act
            filters.ApplyBrightness(photo);
            filters.ApplyBrightness(photo);

            // Assert
            var output = stringWriter.ToString();

            // Count occurrences using Split
            var occurrences = output.Split(new[] { "Apply brightness" }, StringSplitOptions.None).Length - 1;
            Assert.Equal(2, occurrences);

            // Cleanup
            Console.SetOut(originalOut);
        }
    }
}
