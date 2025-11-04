using Xunit;
using Delegates;
using System;
using System.IO;

namespace Delegates.Tests
{
    [Collection("Console Tests")]
    public class PhotoProcessorTests
    {
        [Fact]
        public void Process_WithSingleFilter_ExecutesFilterOnce()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            processor.Process("test.jpg", filterHandler);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void Process_WithMultipleFilters_ExecutesAllFiltersInOrder()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += filters.ApplyContrast;
            filterHandler += filters.Resize;
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            processor.Process("test.jpg", filterHandler);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);
            Assert.Contains("Resize photo", output);

            // Verify order
            int brightnessIndex = output.IndexOf("Apply brightness");
            int contrastIndex = output.IndexOf("Apply contrast");
            int resizeIndex = output.IndexOf("Resize photo");

            Assert.True(brightnessIndex < contrastIndex);
            Assert.True(contrastIndex < resizeIndex);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void Process_WithEmptyPath_ExecutesSuccessfully()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;

            // Act & Assert
            var exception = Record.Exception(() => processor.Process(string.Empty, filterHandler));
            Assert.Null(exception);
        }

        [Fact]
        public void Process_WithNullPath_ExecutesSuccessfully()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;

            // Act & Assert
            var exception = Record.Exception(() => processor.Process(null!, filterHandler));
            Assert.Null(exception);
        }

        [Fact]
        public void Process_WithCustomDelegate_ExecutesCustomAction()
        {
            // Arrange
            var processor = new PhotoProcessor();
            bool customActionExecuted = false;
            Action<Photo> customFilter = (photo) => { customActionExecuted = true; };

            // Act
            processor.Process("test.jpg", customFilter);

            // Assert
            Assert.True(customActionExecuted);
        }

        [Fact]
        public void Process_WithLambdaExpression_ExecutesLambda()
        {
            // Arrange
            var processor = new PhotoProcessor();
            int executionCount = 0;
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            processor.Process("test.jpg", (photo) =>
            {
                executionCount++;
                Console.WriteLine("Lambda executed");
            });

            // Assert
            Assert.Equal(1, executionCount);
            var output = stringWriter.ToString();
            Assert.Contains("Lambda executed", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void Process_WithChainedDelegates_ExecutesAllDelegates()
        {
            // Arrange
            var processor = new PhotoProcessor();
            int counter = 0;
            Action<Photo> filter1 = (p) => counter++;
            Action<Photo> filter2 = (p) => counter++;
            Action<Photo> filter3 = (p) => counter++;
            Action<Photo> chainedFilters = filter1 + filter2 + filter3;

            // Act
            processor.Process("test.jpg", chainedFilters);

            // Assert
            Assert.Equal(3, counter);
        }

        [Fact]
        public void Process_LoadsAndSavesPhoto()
        {
            // Arrange
            var processor = new PhotoProcessor();
            Photo? capturedPhoto = null;
            Action<Photo> captureFilter = (photo) => capturedPhoto = photo;

            // Act
            processor.Process("test.jpg", captureFilter);

            // Assert
            Assert.NotNull(capturedPhoto);
            Assert.IsType<Photo>(capturedPhoto);
        }

        [Fact]
        public void Process_CalledMultipleTimes_ProcessesEachRequest()
        {
            // Arrange
            var processor = new PhotoProcessor();
            int executionCount = 0;
            Action<Photo> countingFilter = (p) => executionCount++;

            // Act
            processor.Process("test1.jpg", countingFilter);
            processor.Process("test2.jpg", countingFilter);
            processor.Process("test3.jpg", countingFilter);

            // Assert
            Assert.Equal(3, executionCount);
        }

        [Fact]
        public void Process_WithMixedFilterTypes_ExecutesAll()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += filters.ApplyContrast;
            filterHandler += (photo) => Console.WriteLine("Custom lambda filter");

            // Act
            processor.Process("test.jpg", filterHandler);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);
            Assert.Contains("Custom lambda filter", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }
    }
}
