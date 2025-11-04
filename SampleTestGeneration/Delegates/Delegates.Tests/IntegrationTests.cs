using Xunit;
using Delegates;
using System;
using System.IO;

namespace Delegates.Tests
{
    [Collection("Console Tests")]
    public class IntegrationTests
    {
        [Fact]
        public void FullWorkflow_SimulatingProgramMain_ExecutesCorrectly()
        {
            // Arrange - Simulating Program.Main
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += filters.ApplyContrast;
            filterHandler += RemoveRedEyeFilter;

            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            processor.Process("photo.jpg", filterHandler);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);
            Assert.Contains("Apply RemoveRedEye", output);

            // Verify execution order matches Program.Main
            int brightnessIndex = output.IndexOf("Apply brightness");
            int contrastIndex = output.IndexOf("Apply contrast");
            int redEyeIndex = output.IndexOf("Apply RemoveRedEye");

            Assert.True(brightnessIndex >= 0);
            Assert.True(contrastIndex > brightnessIndex);
            Assert.True(redEyeIndex > contrastIndex);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void FullWorkflow_WithDifferentFilterCombination_ExecutesCorrectly()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.Resize;
            filterHandler += filters.ApplyBrightness;
            filterHandler += filters.ApplyContrast;

            var stringWriter = new StringWriter();
            var originalOut = Console.Out;
            Console.SetOut(stringWriter);

            // Act
            processor.Process("photo.jpg", filterHandler);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Resize photo", output);
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);

            // Cleanup
            Console.SetOut(originalOut);
        }

        [Fact]
        public void FullWorkflow_ProcessingMultiplePhotos_EachPhotoProcessedIndependently()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;

            var stringWriter = new StringWriter();
            var originalOut = Console.Out;
            Console.SetOut(stringWriter);

            // Act
            processor.Process("photo1.jpg", filterHandler);
            processor.Process("photo2.jpg", filterHandler);
            processor.Process("photo3.jpg", filterHandler);

            // Assert
            var output = stringWriter.ToString();

            // Count occurrences using Split
            var occurrences = output.Split(new[] { "Apply brightness" }, StringSplitOptions.None).Length - 1;
            Assert.Equal(3, occurrences);

            // Cleanup
            Console.SetOut(originalOut);
        }

        [Fact]
        public void FullWorkflow_WithAllAvailableFilters_ExecutesAll()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += filters.ApplyContrast;
            filterHandler += filters.Resize;
            filterHandler += RemoveRedEyeFilter;

            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            processor.Process("photo.jpg", filterHandler);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);
            Assert.Contains("Resize photo", output);
            Assert.Contains("Apply RemoveRedEye", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void FullWorkflow_CreateAndUseMultipleProcessors_AllWorkIndependently()
        {
            // Arrange
            var processor1 = new PhotoProcessor();
            var processor2 = new PhotoProcessor();
            var filters = new PhotoFilters();

            Action<Photo> brightFilter = filters.ApplyBrightness;
            Action<Photo> contrastFilter = filters.ApplyContrast;

            int processor1Executions = 0;
            int processor2Executions = 0;

            Action<Photo> countingFilter1 = (p) => processor1Executions++;
            Action<Photo> countingFilter2 = (p) => processor2Executions++;

            // Act
            processor1.Process("photo1.jpg", brightFilter + countingFilter1);
            processor2.Process("photo2.jpg", contrastFilter + countingFilter2);

            // Assert
            Assert.Equal(1, processor1Executions);
            Assert.Equal(1, processor2Executions);
        }

        [Fact]
        public void FullWorkflow_WithCustomFiltersOnly_ExecutesWithoutBuiltInFilters()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var customFilterExecuted = false;

            Action<Photo> customFilter = (photo) =>
            {
                customFilterExecuted = true;
                Console.WriteLine("Custom filter applied");
            };

            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            processor.Process("photo.jpg", customFilter);

            // Assert
            Assert.True(customFilterExecuted);
            var output = stringWriter.ToString();
            Assert.Contains("Custom filter applied", output);
            Assert.DoesNotContain("Apply brightness", output);
            Assert.DoesNotContain("Apply contrast", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void FullWorkflow_DynamicallyBuildFilterChain_ExecutesAllFilters()
        {
            // Arrange
            var processor = new PhotoProcessor();
            var filters = new PhotoFilters();
            Action<Photo>? filterHandler = null;

            // Dynamically build filter chain
            bool applyBrightness = true;
            bool applyContrast = true;
            bool resize = false;

            if (applyBrightness)
                filterHandler += filters.ApplyBrightness;

            if (applyContrast)
                filterHandler += filters.ApplyContrast;

            if (resize)
                filterHandler += filters.Resize;

            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            processor.Process("photo.jpg", filterHandler!);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);
            Assert.DoesNotContain("Resize photo", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        // Helper method simulating the static RemoveRedEyeFilter from Program.cs
        private static void RemoveRedEyeFilter(Photo photo)
        {
            Console.WriteLine("Apply RemoveRedEye");
        }
    }
}
