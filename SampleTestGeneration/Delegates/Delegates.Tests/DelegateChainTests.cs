using Xunit;
using Delegates;
using System;
using System.IO;

namespace Delegates.Tests
{
    [Collection("Console Tests")]
    public class DelegateChainTests
    {
        [Fact]
        public void DelegateChain_WithThreeFilters_ExecutesInCorrectOrder()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += filters.ApplyContrast;
            filterHandler += (photo) => Console.WriteLine("Apply RemoveRedEye");

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filterHandler(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);
            Assert.Contains("Apply RemoveRedEye", output);

            // Verify execution order
            int brightnessIndex = output.IndexOf("Apply brightness");
            int contrastIndex = output.IndexOf("Apply contrast");
            int redEyeIndex = output.IndexOf("Apply RemoveRedEye");

            Assert.True(brightnessIndex < contrastIndex);
            Assert.True(contrastIndex < redEyeIndex);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void DelegateChain_AddedUsingPlusOperator_ExecutesAllDelegates()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo> filter1 = filters.ApplyBrightness;
            Action<Photo> filter2 = filters.ApplyContrast;
            Action<Photo> combinedFilters = filter1 + filter2;

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            combinedFilters(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void DelegateChain_AddedUsingPlusEqualsOperator_ExecutesAllDelegates()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += filters.ApplyContrast;

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filterHandler(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void DelegateChain_RemoveDelegate_ExecutesRemainingDelegates()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += filters.ApplyContrast;
            filterHandler += filters.Resize;
            filterHandler -= filters.ApplyContrast; // Remove contrast

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filterHandler(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.DoesNotContain("Apply contrast", output);
            Assert.Contains("Resize photo", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void DelegateChain_WithSingleDelegate_ExecutesOnce()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            var originalOut = Console.Out;
            Console.SetOut(stringWriter);

            // Act
            filterHandler(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);

            // Count occurrences using Split
            var occurrences = output.Split(new[] { "Apply brightness" }, StringSplitOptions.None).Length - 1;
            Assert.Equal(1, occurrences);

            // Cleanup
            Console.SetOut(originalOut);
        }

        [Fact]
        public void DelegateChain_WithMixedInstanceAndStaticMethods_ExecutesAll()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += RemoveRedEyeFilter;

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filterHandler(photo);

            // Assert
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply RemoveRedEye", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void DelegateChain_AddSameMethodTwice_ExecutesBothTimes()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += filters.ApplyBrightness; // Add same method twice

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            var originalOut = Console.Out;
            Console.SetOut(stringWriter);

            // Act
            filterHandler(photo);

            // Assert
            var output = stringWriter.ToString();

            // Count occurrences using Split
            var occurrences = output.Split(new[] { "Apply brightness" }, StringSplitOptions.None).Length - 1;
            Assert.Equal(2, occurrences);

            // Cleanup
            Console.SetOut(originalOut);
        }

        [Fact]
        public void DelegateChain_CombineMultipleChainsWithPlusOperator_ExecutesAllDelegates()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo> chain1 = filters.ApplyBrightness;
            chain1 += filters.ApplyContrast;

            Action<Photo> chain2 = filters.Resize;
            chain2 += RemoveRedEyeFilter;

            Action<Photo> combinedChain = chain1 + chain2;

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            combinedChain(photo);

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
        public void DelegateChain_WithLambdaExpressions_ExecutesAllInChain()
        {
            // Arrange
            var filters = new PhotoFilters();
            int lambdaExecutionCount = 0;

            Action<Photo> filterHandler = filters.ApplyBrightness;
            filterHandler += (photo) => lambdaExecutionCount++;
            filterHandler += filters.ApplyContrast;
            filterHandler += (photo) => lambdaExecutionCount++;

            var photo = Photo.Load("test.jpg");
            var stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            // Act
            filterHandler(photo);

            // Assert
            Assert.Equal(2, lambdaExecutionCount);
            var output = stringWriter.ToString();
            Assert.Contains("Apply brightness", output);
            Assert.Contains("Apply contrast", output);

            // Cleanup
            Console.SetOut(Console.Out);
        }

        [Fact]
        public void DelegateChain_EmptyChainAfterRemovingAllDelegates_IsNull()
        {
            // Arrange
            var filters = new PhotoFilters();
            Action<Photo>? filterHandler = filters.ApplyBrightness;
            filterHandler -= filters.ApplyBrightness;

            // Assert
            Assert.Null(filterHandler);
        }

        // Helper method simulating the static RemoveRedEyeFilter from Program.cs
        private static void RemoveRedEyeFilter(Photo photo)
        {
            Console.WriteLine("Apply RemoveRedEye");
        }
    }
}
