using System;
using System.Reflection;
using FluentAssertions;
using Xunit;

namespace Helper.Tests
{
    /// <summary>
    /// Unit tests for CrmServiceHelper class
    /// Tests cover initialization, validation, and service client creation scenarios
    /// </summary>
    public class CrmServiceHelperTests : IDisposable
    {
        // Test constants
        private const string ValidUrl = "https://test.crm.dynamics.com";
        private const string ValidClientId = "test-client-id-12345";
        private const string ValidClientSecret = "test-client-secret-67890";

        public CrmServiceHelperTests()
        {
            // Reset static fields before each test
            ResetStaticFields();
        }

        public void Dispose()
        {
            // Cleanup after each test
            ResetStaticFields();
        }

        #region Helper Methods

        /// <summary>
        /// Resets the static fields in CrmServiceHelper using reflection
        /// This is necessary because the class uses static fields
        /// </summary>
        private void ResetStaticFields()
        {
            var type = typeof(CrmServiceHelper);
            var crmUrlField = type.GetField("crmUrl", BindingFlags.NonPublic | BindingFlags.Static);
            var clientIdField = type.GetField("clientId", BindingFlags.NonPublic | BindingFlags.Static);
            var clientSecretField = type.GetField("clientSecret", BindingFlags.NonPublic | BindingFlags.Static);

            crmUrlField?.SetValue(null, null);
            clientIdField?.SetValue(null, null);
            clientSecretField?.SetValue(null, null);
        }

        /// <summary>
        /// Gets the value of a private static field using reflection
        /// </summary>
        private string GetStaticFieldValue(string fieldName)
        {
            var type = typeof(CrmServiceHelper);
            var field = type.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Static);
            return field?.GetValue(null) as string;
        }

        /// <summary>
        /// Invokes the private ValidateConnectionComponents method using reflection
        /// </summary>
        private void InvokeValidateConnectionComponents()
        {
            var type = typeof(CrmServiceHelper);
            var method = type.GetMethod("ValidateConnectionComponents", BindingFlags.NonPublic | BindingFlags.Static);
            method?.Invoke(null, null);
        }

        #endregion

        #region Initialize Tests

        [Fact]
        public void Initialize_WithValidParameters_ShouldSetAllFields()
        {
            // Arrange & Act
            CrmServiceHelper.Initialize(ValidUrl, ValidClientId, ValidClientSecret);

            // Assert
            GetStaticFieldValue("crmUrl").Should().Be(ValidUrl);
            GetStaticFieldValue("clientId").Should().Be(ValidClientId);
            GetStaticFieldValue("clientSecret").Should().Be(ValidClientSecret);
        }

        [Fact]
        public void Initialize_WithEmptyUrl_ShouldSetEmptyUrl()
        {
            // Arrange & Act
            CrmServiceHelper.Initialize(string.Empty, ValidClientId, ValidClientSecret);

            // Assert
            GetStaticFieldValue("crmUrl").Should().BeEmpty();
            GetStaticFieldValue("clientId").Should().Be(ValidClientId);
            GetStaticFieldValue("clientSecret").Should().Be(ValidClientSecret);
        }

        [Fact]
        public void Initialize_WithNullParameters_ShouldSetNullValues()
        {
            // Arrange & Act
            CrmServiceHelper.Initialize(null, null, null);

            // Assert
            GetStaticFieldValue("crmUrl").Should().BeNull();
            GetStaticFieldValue("clientId").Should().BeNull();
            GetStaticFieldValue("clientSecret").Should().BeNull();
        }

        [Theory]
        [InlineData("https://prod.crm.dynamics.com", "prod-client-id", "prod-secret")]
        [InlineData("https://dev.crm.dynamics.com", "dev-client-id", "dev-secret")]
        [InlineData("https://test.crm.dynamics.com", "test-client-id", "test-secret")]
        public void Initialize_WithDifferentValidValues_ShouldSetCorrectly(string url, string clientId, string clientSecret)
        {
            // Arrange & Act
            CrmServiceHelper.Initialize(url, clientId, clientSecret);

            // Assert
            GetStaticFieldValue("crmUrl").Should().Be(url);
            GetStaticFieldValue("clientId").Should().Be(clientId);
            GetStaticFieldValue("clientSecret").Should().Be(clientSecret);
        }

        [Fact]
        public void Initialize_CalledMultipleTimes_ShouldOverwritePreviousValues()
        {
            // Arrange
            CrmServiceHelper.Initialize("first-url", "first-id", "first-secret");

            // Act
            CrmServiceHelper.Initialize(ValidUrl, ValidClientId, ValidClientSecret);

            // Assert
            GetStaticFieldValue("crmUrl").Should().Be(ValidUrl);
            GetStaticFieldValue("clientId").Should().Be(ValidClientId);
            GetStaticFieldValue("clientSecret").Should().Be(ValidClientSecret);
        }

        #endregion

        #region ValidateConnectionComponents Tests

        [Fact]
        public void ValidateConnectionComponents_WithAllValidParameters_ShouldNotThrow()
        {
            // Arrange
            CrmServiceHelper.Initialize(ValidUrl, ValidClientId, ValidClientSecret);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().NotThrow();
        }

        [Fact]
        public void ValidateConnectionComponents_WithNullUrl_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(null, ValidClientId, ValidClientSecret);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().Throw<TargetInvocationException>()
                .WithInnerException<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Fact]
        public void ValidateConnectionComponents_WithEmptyUrl_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(string.Empty, ValidClientId, ValidClientSecret);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().Throw<TargetInvocationException>()
                .WithInnerException<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Fact]
        public void ValidateConnectionComponents_WithNullClientId_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(ValidUrl, null, ValidClientSecret);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().Throw<TargetInvocationException>()
                .WithInnerException<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Fact]
        public void ValidateConnectionComponents_WithEmptyClientId_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(ValidUrl, string.Empty, ValidClientSecret);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().Throw<TargetInvocationException>()
                .WithInnerException<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Fact]
        public void ValidateConnectionComponents_WithNullClientSecret_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(ValidUrl, ValidClientId, null);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().Throw<TargetInvocationException>()
                .WithInnerException<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Fact]
        public void ValidateConnectionComponents_WithEmptyClientSecret_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(ValidUrl, ValidClientId, string.Empty);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().Throw<TargetInvocationException>()
                .WithInnerException<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Fact]
        public void ValidateConnectionComponents_WithAllNullParameters_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(null, null, null);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().Throw<TargetInvocationException>()
                .WithInnerException<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Fact]
        public void ValidateConnectionComponents_WithAllEmptyParameters_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(string.Empty, string.Empty, string.Empty);

            // Act
            Action act = () => InvokeValidateConnectionComponents();

            // Assert
            act.Should().Throw<TargetInvocationException>()
                .WithInnerException<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        #endregion

        #region GetCrmServiceClient Tests

        [Fact]
        public void GetCrmServiceClient_WithoutInitialization_ShouldThrowArgumentException()
        {
            // Arrange - Don't call Initialize

            // Act
            Action act = () => CrmServiceHelper.GetCrmServiceClient();

            // Assert
            act.Should().Throw<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Fact]
        public void GetCrmServiceClient_WithPartialInitialization_ShouldThrowArgumentException()
        {
            // Arrange
            CrmServiceHelper.Initialize(ValidUrl, null, ValidClientSecret);

            // Act
            Action act = () => CrmServiceHelper.GetCrmServiceClient();

            // Assert
            act.Should().Throw<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        [Theory]
        [InlineData(null, "client-id", "client-secret")]
        [InlineData("url", null, "client-secret")]
        [InlineData("url", "client-id", null)]
        [InlineData("", "client-id", "client-secret")]
        [InlineData("url", "", "client-secret")]
        [InlineData("url", "client-id", "")]
        public void GetCrmServiceClient_WithInvalidParameters_ShouldThrowArgumentException(
            string url, string clientId, string clientSecret)
        {
            // Arrange
            CrmServiceHelper.Initialize(url, clientId, clientSecret);

            // Act
            Action act = () => CrmServiceHelper.GetCrmServiceClient();

            // Assert
            act.Should().Throw<ArgumentException>()
                .WithMessage("*CrmServiceClient connection components cannot be null or empty*");
        }

        #endregion

        #region Edge Cases and Security Tests

        [Fact]
        public void Initialize_WithWhitespaceUrl_ShouldSetWhitespaceValue()
        {
            // Arrange
            var whitespaceUrl = "   ";

            // Act
            CrmServiceHelper.Initialize(whitespaceUrl, ValidClientId, ValidClientSecret);

            // Assert
            GetStaticFieldValue("crmUrl").Should().Be(whitespaceUrl);
        }

        [Fact]
        public void Initialize_WithVeryLongStrings_ShouldHandleCorrectly()
        {
            // Arrange
            var longUrl = new string('a', 10000);
            var longClientId = new string('b', 10000);
            var longClientSecret = new string('c', 10000);

            // Act
            CrmServiceHelper.Initialize(longUrl, longClientId, longClientSecret);

            // Assert
            GetStaticFieldValue("crmUrl").Should().HaveLength(10000);
            GetStaticFieldValue("clientId").Should().HaveLength(10000);
            GetStaticFieldValue("clientSecret").Should().HaveLength(10000);
        }

        [Fact]
        public void Initialize_WithSpecialCharacters_ShouldHandleCorrectly()
        {
            // Arrange
            var specialUrl = "https://test.crm.dynamics.com?param=value&special=<>\"'";
            var specialClientId = "client-id-!@#$%^&*()";
            var specialClientSecret = "secret-{}[]|\\;:,.<>?/~`";

            // Act
            CrmServiceHelper.Initialize(specialUrl, specialClientId, specialClientSecret);

            // Assert
            GetStaticFieldValue("crmUrl").Should().Be(specialUrl);
            GetStaticFieldValue("clientId").Should().Be(specialClientId);
            GetStaticFieldValue("clientSecret").Should().Be(specialClientSecret);
        }

        [Fact]
        public void Initialize_WithUnicodeCharacters_ShouldHandleCorrectly()
        {
            // Arrange
            var unicodeUrl = "https://test.crm.dynamics.com/测试";
            var unicodeClientId = "client-id-αβγδε";
            var unicodeClientSecret = "secret-日本語";

            // Act
            CrmServiceHelper.Initialize(unicodeUrl, unicodeClientId, unicodeClientSecret);

            // Assert
            GetStaticFieldValue("crmUrl").Should().Be(unicodeUrl);
            GetStaticFieldValue("clientId").Should().Be(unicodeClientId);
            GetStaticFieldValue("clientSecret").Should().Be(unicodeClientSecret);
        }

        #endregion
    }
}
