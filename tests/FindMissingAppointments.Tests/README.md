# FindMissingAppointments.Tests

Unit tests for the FindMissingAppointments Helper library.

## Test Framework

- **xUnit**: Primary testing framework
- **FluentAssertions**: For readable assertions
- **Moq**: For mocking dependencies (if needed for future tests)

## Running Tests

### Command Line
```bash
# Run all tests
dotnet test

# Run with detailed output
dotnet test --verbosity detailed

# Run with code coverage
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
```

### Visual Studio
1. Open Test Explorer (Test > Test Explorer)
2. Click "Run All" to execute all tests
3. View test results in the Test Explorer window

## Test Coverage

### CrmServiceHelperTests
Tests for the `CrmServiceHelper` class covering:

#### Initialize Method
- ✅ Valid parameter initialization
- ✅ Empty parameter handling
- ✅ Null parameter handling
- ✅ Multiple initialization calls (overwriting)
- ✅ Various valid configurations

#### ValidateConnectionComponents Method
- ✅ Valid parameters pass validation
- ✅ Null URL throws ArgumentException
- ✅ Empty URL throws ArgumentException
- ✅ Null ClientId throws ArgumentException
- ✅ Empty ClientId throws ArgumentException
- ✅ Null ClientSecret throws ArgumentException
- ✅ Empty ClientSecret throws ArgumentException
- ✅ All null parameters throw ArgumentException
- ✅ All empty parameters throw ArgumentException

#### GetCrmServiceClient Method
- ✅ Uninitialized state throws ArgumentException
- ✅ Partial initialization throws ArgumentException
- ✅ Various invalid parameter combinations throw ArgumentException

#### Edge Cases & Security
- ✅ Whitespace handling
- ✅ Very long strings (10,000 characters)
- ✅ Special characters handling
- ✅ Unicode characters handling

## Test Structure

Each test follows the AAA pattern:
- **Arrange**: Set up test data and prerequisites
- **Act**: Execute the method being tested
- **Assert**: Verify the expected outcome

## Notes

⚠️ **Static Fields**: The `CrmServiceHelper` class uses static fields, so tests include setup and teardown logic to reset state between tests.

⚠️ **Reflection Usage**: Some tests use reflection to access private methods and fields for thorough testing. This is intentional for unit testing purposes.

⚠️ **Integration Tests**: Tests for `GetCrmServiceClient` that require actual CRM connection are not included as they would require live credentials and network access. These should be implemented as integration tests in a separate project.

## Future Enhancements

1. Add integration tests for actual CRM connections
2. Add tests for retry logic and timeout scenarios
3. Add performance tests for connection establishment
4. Mock `CrmServiceClient` for more comprehensive unit testing
