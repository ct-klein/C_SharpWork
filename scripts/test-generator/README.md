# C# Automated Test Generator

An intelligent test generation system that uses AI agents to automatically create comprehensive unit tests for your C# projects.

## 🌟 Features

- **Automatic Code Analysis**: Analyzes your C# code to understand structure, dependencies, and patterns
- **Comprehensive Test Generation**: Creates thorough unit tests with edge cases
- **Multi-Agent Architecture**: Uses specialized agents (analyzer, generator, reviewer)
- **Configurable**: Customize test frameworks, patterns, and coverage targets
- **Edge Case Detection**: Automatically includes null checks, boundary values, special characters
- **Framework Support**: xUnit, NUnit, MSTest
- **Mocking Support**: Moq, NSubstitute, FakeItEasy

## 🚀 Quick Start

### 1. Generate Tests for Your Project

```bash
# Navigate to the test generator directory
cd scripts/test-generator

# Generate tests for entire project
node generate-tests.js ../../MyProject

# Generate tests for specific file
node generate-tests.js ../../MyProject --file Helper.cs

# Show help
node generate-tests.js --help
```

### 2. Follow the Generated Instructions

The script will provide you with a complete prompt to paste into Claude Code, which will spawn all agents in parallel to analyze and generate tests.

## 📋 What Gets Generated

For each C# source file, the system creates:

### ✅ Test Project Structure
```
tests/
└── YourProject.Tests/
    ├── YourProject.Tests.csproj
    ├── ClassNameTests.cs
    ├── AnotherClassTests.cs
    └── README.md
```

### ✅ Test Coverage Includes

**Method Testing:**
- All public methods
- Private methods (using reflection)
- Static methods
- Async methods
- Property getters/setters
- Constructors

**Edge Cases:**
- ✅ Null parameters
- ✅ Empty/whitespace strings
- ✅ Boundary values (min/max)
- ✅ Exception scenarios
- ✅ Unicode characters
- ✅ Special characters
- ✅ Very long strings
- ✅ Concurrent access (if applicable)

**Test Patterns:**
- AAA pattern (Arrange, Act, Assert)
- Parameterized tests with [Theory]
- Single case tests with [Fact]
- Proper mocking of dependencies
- Fluent assertions
- Comprehensive documentation

## ⚙️ Configuration

Edit `config.json` to customize:

```json
{
  "testGenerator": {
    "framework": "xUnit",              // xUnit, NUnit, MSTest
    "mockingLibrary": "Moq",           // Moq, NSubstitute, FakeItEasy
    "assertionLibrary": "FluentAssertions",
    "targetFramework": "net6.0",
    "coverageThreshold": 80,
    "testPatterns": {
      "publicMethods": { "enabled": true },
      "privateMethods": { "enabled": true, "useReflection": true },
      "edgeCases": {
        "nullParameters": true,
        "emptyStrings": true,
        "boundaryValues": true,
        "exceptions": true,
        "unicodeAndSpecialChars": true
      }
    }
  }
}
```

## 🤖 How It Works

### Multi-Agent Architecture

The system uses three specialized agents that work concurrently:

#### 1. **Code Analyzer Agent** (`code-analyzer`)
- Parses C# source files
- Extracts classes, methods, properties
- Identifies dependencies and patterns
- Detects complexity and test requirements
- Stores analysis in shared memory

#### 2. **Test Generator Agent** (`tester`)
- Retrieves code analysis from memory
- Creates test project structure
- Generates comprehensive test classes
- Implements edge case tests
- Creates mocks for dependencies
- Stores test metadata in memory

#### 3. **Test Reviewer Agent** (`reviewer`)
- Reviews generated tests
- Validates coverage targets
- Checks for missing edge cases
- Ensures best practices
- Suggests improvements
- Produces quality report

### Agent Coordination

```javascript
// All agents run in parallel (single message)
Task("Code Analyzer", "Analyze Helper.cs...", "code-analyzer")
Task("Test Generator", "Generate tests with 80% coverage...", "tester")
Task("Test Reviewer", "Review test quality and completeness...", "reviewer")
```

Agents coordinate via:
- **Shared Memory**: Analysis and metadata stored/retrieved
- **Hooks**: Session management and progress tracking
- **Results**: Each agent produces output for the next

## 📊 Usage Examples

### Example 1: Generate Tests for Single File

```bash
node generate-tests.js ../FindMissingAppointments --file Helper.cs
```

**Output:**
- Analyzes `Helper.cs`
- Generates `tests/FindMissingAppointments.Tests/HelperTests.cs`
- Creates test project if needed
- Provides execution instructions

### Example 2: Generate Tests for Entire Project

```bash
node generate-tests.js ../MyEntireProject
```

**Output:**
- Scans all `.cs` files
- Generates tests for each file
- Creates comprehensive test suite
- Generates documentation

### Example 3: Integration with CI/CD

```bash
# Add to your build pipeline
npm run generate -- ./src --watch
dotnet test ./tests --collect:"XPlat Code Coverage"
```

## 🎯 Best Practices

### For Best Results:

1. **Run Early**: Generate tests when creating new classes
2. **Review Generated Tests**: AI does great work, but review for domain logic
3. **Customize Config**: Adjust coverage targets and patterns for your needs
4. **Integrate with CI**: Run test generation in pre-commit hooks
5. **Update Tests**: Re-generate when code changes significantly

### Naming Conventions:

The generator follows C# testing conventions:
- Test class: `{ClassName}Tests`
- Test method: `{MethodName}_{Scenario}_{ExpectedBehavior}`
- Example: `Initialize_WithNullUrl_ShouldThrowArgumentException`

## 🔧 Advanced Features

### Watch Mode (Coming Soon)

```bash
node generate-tests.js ./MyProject --watch
```

Automatically generates tests when source files change.

### Custom Templates

Create custom test templates in `templates/`:
- `templates/xunit-template.cs`
- `templates/nunit-template.cs`
- `templates/mstest-template.cs`

### VS Code Integration (Optional)

Add to `.vscode/tasks.json`:

```json
{
  "label": "Generate Unit Tests",
  "type": "shell",
  "command": "node scripts/test-generator/generate-tests.js ${fileDirname}",
  "problemMatcher": []
}
```

## 📈 Benefits

### Time Savings
- **Manual**: 30-60 minutes per class
- **Automated**: 2-5 minutes per class
- **Savings**: 90%+ time reduction

### Quality Improvements
- Consistent test patterns
- Comprehensive edge case coverage
- Best practice enforcement
- Reduced human error

### Coverage Goals
- Target: 80%+ code coverage
- Includes: All public APIs
- Validates: Exception handling
- Tests: Edge cases automatically

## 🛠️ Troubleshooting

### Issue: "No C# files found"
**Solution**: Ensure you're pointing to the correct project directory

### Issue: "Cannot create test directory"
**Solution**: Check write permissions in the target directory

### Issue: "Agent coordination failed"
**Solution**: Ensure Claude Flow is installed: `npm install -g claude-flow@alpha`

### Issue: "Tests don't compile"
**Solution**: Review generated code and adjust namespaces/dependencies

## 🔗 Integration with Claude Code

This system is designed to work seamlessly with Claude Code:

1. Run the generator script
2. Copy the generated prompt
3. Paste into Claude Code
4. Agents spawn in parallel and generate tests
5. Review and commit the generated tests

## 📚 Additional Resources

- [xUnit Documentation](https://xunit.net/)
- [Moq Documentation](https://github.com/moq/moq4)
- [FluentAssertions Documentation](https://fluentassertions.com/)
- [Claude Flow Documentation](https://github.com/ruvnet/claude-flow)

## 🤝 Contributing

To extend the test generator:

1. Add new test patterns in `config.json`
2. Create custom agent prompts in `generate-tests.js`
3. Add templates for different frameworks
4. Submit improvements!

## 📝 License

MIT License - Feel free to use and modify for your projects!

---

**Happy Testing! 🧪✨**

*Generated tests are a great starting point, but always review for domain-specific logic and business rules.*
