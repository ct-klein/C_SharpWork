# Quick Start Guide - Automated Test Generation

## 🚀 Generate Tests in 3 Steps

### Step 1: Run the Generator

```bash
cd scripts/test-generator
node generate-tests.js ../../FindMissingAppointments
```

### Step 2: Copy the Generated Prompt

The script will output a prompt like this:

```
🚀 Generate comprehensive unit tests for the following C# files:

  - FindMissingAppointments/Helper.cs
  - FindMissingAppointments/Program.cs

📋 Requirements:
- Framework: xUnit
- Mocking: Moq
- Assertions: FluentAssertions
- Target Coverage: 80%
- Test Directory: tests/

🤖 Use the following agents IN PARALLEL (single message):
1. Task("Code Analyzer", "Analyze all source files...", "code-analyzer")
2. Task("Test Generator", "Generate comprehensive unit tests...", "tester")
3. Task("Test Reviewer", "Review generated tests...", "reviewer")

✅ Include ALL edge cases:
- Null parameters
- Empty strings
- Boundary values
- Exceptions
- Unicode & special characters
```

### Step 3: Paste into Claude Code

1. Open Claude Code
2. Paste the entire prompt
3. Claude Code will spawn all agents in parallel
4. Wait for tests to be generated
5. Review and run the tests!

## 📱 VS Code Integration

Press `Ctrl+Shift+P` (or `Cmd+Shift+P` on Mac) and search for:

- **"Tasks: Run Task"** → **"Generate Unit Tests (Current File)"**
- **"Tasks: Run Task"** → **"Generate Unit Tests (Current Project)"**
- **"Tasks: Run Task"** → **"Generate Unit Tests (Workspace)"**

## 🎯 Example: Test Your Helper.cs

```bash
# From test-generator directory
node generate-tests.js ../../FindMissingAppointments --file Helper.cs
```

This will generate a prompt specifically for Helper.cs with all edge cases!

## ✅ Verify Tests Work

```bash
# Run the generated tests
dotnet test tests/FindMissingAppointments.Tests/

# Run with coverage
dotnet test tests/FindMissingAppointments.Tests/ /p:CollectCoverage=true
```

## 🔧 Customize Configuration

Edit `config.json` to change:
- Test framework (xUnit, NUnit, MSTest)
- Coverage threshold (default: 80%)
- Edge case patterns
- Mocking library

## 💡 Pro Tips

1. **Generate Early**: Create tests when you write new classes
2. **Single File Mode**: Use `--file` flag for quick iteration
3. **Review Output**: AI-generated tests are great, but review for domain logic
4. **Re-run Anytime**: Safe to re-generate - creates new files without overwriting

## 🐛 Troubleshooting

**Issue**: Script doesn't find files
**Fix**: Check your path - should point to project root

**Issue**: Tests don't compile
**Fix**: Ensure NuGet packages are restored with `dotnet restore`

**Issue**: Claude Flow not working
**Fix**: Install with `npm install -g claude-flow@alpha`

---

**Need Help?** Check the full [README.md](README.md) for detailed documentation.
