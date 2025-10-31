#!/usr/bin/env node

/**
 * Automated Unit Test Generator for C# Projects
 * Uses Claude Flow agents to analyze code and generate comprehensive tests
 *
 * Usage:
 *   node generate-tests.js <project-path>
 *   node generate-tests.js <project-path> --file <specific-file.cs>
 *   node generate-tests.js <project-path> --watch (continuous mode)
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Configuration
const CONFIG_FILE = path.join(__dirname, 'config.json');
const config = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8')).testGenerator;

class TestGeneratorOrchestrator {
  constructor(projectPath, options = {}) {
    this.projectPath = path.resolve(projectPath);
    this.options = options;
    this.config = config;
    this.results = {
      analyzed: [],
      generated: [],
      errors: []
    };
  }

  /**
   * Main orchestration method - spawns all agents concurrently
   */
  async generate() {
    console.log('🚀 Starting Automated Test Generation...\n');
    console.log(`📁 Project Path: ${this.projectPath}`);
    console.log(`⚙️  Framework: ${this.config.framework}`);
    console.log(`📊 Target Coverage: ${this.config.coverageThreshold}%\n`);

    try {
      // Step 1: Discover C# files
      const sourceFiles = this.discoverSourceFiles();
      console.log(`✅ Found ${sourceFiles.length} C# file(s) to analyze\n`);

      if (sourceFiles.length === 0) {
        console.log('⚠️  No C# files found. Exiting.');
        return;
      }

      // Step 2: Generate Claude Flow agent tasks
      const agentTasks = this.buildAgentTasks(sourceFiles);

      // Step 3: Write agent instructions to file
      const taskFile = this.writeAgentTasks(agentTasks);
      console.log(`📝 Agent tasks written to: ${taskFile}\n`);

      // Step 4: Display instructions for manual execution
      this.displayExecutionInstructions(sourceFiles);

      return {
        success: true,
        filesAnalyzed: sourceFiles.length,
        taskFile: taskFile
      };

    } catch (error) {
      console.error('❌ Error during test generation:', error.message);
      this.results.errors.push(error.message);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Discover all C# source files in the project
   */
  discoverSourceFiles() {
    const sourceFiles = [];
    const excludeDirs = ['bin', 'obj', 'node_modules', 'tests', '.git'];

    const scanDirectory = (dir) => {
      try {
        const items = fs.readdirSync(dir);

        for (const item of items) {
          const fullPath = path.join(dir, item);
          const stat = fs.statSync(fullPath);

          if (stat.isDirectory()) {
            const dirName = path.basename(fullPath);
            if (!excludeDirs.includes(dirName) && !dirName.startsWith('.')) {
              scanDirectory(fullPath);
            }
          } else if (stat.isFile() && item.endsWith('.cs')) {
            // Exclude already existing test files
            if (!item.includes('Test') && !item.includes('Mock')) {
              sourceFiles.push(fullPath);
            }
          }
        }
      } catch (error) {
        console.warn(`⚠️  Cannot scan directory ${dir}: ${error.message}`);
      }
    };

    if (this.options.file) {
      // Single file mode
      const filePath = path.resolve(this.projectPath, this.options.file);
      if (fs.existsSync(filePath)) {
        sourceFiles.push(filePath);
      } else {
        throw new Error(`File not found: ${filePath}`);
      }
    } else {
      // Scan entire project
      scanDirectory(this.projectPath);
    }

    return sourceFiles;
  }

  /**
   * Build agent task descriptions for each source file
   */
  buildAgentTasks(sourceFiles) {
    const tasks = [];

    for (const filePath of sourceFiles) {
      const fileName = path.basename(filePath);
      const relativePath = path.relative(this.projectPath, filePath);

      tasks.push({
        type: 'analyze-and-generate',
        file: filePath,
        relativePath: relativePath,
        agents: [
          {
            name: 'code-analyzer',
            description: `Analyze ${fileName}`,
            prompt: this.buildAnalyzerPrompt(filePath, relativePath)
          },
          {
            name: 'tester',
            description: `Generate tests for ${fileName}`,
            prompt: this.buildTestGeneratorPrompt(filePath, relativePath)
          },
          {
            name: 'reviewer',
            description: `Review tests for ${fileName}`,
            prompt: this.buildReviewerPrompt(filePath, relativePath)
          }
        ]
      });
    }

    return tasks;
  }

  /**
   * Build prompt for code analyzer agent
   */
  buildAnalyzerPrompt(filePath, relativePath) {
    return `
# Code Analysis Task

**File**: ${relativePath}
**Full Path**: ${filePath}

## Objective
Analyze the C# source file and extract comprehensive metadata for test generation.

## Analysis Requirements

1. **Class Information**
   - Class name(s)
   - Access modifiers
   - Base classes and interfaces
   - Generic type parameters

2. **Methods**
   - Public methods (all)
   - Private methods (for reflection-based testing)
   - Static methods
   - Async methods
   - Method parameters and return types
   - Method complexity

3. **Properties**
   - Public properties
   - Getters and setters
   - Auto-properties vs computed

4. **Constructors**
   - Parameter counts
   - Dependency injection patterns

5. **Dependencies**
   - External libraries
   - Internal dependencies
   - Interfaces to mock

6. **Patterns & Features**
   - Static fields/methods
   - Singleton patterns
   - Factory patterns
   - Async/await usage
   - Exception handling

## Output Format
Provide a structured JSON report with all findings. Store in memory with key: "analysis/${relativePath}"

## Next Step
After analysis, coordinate with the 'tester' agent to generate comprehensive unit tests.
`.trim();
  }

  /**
   * Build prompt for test generator agent
   */
  buildTestGeneratorPrompt(filePath, relativePath) {
    const testPath = this.calculateTestPath(relativePath);

    return `
# Test Generation Task

**Source File**: ${relativePath}
**Test File**: ${testPath}

## Objective
Generate comprehensive unit tests with ${this.config.coverageThreshold}% coverage target.

## Test Generation Requirements

1. **Test Project Setup**
   - Create test project if not exists
   - Add necessary NuGet packages:
     - ${this.config.framework}
     - ${this.config.mockingLibrary}
     - ${this.config.assertionLibrary}
   - Configure target framework: ${this.config.targetFramework}

2. **Test Class Structure**
   - Follow naming convention: [ClassName]Tests
   - Use IDisposable for cleanup (if needed)
   - Include test categories/traits

3. **Test Coverage**
   ${this.config.testPatterns.publicMethods.enabled ? '- ✅ All public methods' : ''}
   ${this.config.testPatterns.privateMethods.enabled ? '- ✅ Private methods (using reflection)' : ''}
   ${this.config.testPatterns.properties.enabled ? '- ✅ Properties (getters/setters)' : ''}
   ${this.config.testPatterns.constructors.enabled ? '- ✅ Constructors' : ''}

4. **Edge Cases** (CRITICAL)
   ${this.config.testPatterns.edgeCases.nullParameters ? '- ✅ Null parameters' : ''}
   ${this.config.testPatterns.edgeCases.emptyStrings ? '- ✅ Empty strings' : ''}
   ${this.config.testPatterns.edgeCases.boundaryValues ? '- ✅ Boundary values (min/max)' : ''}
   ${this.config.testPatterns.edgeCases.exceptions ? '- ✅ Exception scenarios' : ''}
   ${this.config.testPatterns.edgeCases.unicodeAndSpecialChars ? '- ✅ Unicode & special characters' : ''}

5. **Test Patterns**
   - Use AAA pattern (Arrange, Act, Assert)
   - Use [Theory] for parameterized tests
   - Use [Fact] for single test cases
   - Add descriptive test names
   - Include XML documentation

6. **Mocking**
   - Mock external dependencies
   - Use ${this.config.mockingLibrary} for mocking
   - Setup mock behaviors appropriately

7. **Assertions**
   - Use ${this.config.assertionLibrary} for fluent assertions
   - Test both positive and negative cases
   - Verify exception messages

## Retrieve Analysis
Get the code analysis from memory key: "analysis/${relativePath}"

## Output
- Create complete test file at: ${testPath}
- Generate test project file if needed
- Store test metadata in memory: "tests/${relativePath}"

## Next Step
After generation, the 'reviewer' agent will review the tests.
`.trim();
  }

  /**
   * Build prompt for reviewer agent
   */
  buildReviewerPrompt(filePath, relativePath) {
    return `
# Test Review Task

**Source File**: ${relativePath}
**Test File**: ${this.calculateTestPath(relativePath)}

## Objective
Review generated tests for quality, completeness, and best practices.

## Review Checklist

1. **Coverage Analysis**
   - Verify ${this.config.coverageThreshold}% coverage target
   - Check all public methods tested
   - Verify edge cases included

2. **Test Quality**
   - Proper test naming conventions
   - Clear Arrange-Act-Assert structure
   - Appropriate use of [Fact] vs [Theory]
   - Good XML documentation

3. **Assertions**
   - Using ${this.config.assertionLibrary} properly
   - Assertions are specific and meaningful
   - Exception testing is thorough

4. **Mocking**
   - Dependencies properly mocked
   - Mock setups are correct
   - Verifications are appropriate

5. **Edge Cases**
   - Null handling tested
   - Empty/whitespace tested
   - Boundary values tested
   - Unicode/special chars tested
   - Exception scenarios covered

6. **Best Practices**
   - No test interdependencies
   - Proper cleanup (IDisposable)
   - Fast execution (no unnecessary delays)
   - Deterministic results

## Retrieve Context
Get test metadata from memory: "tests/${relativePath}"

## Output
- Provide detailed review report
- List any gaps or improvements needed
- Suggest additional test cases if needed
- Store review in memory: "review/${relativePath}"

## Final Step
If review passes, mark tests as complete. Otherwise, suggest improvements to the tester agent.
`.trim();
  }

  /**
   * Calculate test file path from source path
   */
  calculateTestPath(relativePath) {
    const parsedPath = path.parse(relativePath);
    const projectName = path.basename(this.projectPath);
    const testFileName = `${parsedPath.name}Tests${parsedPath.ext}`;
    return path.join(this.config.testDirectory, `${projectName}.Tests`, testFileName);
  }

  /**
   * Write agent tasks to a file for execution
   */
  writeAgentTasks(tasks) {
    const outputPath = path.join(__dirname, 'agent-tasks.json');
    fs.writeFileSync(outputPath, JSON.stringify(tasks, null, 2));
    return outputPath;
  }

  /**
   * Display instructions for executing the agent tasks
   */
  displayExecutionInstructions(sourceFiles) {
    console.log('━'.repeat(70));
    console.log('📋 EXECUTION INSTRUCTIONS');
    console.log('━'.repeat(70));
    console.log('\n🤖 To generate tests using Claude Code agents:\n');

    console.log('Option 1: Use Claude Code directly');
    console.log('─'.repeat(70));
    console.log('Open Claude Code and paste this prompt:\n');

    const prompt = this.buildClaudeCodePrompt(sourceFiles);
    console.log(prompt);

    console.log('\n' + '─'.repeat(70));
    console.log('\nOption 2: Use Claude Flow CLI (if available)');
    console.log('─'.repeat(70));
    console.log('npx claude-flow sparc tdd "Generate tests for all files"');

    console.log('\n' + '─'.repeat(70));
    console.log('\nOption 3: Manual agent spawning');
    console.log('─'.repeat(70));
    console.log('Review agent-tasks.json and spawn agents individually\n');

    console.log('━'.repeat(70));
  }

  /**
   * Build a comprehensive prompt for Claude Code
   */
  buildClaudeCodePrompt(sourceFiles) {
    const fileList = sourceFiles.map(f => `  - ${path.relative(this.projectPath, f)}`).join('\n');

    return `
🚀 Generate comprehensive unit tests for the following C# files:

${fileList}

📋 Requirements:
- Framework: ${this.config.framework}
- Mocking: ${this.config.mockingLibrary}
- Assertions: ${this.config.assertionLibrary}
- Target Coverage: ${this.config.coverageThreshold}%
- Test Directory: ${this.config.testDirectory}/

🤖 Use the following agents IN PARALLEL (single message):
1. Task("Code Analyzer", "Analyze all source files and extract metadata", "code-analyzer")
2. Task("Test Generator", "Generate comprehensive unit tests with edge cases", "tester")
3. Task("Test Reviewer", "Review generated tests for quality and completeness", "reviewer")

✅ Include ALL edge cases:
- Null parameters
- Empty strings
- Boundary values
- Exceptions
- Unicode & special characters

📊 Generate:
- Complete test project
- Test classes for each source file
- README.md with documentation
- Code coverage configuration

🔧 Test patterns:
- AAA pattern (Arrange, Act, Assert)
- [Theory] for parameterized tests
- [Fact] for single cases
- Reflection for private methods
- Comprehensive documentation
`.trim();
  }
}

// CLI Execution
if (require.main === module) {
  const args = process.argv.slice(2);

  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    console.log(`
Automated Unit Test Generator for C# Projects

Usage:
  node generate-tests.js <project-path> [options]

Options:
  --file <path>      Generate tests for a specific file only
  --watch            Watch mode (continuous generation)
  --help, -h         Show this help message

Examples:
  node generate-tests.js ./MyProject
  node generate-tests.js ./MyProject --file Helper.cs
  node generate-tests.js ./MyProject --watch

Features:
  - Analyzes C# code structure
  - Generates comprehensive unit tests
  - Includes edge cases automatically
  - Reviews test quality
  - Uses xUnit, Moq, and FluentAssertions
    `);
    process.exit(0);
  }

  const projectPath = args[0];
  const options = {
    file: args.includes('--file') ? args[args.indexOf('--file') + 1] : null,
    watch: args.includes('--watch')
  };

  const orchestrator = new TestGeneratorOrchestrator(projectPath, options);

  orchestrator.generate()
    .then(result => {
      if (result.success) {
        console.log('\n✅ Test generation setup complete!');
        console.log(`📊 Files analyzed: ${result.filesAnalyzed}`);
      } else {
        console.error('\n❌ Test generation failed:', result.error);
        process.exit(1);
      }
    })
    .catch(error => {
      console.error('\n❌ Fatal error:', error);
      process.exit(1);
    });
}

module.exports = TestGeneratorOrchestrator;
