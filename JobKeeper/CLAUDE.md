# Claude Code Configuration - SPARC Development Environment (C# Edition)

## 🚨 CRITICAL: CONCURRENT EXECUTION & FILE MANAGEMENT

**ABSOLUTE RULES**:
1. ALL operations MUST be concurrent/parallel in a single message
2. **NEVER save working files, text/mds and tests to the root folder**
3. ALWAYS organize files in appropriate subdirectories
4. **USE CLAUDE CODE'S TASK TOOL** for spawning agents concurrently, not just MCP

### ⚡ GOLDEN RULE: "1 MESSAGE = ALL RELATED OPERATIONS"

**MANDATORY PATTERNS:**
- **TodoWrite**: ALWAYS batch ALL todos in ONE call (5-10+ todos minimum)
- **Task tool (Claude Code)**: ALWAYS spawn ALL agents in ONE message with full instructions
- **File operations**: ALWAYS batch ALL reads/writes/edits in ONE message
- **Bash commands**: ALWAYS batch ALL terminal operations in ONE message
- **Memory operations**: ALWAYS batch ALL memory store/retrieve in ONE message

### 🎯 CRITICAL: Claude Code Task Tool for Agent Execution

**Claude Code's Task tool is the PRIMARY way to spawn agents:**
```csharp
// ✅ CORRECT: Use Claude Code's Task tool for parallel agent execution
[Single Message]:
  Task("Research agent", "Analyze requirements and patterns...", "researcher")
  Task("Coder agent", "Implement core features...", "coder")
  Task("Tester agent", "Create comprehensive tests...", "tester")
  Task("Reviewer agent", "Review code quality...", "reviewer")
  Task("Architect agent", "Design system architecture...", "system-architect")
```

**MCP tools are ONLY for coordination setup:**
- `mcp__claude-flow__swarm_init` - Initialize coordination topology
- `mcp__claude-flow__agent_spawn` - Define agent types for coordination
- `mcp__claude-flow__task_orchestrate` - Orchestrate high-level workflows

### 📁 File Organization Rules

**NEVER save to root folder. Use these directories:**
- `/src` - Source code files (.cs, .csproj)
- `/tests` - Test projects (xUnit, NUnit, MSTest)
- `/docs` - Documentation and markdown files
- `/config` - Configuration files (appsettings.json, etc.)
- `/scripts` - PowerShell/batch utility scripts
- `/examples` - Example code and samples

## Project Overview

This project uses SPARC (Specification, Pseudocode, Architecture, Refinement, Completion) methodology with Claude-Flow orchestration for systematic Test-Driven Development in C#.

## SPARC Commands

### Core Commands
- `npx claude-flow sparc modes` - List available modes
- `npx claude-flow sparc run <mode> "<task>"` - Execute specific mode
- `npx claude-flow sparc tdd "<feature>"` - Run complete TDD workflow
- `npx claude-flow sparc info <mode>` - Get mode details

### Batchtools Commands
- `npx claude-flow sparc batch <modes> "<task>"` - Parallel execution
- `npx claude-flow sparc pipeline "<task>"` - Full pipeline processing
- `npx claude-flow sparc concurrent <mode> "<tasks-file>"` - Multi-task processing

### C# Build Commands
- `dotnet build` - Build solution/project
- `dotnet test` - Run all tests
- `dotnet run --project <path>` - Run application
- `dotnet clean` - Clean build artifacts
- `dotnet restore` - Restore NuGet packages
- `dotnet format` - Format code
- `dotnet publish` - Publish for deployment
- `dotnet add package <name>` - Add NuGet package

### Code Quality Commands
- `dotnet format --verify-no-changes` - Verify formatting
- `dotnet test --collect:"XPlat Code Coverage"` - Test with coverage
- `dotnet build /warnaserror` - Treat warnings as errors

## SPARC Workflow Phases

1. **Specification** - Requirements analysis (`sparc run spec-pseudocode`)
2. **Pseudocode** - Algorithm design (`sparc run spec-pseudocode`)
3. **Architecture** - System design (`sparc run architect`)
4. **Refinement** - TDD implementation (`sparc tdd`)
5. **Completion** - Integration (`sparc run integration`)

## Code Style & Best Practices

- **Modular Design**: Files under 500 lines, classes with single responsibility
- **Environment Safety**: Use appsettings.json, never hardcode secrets
- **Test-First**: Write tests before implementation (TDD)
- **Clean Architecture**: Separate concerns (UI, Business Logic, Data Access)
- **Documentation**: XML documentation comments for public APIs
- **Naming Conventions**: PascalCase for classes/methods, camelCase for private fields
- **Async/Await**: Use async operations for I/O bound work
- **Dependency Injection**: Use DI container for loose coupling
- **Exception Handling**: Use specific exceptions, avoid catching generic Exception

## 🚀 Available Agents (54 Total + C# Specific)

### Core Development
`coder`, `reviewer`, `tester`, `planner`, `researcher`

### C#-Specific Agents
`dotnet-dev`, `aspnet-core-dev`, `winforms-dev`, `wpf-dev`, `blazor-dev`, `entity-framework-dev`, `xamarin-dev`, `maui-dev`, `azure-dev`, `sql-server-dev`

### Swarm Coordination
`hierarchical-coordinator`, `mesh-coordinator`, `adaptive-coordinator`, `collective-intelligence-coordinator`, `swarm-memory-manager`

### Consensus & Distributed
`byzantine-coordinator`, `raft-manager`, `gossip-coordinator`, `consensus-builder`, `crdt-synchronizer`, `quorum-manager`, `security-manager`

### Performance & Optimization
`perf-analyzer`, `performance-benchmarker`, `task-orchestrator`, `memory-coordinator`, `smart-agent`

### GitHub & Repository
`github-modes`, `pr-manager`, `code-review-swarm`, `issue-tracker`, `release-manager`, `workflow-automation`, `project-board-sync`, `repo-architect`, `multi-repo-swarm`

### SPARC Methodology
`sparc-coord`, `sparc-coder`, `specification`, `pseudocode`, `architecture`, `refinement`

### Specialized Development
`backend-dev`, `mobile-dev`, `ml-developer`, `cicd-engineer`, `api-docs`, `system-architect`, `code-analyzer`, `base-template-generator`

### Testing & Validation
`tdd-london-swarm`, `production-validator`

### Migration & Planning
`migration-planner`, `swarm-init`

## 🎯 Claude Code vs MCP Tools

### Claude Code Handles ALL EXECUTION:
- **Task tool**: Spawn and run agents concurrently for actual work
- File operations (Read, Write, Edit, MultiEdit, Glob, Grep)
- Code generation and C# programming
- Bash/PowerShell commands and system operations
- Implementation work
- Project navigation and analysis
- TodoWrite and task management
- Git operations
- NuGet package management
- Testing and debugging

### MCP Tools ONLY COORDINATE:
- Swarm initialization (topology setup)
- Agent type definitions (coordination patterns)
- Task orchestration (high-level planning)
- Memory management
- Neural features
- Performance tracking
- GitHub integration

**KEY**: MCP coordinates the strategy, Claude Code's Task tool executes with real agents.

## 🚀 Quick Setup

```bash
# Add MCP servers (Claude Flow required, others optional)
claude mcp add claude-flow npx claude-flow@alpha mcp start
claude mcp add ruv-swarm npx ruv-swarm mcp start  # Optional: Enhanced coordination
claude mcp add flow-nexus npx flow-nexus@latest mcp start  # Optional: Cloud features
```

## MCP Tool Categories

### Coordination
`swarm_init`, `agent_spawn`, `task_orchestrate`

### Monitoring
`swarm_status`, `agent_list`, `agent_metrics`, `task_status`, `task_results`

### Memory & Neural
`memory_usage`, `neural_status`, `neural_train`, `neural_patterns`

### GitHub Integration
`github_swarm`, `repo_analyze`, `pr_enhance`, `issue_triage`, `code_review`

### System
`benchmark_run`, `features_detect`, `swarm_monitor`

### Flow-Nexus MCP Tools (Optional Advanced Features)
Flow-Nexus extends MCP capabilities with 70+ cloud-based orchestration tools:

**Key MCP Tool Categories:**
- **Swarm & Agents**: `swarm_init`, `swarm_scale`, `agent_spawn`, `task_orchestrate`
- **Sandboxes**: `sandbox_create`, `sandbox_execute`, `sandbox_upload` (cloud execution)
- **Templates**: `template_list`, `template_deploy` (pre-built project templates)
- **Neural AI**: `neural_train`, `neural_patterns`, `seraphina_chat` (AI assistant)
- **GitHub**: `github_repo_analyze`, `github_pr_manage` (repository management)
- **Real-time**: `execution_stream_subscribe`, `realtime_subscribe` (live monitoring)
- **Storage**: `storage_upload`, `storage_list` (cloud file management)

**Authentication Required:**
- Register: `mcp__flow-nexus__user_register` or `npx flow-nexus@latest register`
- Login: `mcp__flow-nexus__user_login` or `npx flow-nexus@latest login`
- Access 70+ specialized MCP tools for advanced orchestration

## 🚀 Agent Execution Flow with Claude Code

### The Correct Pattern:

1. **Optional**: Use MCP tools to set up coordination topology
2. **REQUIRED**: Use Claude Code's Task tool to spawn agents that do actual work
3. **REQUIRED**: Each agent runs hooks for coordination
4. **REQUIRED**: Batch all operations in single messages

### Example C# Windows Forms Development:

```csharp
// Single message with all agent spawning via Claude Code's Task tool
[Parallel Agent Execution]:
  Task("WinForms Developer", "Build main form with data grid. Use hooks for coordination.", "winforms-dev")
  Task("Database Architect", "Design SQL Server schema. Coordinate with backend via memory.", "sql-server-dev")
  Task("Business Logic Developer", "Create service layer with validation. Store contracts in memory.", "dotnet-dev")
  Task("Test Engineer", "Write xUnit tests with 90% coverage. Check memory for API contracts.", "tester")
  Task("Code Reviewer", "Review C# code quality and SOLID principles. Report findings via hooks.", "reviewer")

  // All todos batched together
  TodoWrite { todos: [...8-10 todos...] }

  // All file operations together
  Write "src/JobKeeper.WinForms/MainForm.cs"
  Write "src/JobKeeper.WinForms/Models/Job.cs"
  Write "tests/JobKeeper.Tests/Unit/JobServiceTests.cs"
  Write "config/appsettings.json"
```

## 📋 Agent Coordination Protocol

### Every Agent Spawned via Task Tool MUST:

**1️⃣ BEFORE Work:**
```bash
npx claude-flow@alpha hooks pre-task --description "[task]"
npx claude-flow@alpha hooks session-restore --session-id "swarm-[id]"
```

**2️⃣ DURING Work:**
```bash
npx claude-flow@alpha hooks post-edit --file "[file]" --memory-key "swarm/[agent]/[step]"
npx claude-flow@alpha hooks notify --message "[what was done]"
```

**3️⃣ AFTER Work:**
```bash
npx claude-flow@alpha hooks post-task --task-id "[task]"
npx claude-flow@alpha hooks session-end --export-metrics true
```

## 🎯 Concurrent Execution Examples

### ✅ CORRECT WORKFLOW: MCP Coordinates, Claude Code Executes

```csharp
// Step 1: MCP tools set up coordination (optional, for complex tasks)
[Single Message - Coordination Setup]:
  mcp__claude-flow__swarm_init { topology: "mesh", maxAgents: 6 }
  mcp__claude-flow__agent_spawn { type: "researcher" }
  mcp__claude-flow__agent_spawn { type: "dotnet-dev" }
  mcp__claude-flow__agent_spawn { type: "tester" }

// Step 2: Claude Code Task tool spawns ACTUAL agents that do the work
[Single Message - Parallel Agent Execution]:
  // Claude Code's Task tool spawns real agents concurrently
  Task("Research agent", "Analyze C# design patterns and best practices. Check memory for prior decisions.", "researcher")
  Task("WinForms Developer", "Implement Windows Forms UI with modern controls. Coordinate via hooks.", "winforms-dev")
  Task("Database Developer", "Design and implement Entity Framework Core models. Store decisions in memory.", "entity-framework-dev")
  Task("Test Engineer", "Create comprehensive xUnit test suite with 90% coverage.", "tester")
  Task("Code Reviewer", "Review C# code quality, SOLID principles, and security. Document findings.", "reviewer")

  // Batch ALL todos in ONE call
  TodoWrite { todos: [
    {id: "1", content: "Research WinForms best practices", status: "in_progress", priority: "high"},
    {id: "2", content: "Design database schema", status: "in_progress", priority: "high"},
    {id: "3", content: "Implement main form UI", status: "pending", priority: "high"},
    {id: "4", content: "Create data models", status: "pending", priority: "high"},
    {id: "5", content: "Write unit tests", status: "pending", priority: "medium"},
    {id: "6", content: "Integration tests", status: "pending", priority: "medium"},
    {id: "7", content: "XML documentation", status: "pending", priority: "low"},
    {id: "8", content: "Performance optimization", status: "pending", priority: "low"}
  ]}

  // Parallel file operations
  Bash "mkdir -p src/JobKeeper.WinForms/{Forms,Models,Services,Utilities}"
  Write "src/JobKeeper.WinForms/JobKeeper.WinForms.csproj"
  Write "src/JobKeeper.WinForms/Program.cs"
  Write "src/JobKeeper.WinForms/MainForm.cs"
  Write "tests/JobKeeper.Tests/JobKeeper.Tests.csproj"
```

### ❌ WRONG (Multiple Messages):
```csharp
Message 1: mcp__claude-flow__swarm_init
Message 2: Task("agent 1")
Message 3: TodoWrite { todos: [single todo] }
Message 4: Write "file.cs"
// This breaks parallel coordination!
```

## Performance Benefits

- **84.8% SWE-Bench solve rate**
- **32.3% token reduction**
- **2.8-4.4x speed improvement**
- **27+ neural models**

## Hooks Integration

### Pre-Operation
- Auto-assign agents by file type (.cs, .csproj, .sln)
- Validate commands for safety
- Prepare resources automatically
- Optimize topology by complexity
- Cache searches

### Post-Operation
- Auto-format code (dotnet format)
- Train neural patterns
- Update memory
- Analyze performance
- Track token usage

### Session Management
- Generate summaries
- Persist state
- Track metrics
- Restore context
- Export workflows

## Advanced Features (v2.0.0)

- 🚀 Automatic Topology Selection
- ⚡ Parallel Execution (2.8-4.4x speed)
- 🧠 Neural Training
- 📊 Bottleneck Analysis
- 🤖 Smart Auto-Spawning
- 🛡️ Self-Healing Workflows
- 💾 Cross-Session Memory
- 🔗 GitHub Integration

## C# Project Structure

```
JobKeeper/
├── src/
│   └── JobKeeper.WinForms/
│       ├── Forms/           # Windows Forms (UI)
│       ├── Models/          # Data models
│       ├── Services/        # Business logic
│       ├── Utilities/       # Helper classes
│       ├── Program.cs       # Entry point
│       └── *.csproj         # Project file
├── tests/
│   └── JobKeeper.Tests/
│       ├── Unit/            # Unit tests
│       ├── Integration/     # Integration tests
│       └── *.csproj         # Test project file
├── docs/                    # Documentation
├── config/                  # Configuration files
├── scripts/                 # Build/deployment scripts
└── examples/                # Sample code
```

## Integration Tips

1. Start with basic swarm init
2. Scale agents gradually
3. Use memory for context
4. Monitor progress regularly
5. Train patterns from success
6. Enable hooks automation
7. Use GitHub tools first
8. Leverage .NET CLI for automation
9. Use appsettings.json for configuration
10. Follow C# coding conventions

## Support

- Documentation: https://github.com/ruvnet/claude-flow
- Issues: https://github.com/ruvnet/claude-flow/issues
- Flow-Nexus Platform: https://flow-nexus.ruv.io (registration required for cloud features)
- .NET Documentation: https://learn.microsoft.com/dotnet/

---

Remember: **Claude Flow coordinates, Claude Code creates!**

# important-instruction-reminders
Do what has been asked; nothing more, nothing less.
NEVER create files unless they're absolutely necessary for achieving your goal.
ALWAYS prefer editing an existing file to creating a new one.
NEVER proactively create documentation files (*.md) or README files. Only create documentation files if explicitly requested by the User.
Never save working files, text/mds and tests to the root folder.
