# AGENTS.md - Development Guidelines for FanucFocasTutorial

This file contains guidelines and commands for agentic coding agents working in this repository.

## Build Commands

### Primary Build
```bash
dotnet build
```

### Release Build
```bash
dotnet build -c Release
```

### Run Application
```bash
dotnet run
```

### Clean Build
```bash
dotnet clean && dotnet build
```

## Testing Commands

This project currently has no formal test framework. To add testing:
1. Create test project with `dotnet new xunit -n FanucFocasTutorial.Tests`
2. Add reference: `dotnet add reference FanucFocasTutorial-master.csproj`
3. Run tests with: `dotnet test`

## Code Style Guidelines

### Import Organization
- System imports first, grouped by namespace
- Third-party imports second
- Local/namespace imports last
- Use `using System;` before `using System.Collections.Generic;`
- Remove unused imports

### Naming Conventions
- **Classes**: PascalCase (e.g., `CNCConnectionManager`)
- **Methods**: PascalCase (e.g., `ConnectAll()`)
- **Fields**: camelCase with underscore prefix for private fields (e.g., `_connections`)
- **Properties**: PascalCase (e.g., `IpAddress`)
- **Constants**: UPPER_CASE with underscores (e.g., `DEFAULT_TIMEOUT`)
- **Local variables**: camelCase (e.g., `connection`)

### Type Guidelines
- Enable nullable reference types (`<Nullable>enable</Nullable>`)
- Use `var` for local variable type inference when type is obvious
- Prefer explicit types for public APIs and complex expressions
- Use interfaces for public contracts where appropriate

### Error Handling
- Use try-catch blocks for external library calls (FOCAS operations)
- Log errors with meaningful messages
- Return boolean or status codes for operation success/failure
- Handle disposal patterns with `using` statements or `IDisposable`

### Formatting
- 4 spaces for indentation (no tabs)
- Opening braces on new line for class/method declarations
- Opening braces on same line for control structures (if, for, while)
- One blank line between method definitions
- Maximum line length: 120 characters

### Documentation
- Use XML documentation comments for public APIs
- Include parameter descriptions and return value meanings
- Add TODO comments with developer initials and date

### FANUC FOCAS Specific Guidelines
- Always check return codes from FOCAS library calls
- Use constants from `Focas1` class for error codes and limits
- Handle connection timeouts gracefully
- Implement proper disposal for CNC connections
- Use `Marshal` class for P/Invoke operations carefully

### Database Guidelines
- Use Entity Framework Core with SQLite
- Define entities with `[Table]` and `[Key]` attributes
- Use `DateTime.Now` for timestamp fields
- Implement proper connection string management

### UI Guidelines (Windows Forms)
- Use meaningful control names with prefixes (`_btn`, `_txt`, `_grid`)
- Implement proper event handler patterns
- Use `InvokeRequired` for cross-thread UI updates
- Follow Windows Forms design patterns

## Project Structure Notes

### Key Directories
- `libs/` - FANUC FOCAS library files
- `Forms/` - Windows Forms UI components
- Root level - Main application files

### Important Files
- `fwlib32.cs` - FANUC FOCAS P/Invoke declarations (DO NOT MODIFY)
- `CNCConnectionManager.cs` - CNC connection management
- `MainForm.cs` - Primary application interface

## Development Workflow

1. Always build before committing changes
2. Test CNC connections with actual hardware when possible
3. Validate data integrity for database operations
4. Check FOCAS return codes and handle errors appropriately
5. Follow disposal patterns for external resources

## Platform Considerations

- Target: x86 Windows (required for FANUC libraries)
- Framework: .NET 9.0 Windows Forms
- Self-contained deployment supported
- FANUC libraries require 32-bit compatibility