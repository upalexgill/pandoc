# GitHub Copilot Instructions for PowerShell Pandoc Markdown to Word Converter

This document provides instructions for using GitHub Copilot with the PowerShell Pandoc Markdown to Word Converter project.

## Project Overview

The PowerShell Pandoc Markdown to Word Converter is a script that converts `.md` files to `.docx` format using Pandoc.

### Current Version: 1.1.0 (2025-06-11)

**Key Features:**
- Convert `.md` files to `.docx` format
- Automatic Pandoc installation via winget
- Reference document support for custom styling
- Comprehensive test suite (`Test-Converter.ps1`)
- VS Code tasks and launch configurations
- Auto-generate output filenames
- Handle relative and absolute paths
- Attempt to open converted files automatically

### Project Structure
```
CHANGELOG.md          # Version history and changes
ConvertToWord.ps1     # Main conversion script
README.md             # Project documentation
Test-Converter.ps1    # Test suite for validation
examples/             # Sample files and test outputs
github/               # Development configuration
```

### Recent Changes (v1.1.0)
- Added comprehensive test suite for validation
- Added VS Code tasks and launch configurations
- Added changelog file
- Updated README with complete documentation
- Fixed test suite syntax errors and function loading issues
- Improved project structure documentation

## Changelog Management

The project follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) format and adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

### Versioning Guidelines
- **MAJOR** version when you make incompatible API changes
- **MINOR** version when you add functionality in a backward compatible manner
- **PATCH** version when you make backward compatible bug fixes

### Changelog Format
The `CHANGELOG.md` file maintains a structured format with:
- **[Unreleased]** section for upcoming changes
- Version sections with release dates
- Categories: Added, Changed, Deprecated, Removed, Fixed, Security
- Links to version comparisons and releases

### Best Practices
- Update changelog with every significant change
- Use present tense for changelog entries ("Add feature" not "Added feature")
- Group similar changes together
- Include migration notes for breaking changes
- Reference issue numbers when applicable

## Prerequisites

- Ensure you have a GitHub account with access to Copilot
- Install Visual Studio Code or another supported IDE
- Install the GitHub Copilot extension from the marketplace
- PowerShell 5.1+ (Windows PowerShell or PowerShell Core)
- Pandoc (auto-installed by script if missing)

## Setup

1. Open your IDE
2. Install the GitHub Copilot extension
3. Sign in with your GitHub account when prompted
4. Enable Copilot in your IDE settings if necessary

## Usage

- Start typing code in your editor
- Copilot will automatically suggest code completions
- Press `Tab` to accept a suggestion, or `Esc` to dismiss
- Use `Ctrl+Space` (or your IDE's shortcut) to trigger suggestions manually

## Tips

- Write clear comments to guide Copilot's suggestions
- Review and edit Copilot's code for correctness and security
- Use Copilot Labs for advanced features like code explanation and translation

## Troubleshooting

- If suggestions are not appearing, ensure you are signed in and the extension is enabled
- Restart your IDE if issues persist
- Refer to the [official documentation](https://docs.github.com/en/copilot) for more help

## PowerShell Development Best Practices

When working with PowerShell scripts, GitHub Copilot can be particularly helpful. Here are key insights from real development experience:

### Parameter Design
- **Keep it simple**: Start with essential parameters only (e.g., mandatory input file)
- **Make optional parameters truly optional**: Provide sensible defaults and auto-generation
- **Avoid unnecessary complexity**: Parameters like `$TargetFolder` often add confusion without significant benefit
- **Use proper validation**: Add `ValidateScript` attributes to catch issues early

### Path Handling
- **Always use absolute paths** for external commands to avoid permission issues
- **Let PowerShell handle quoting**: When using the call operator `&`, pass variables as separate parameters
  ```powershell
  # ✅ Correct - PowerShell handles quoting automatically
  & $pandocExe $inputFile -o $outputFile
  
  # ❌ Wrong - Manual quoting can cause parsing issues
  & $pandocExe "`"$inputFile`"" -o "`"$outputFile`""
  ```

### Error Handling
- **Check file existence** before processing, especially after path resolution
- **Use proper exit codes** for scripting scenarios (0 = success, 1+ = error)
- **Provide meaningful error messages** with colored output for better UX
- **Handle external command failures** by checking `$LASTEXITCODE`

### Code Organization
- **Break complex logic into functions** for better maintainability
- **Use script-scoped variables** instead of global variables when possible
- **Add comprehensive help documentation** with `.SYNOPSIS`, `.DESCRIPTION`, and `.EXAMPLE`
- **Implement proper resource cleanup** using `try/finally` blocks

### Variable Management
- **Avoid unused variables**: Remove or use `$null = Get-Command` pattern
- **Use consistent naming**: Prefer descriptive names like `$MarkdownFileAbsolute`
- **Initialize variables properly**: Avoid potential null reference issues

### External Tool Integration
- **Detect tool availability** before attempting installation
- **Provide multiple installation methods** (package managers, manual paths)
- **Search common installation locations** as fallback options
- **Handle PATH refresh issues** gracefully after new installations

### User Experience
- **Provide colored output** for different message types (Success=Green, Warning=Yellow, Error=Red)
- **Show progress indicators** for long-running operations
- **Add debug output** for troubleshooting (input/output file paths)
- **Handle file opening failures** gracefully with try/catch
