# PowerShell Pandoc Markdown to Word Converter

A PowerShell script that converts Markdown files to Word documents using Pandoc, with automatic installation, comprehensive error handling, and modular architecture.

## Features

- **Automatic Pandoc Installation**: Installs Pandoc via winget if not already available
- **Intelligent Path Handling**: Works with both relative and absolute file paths
- **Reference Document Support**: Uses custom styling when a reference document is available
- **Auto-opens Output**: Automatically opens the converted Word document
- **Comprehensive Error Handling**: Detailed error checking and user-friendly messages
- **Cross-session Compatibility**: Handles PATH issues after fresh Pandoc installations
- **Test Suite**: Automated testing for reliability and quality assurance

## Prerequisites

- Windows 10/11 with PowerShell 5.1 or later
- winget (Windows Package Manager) - typically pre-installed on modern Windows

## Installation

1. Clone or download this repository
2. No additional setup required - the script will install Pandoc automatically if needed

## Usage

### Get Help

View comprehensive help documentation:

```powershell
Get-Help .\ConvertToWord.ps1 -Full
```

### Quick Examples

```powershell
# Convert with auto-generated output name
.\ConvertToWord.ps1 "example.md"

# Convert with specific output name  
.\ConvertToWord.ps1 "input.md" "output.docx"

# Convert example files
.\ConvertToWord.ps1 "examples\example-a.md"
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `MarkdownFile` | String | Yes | Path to the Markdown file to convert |
| `WordFile` | String | No | Output Word file path (auto-generated if not specified) |

## Reference Document Styling

The script automatically looks for a reference document at `docs\reference.docx` to apply custom styling to the output. If found, this document's styles will be applied to the converted Word document.

To use custom styling:
1. Create a `docs` folder in the script directory
2. Place your styled Word document as `docs\reference.docx`
3. Run the conversion - the script will automatically use your reference document

## Examples

The `examples` folder contains sample files for testing:
- `example-a.md` - Basic Markdown with headers, lists, and code blocks
- `example-b.md` - More complex content for testing
- Corresponding `.docx` output files for comparison

## Error Handling

The script handles various error scenarios:

- **Missing Pandoc**: Automatically installs via winget
- **PATH Issues**: Uses full path to Pandoc executable when needed
- **Missing Input File**: Clear error message with file path
- **Permission Issues**: Uses absolute paths to avoid directory permission problems
- **Conversion Failures**: Displays detailed error information

## Troubleshooting

### "Pandoc not found" after installation
- Restart your PowerShell session
- The script will try to use the full path to Pandoc as a fallback

### "Cannot find markdown file"
- Verify the file path is correct
- Use absolute paths if relative paths aren't working
- Check file permissions

### Conversion fails
- Ensure the Markdown file is valid
- Check that you have write permissions to the output directory
- Review any error messages displayed by the script

## Script Architecture

The script follows a straightforward, reliable approach:

### Core Functionality
- Automatic Pandoc installation with winget integration
- Intelligent file path resolution (relative and absolute paths)
- Reference document styling support
- Comprehensive error handling with colored output

### Error Handling
- Proper exit codes for scripting scenarios (0 = success, 1+ = error)
- Graceful handling of PATH issues after Pandoc installation
- Clear error messages for troubleshooting
- Automatic retry mechanisms for common issues

### Testing Infrastructure
- Automated test suite (`Test-Converter.ps1`) validates core functionality
- Tests cover conversion, file handling, help documentation, and error scenarios
- VS Code integration for easy development and testing

## Advanced Usage

### Custom Styling with Reference Documents

1. Create a `docs` folder in the script directory
2. Place your styled Word document as `docs\reference.docx`
3. The script will automatically apply your custom styles

### Command Line Parameters

```powershell
# View all available parameters
Get-Help .\ConvertToWord.ps1 -Parameter *
```

### Integration with Other Scripts

The script returns proper exit codes for integration:
- `0`: Successful conversion
- `1`: Error occurred (file not found, conversion failed, etc.)

### VS Code Integration

For VS Code users, the project includes:
- **Tasks** (`Ctrl+Shift+P` → "Run Task"):
  - "Test Converter" - Run the full test suite
  - "Convert Example A/B" - Quick conversion testing
  - "Show Help" - Display full help documentation
- **Debug Configurations** (`F5`):
  - Debug the main script with example files
  - Debug the test suite for development

## Project Structure

```
powershell-pandoc/
├── ConvertToWord.ps1          # Main conversion script
├── Test-Converter.ps1         # Test suite for validation
├── README.md                  # Documentation
├── CHANGELOG.md              # Version history and changes
├── examples/                 # Sample files for testing
│   ├── example-a.md         # Basic Markdown sample
│   ├── example-a.docx       # Converted output example
│   ├── example-b.md         # Advanced Markdown sample
│   └── example-b.docx       # Converted output example
├── docs/                    # Optional styling documents
│   └── reference.docx       # Custom Word styles (optional)
├── .vscode/                 # VS Code configuration
│   ├── tasks.json          # Build and test tasks
│   └── launch.json         # Debug configurations
└── github/
    └── copilot-instructions.md  # Development guidelines
```

## Contributing

Feel free to submit issues and enhancement requests! When contributing:

1. Follow the PowerShell best practices outlined in `github/copilot-instructions.md`
2. Run the test suite with `.\Test-Converter.ps1` before submitting
3. Update the changelog for any significant changes
4. Ensure all functions have proper help documentation

## License

This project is open source. See the repository for license details.

## Related Tools

- [Pandoc](https://pandoc.org/) - Universal document converter
- [winget](https://docs.microsoft.com/en-us/windows/package-manager/winget/) - Windows Package Manager
