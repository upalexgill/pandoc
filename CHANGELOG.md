# Changelog

All notable changes to the PowerShell Pandoc Markdown to Word Converter will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2025-06-11

### Added
- Comprehensive test suite (`Test-Converter.ps1`) for validation
- VS Code tasks and launch configurations for development  
- This changelog file
- Updated README with complete project documentation

### Fixed
- Test suite syntax errors and function loading issues
- README documentation updated to match current codebase
- Project structure documentation

### Note
- This version maintains the original script functionality while adding testing infrastructure
- The script uses the basic parameter structure (mandatory MarkdownFile parameter)

## [1.0.0] - Previous Version

### Added
- Basic Markdown to Word conversion functionality
- Automatic Pandoc installation via winget
- Reference document support for custom styling
- Basic error handling and user feedback
- README documentation

### Features
- Convert `.md` files to `.docx` format
- Auto-generate output filenames
- Handle relative and absolute paths
- Attempt to open converted files automatically
