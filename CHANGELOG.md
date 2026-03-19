# Changelog

All notable changes to AI-Passbolt Migration Tool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Planned
- Support for 1Password CSV export format
- Support for LastPass CSV export format
- Bitwarden import support
- Encrypted CSV export option
- Command-line interface (CLI) mode
- Batch processing for multiple Excel files

---

## [1.0.0] - 2026-03-19

### Added
- **Excel Parser** with support for:
  - Merged cells handling
  - Multi-row headers
  - Multiple sheets support
  - Automatic structure detection
  - Vertical format detection

- **AI Analysis** using Groq API:
  - Smart column detection
  - Context-aware data understanding
  - Automatic Notes population
  - URL and IP validation

- **GUI Application** (customtkinter):
  - File-based conversion tab
  - Clipboard paste tab
  - Settings tab with API key management
  - Preview tab for data validation
  - Dark mode support

- **Export to Passbolt CSV**:
  - All required fields (Group, Title, Username, Password, URL, Notes)
  - UTF-8 encoding
  - Compatible with Passbolt KeePass CSV import

- **Security Features**:
  - Environment variable support for API keys
  - .env file configuration
  - Secure credential handling

- **Testing**:
  - Unit tests for Excel parser
  - Test coverage for standard formats
  - Test coverage for merged cells
  - Test coverage for vertical formats

### Documentation
- Comprehensive README with usage examples
- Installation instructions
- Groq API key setup guide
- Supported formats documentation
- Troubleshooting guide

---

## [0.1.0] - 2026-02-01

### Added
- Initial prototype
- Basic Excel parsing
- Simple CSV export

---

## Version History

| Version | Date | Key Features |
|---------|------|--------------|
| 1.0.0 | 2026-03-19 | Full release with AI, GUI, tests |
| 0.1.0 | 2026-02-01 | Initial prototype |

---

## Upcoming Features (Roadmap)

### v1.1.0 (Q2 2026)
- [ ] 1Password import support
- [ ] LastPass import support
- [ ] Enhanced error handling

### v1.2.0 (Q3 2026)
- [ ] Bitwarden import support
- [ ] Encrypted export option
- [ ] CLI mode

### v2.0.0 (Q4 2026)
- [ ] Batch processing
- [ ] Web interface option
- [ ] Direct Passbolt API integration

---

## Breaking Changes

### v1.0.0
- None (initial release)

---

## Security Notice

⚠️ **Important:** CSV files exported by this tool contain passwords in plain text.
Always delete exported CSV files immediately after importing to Passbolt.

---

## Contributors

- Initial version by AI-Passbolt Contributors

---

**Last updated:** 2026-03-19
