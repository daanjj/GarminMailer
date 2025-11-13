# Semantic Versioning Guide for GarminMailer

GarminMailer follows [Semantic Versioning](https://semver.org/) using the format `vMAJOR.MINOR.PATCH`:

## Version Types

- **PATCH** (`v1.4.1`): Bug fixes, security patches, minor improvements
  - No new features
  - Backward compatible
  - Safe to auto-update

- **MINOR** (`v1.5.0`): New features, enhancements
  - Backward compatible
  - May add new functionality
  - Existing workflows continue to work

- **MAJOR** (`v2.0.0`): Breaking changes
  - May require user action
  - Configuration changes
  - UI/workflow changes

## When to Bump Each Version

### Patch (v1.4.0 → v1.4.1)
- Fix email sending bugs
- Fix USB detection issues  
- Improve error messages
- Performance optimizations
- Security fixes

### Minor (v1.4.0 → v1.5.0)
- Add new export formats
- Add configuration options
- Improve UI with new features
- Add support for new Garmin devices
- Add new workflow modes

### Major (v1.4.0 → v2.0.0)
- Change configuration file format
- Remove deprecated features
- Completely redesign UI
- Change default behavior significantly
- Require new system dependencies

## Release Process

### Using the Version Helper (Recommended)

```bash
# Check current version
python version_helper.py current

# Preview next version
python version_helper.py next patch    # Shows v1.4.1
python version_helper.py next minor    # Shows v1.5.0  
python version_helper.py next major    # Shows v2.0.0

# Create and push release
python version_helper.py bump patch "Fix USB detection on macOS"
python version_helper.py bump minor "Add archive-only mode"
python version_helper.py bump major "Redesign configuration system"
```

### Manual Process

```bash
# Create annotated tag
git tag -a v1.4.1 -m "Fix USB detection bug"

# Push tag to trigger GitHub Actions
git push origin v1.4.1
```

## GitHub Actions Integration

- **Automatic builds**: Triggered by pushing any `v*` tag
- **Windows builds**: Always created for every release
- **macOS builds**: Only when manually triggered (to save resources)

## Version Display

The app automatically displays the correct version:
- **Released builds**: Show the git tag version (e.g., "Version: v1.4.0")
- **Development builds**: Show latest tag + "(local)" (e.g., "Version: v1.4.0 (local)")

## Best Practices

1. **Always test before releasing**: Use local builds for testing
2. **Write clear tag messages**: Describe what changed
3. **Follow semver strictly**: Users depend on version meaning
4. **Document breaking changes**: In release notes for major versions
5. **Keep patch releases small**: Focus on single issues when possible

## Examples from GarminMailer History

- `v1.0.0`: Initial release
- `v1.1.0`: Added device labeling system (minor - new feature)
- `v1.2.0`: Added archive mode (minor - new feature)  
- `v1.3.0`: Improved error handling (minor - enhancement)
- `v1.3.1`: Fixed Windows USB detection (patch - bug fix)
- `v1.4.0`: Added email configuration validation (minor - new feature)

Future major version might be:
- `v2.0.0`: Switch to JSON configuration format (breaking change)