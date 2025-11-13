#!/usr/bin/env python3
"""
Semantic Version Helper for GarminMailer

Usage:
  python version_helper.py current                    # Show current version
  python version_helper.py next patch                 # Show next patch version
  python version_helper.py next minor                 # Show next minor version  
  python version_helper.py next major                 # Show next major version
  python version_helper.py bump patch "Bug fixes"    # Create patch release
  python version_helper.py bump minor "New features" # Create minor release
  python version_helper.py bump major "Breaking changes" # Create major release
"""

import subprocess
import sys
from pathlib import Path

def parse_semver(version_str: str) -> tuple[int, int, int]:
    """Parse semantic version string into major, minor, patch integers."""
    try:
        clean_version = version_str[1:] if version_str.startswith('v') else version_str
        parts = clean_version.split('.')
        if len(parts) >= 3:
            return int(parts[0]), int(parts[1]), int(parts[2])
        elif len(parts) == 2:
            return int(parts[0]), int(parts[1]), 0
        elif len(parts) == 1:
            return int(parts[0]), 0, 0
        else:
            return 0, 0, 0
    except (ValueError, IndexError):
        return 0, 0, 0

def get_current_version() -> str:
    """Get the current highest semantic version tag."""
    try:
        result = subprocess.run(
            ["git", "tag", "--sort=-version:refname"],
            capture_output=True,
            text=True,
            check=True
        )
        tags = [t.strip() for t in result.stdout.split('\n') if t.strip().startswith('v')]
        return tags[0] if tags else "v0.0.0"
    except subprocess.CalledProcessError:
        return "v0.0.0"

def next_version(current: str, bump_type: str) -> str:
    """Calculate next version based on bump type."""
    major, minor, patch = parse_semver(current)
    
    if bump_type == "major":
        return f"v{major + 1}.0.0"
    elif bump_type == "minor":
        return f"v{major}.{minor + 1}.0"
    elif bump_type == "patch":
        return f"v{major}.{minor}.{patch + 1}"
    else:
        raise ValueError(f"Invalid bump type: {bump_type}")

def create_tag(version: str, message: str):
    """Create and push a new git tag."""
    try:
        # Create annotated tag
        subprocess.run(
            ["git", "tag", "-a", version, "-m", message],
            check=True
        )
        print(f"✅ Created tag {version}")
        
        # Push tag
        subprocess.run(
            ["git", "push", "origin", version],
            check=True
        )
        print(f"✅ Pushed tag {version} to origin")
        print(f"🚀 GitHub Actions will now build release {version}")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error creating/pushing tag: {e}")
        sys.exit(1)

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    
    command = sys.argv[1]
    current = get_current_version()
    
    if command == "current":
        print(f"Current version: {current}")
        
    elif command == "next":
        if len(sys.argv) < 3:
            print("Usage: python version_helper.py next <patch|minor|major>")
            sys.exit(1)
        bump_type = sys.argv[2]
        try:
            next_ver = next_version(current, bump_type)
            print(f"Next {bump_type} version: {next_ver}")
        except ValueError as e:
            print(f"❌ {e}")
            sys.exit(1)
            
    elif command == "bump":
        if len(sys.argv) < 4:
            print("Usage: python version_helper.py bump <patch|minor|major> \"Description\"")
            sys.exit(1)
        bump_type = sys.argv[2] 
        message = sys.argv[3]
        
        try:
            next_ver = next_version(current, bump_type)
            print(f"Current version: {current}")
            print(f"Creating {bump_type} release: {next_ver}")
            print(f"Message: {message}")
            
            confirm = input("Proceed? (y/N): ").lower().strip()
            if confirm == 'y':
                create_tag(next_ver, message)
            else:
                print("Cancelled")
        except ValueError as e:
            print(f"❌ {e}")
            sys.exit(1)
    else:
        print(f"❌ Unknown command: {command}")
        print(__doc__)
        sys.exit(1)

if __name__ == "__main__":
    main()