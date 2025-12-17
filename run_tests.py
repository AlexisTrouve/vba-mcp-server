#!/usr/bin/env python
"""
Test runner for VBA MCP monorepo.

Runs tests for each package separately to avoid conftest.py conflicts.
"""
import subprocess
import sys
from pathlib import Path


def run_tests(package_name: str, markers: str = "not integration and not windows_only and not slow") -> bool:
    """
    Run tests for a specific package.

    Args:
        package_name: Name of package (core, lite, pro)
        markers: Pytest markers to filter tests

    Returns:
        True if tests passed, False otherwise
    """
    package_dir = Path("packages") / package_name / "tests"

    if not package_dir.exists():
        print(f"[WARNING] No tests found for {package_name}")
        return True

    print(f"\n{'='*70}")
    print(f"Testing: {package_name}")
    print(f"{'='*70}")

    cmd = [
        "pytest",
        str(package_dir),
        "-v",
        "--tb=short",
        "-m", markers
    ]

    result = subprocess.run(cmd)
    return result.returncode == 0


def main():
    """Run all tests."""
    packages = ["core", "lite", "pro"]
    results = {}

    print("VBA MCP Test Suite")
    print("="*70)

    # Run unit tests for each package
    for package in packages:
        results[package] = run_tests(package)

    # Summary
    print(f"\n{'='*70}")
    print("SUMMARY")
    print(f"{'='*70}")

    for package, passed in results.items():
        status = "[PASSED]" if passed else "[FAILED]"
        print(f"{package:10} {status}")

    # Return exit code
    all_passed = all(results.values())
    return 0 if all_passed else 1


if __name__ == "__main__":
    sys.exit(main())
