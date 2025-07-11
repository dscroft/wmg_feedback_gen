"""Setup script for backward compatibility.

Modern packaging uses pyproject.toml, but this setup.py is provided
for compatibility with older tools.
"""

from setuptools import setup

if __name__ == "__main__":
    setup()
