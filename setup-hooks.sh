#!/bin/bash

# Script to set up git hooks

echo "Setting up git hooks..."

# Make pre-commit hook executable
chmod +x .githooks/pre-commit

# Install the hook
if [ -d ".git" ]; then
    cp .githooks/pre-commit .git/hooks/pre-commit
    chmod +x .git/hooks/pre-commit
    echo "✓ Pre-commit hook installed successfully"
else
    echo "Error: .git directory not found. Make sure you're in a git repository."
    exit 1
fi

echo "Git hooks setup complete!"

