#!/bin/bash

# Script to set up git hooks

echo "Setting up git hooks..."

# Make hooks executable
chmod +x .githooks/pre-commit
chmod +x .githooks/pre-push

# Install pre-commit hook
if [ -d ".git" ]; then
    cp .githooks/pre-commit .git/hooks/pre-commit
    chmod +x .git/hooks/pre-commit
    echo "✓ Pre-commit hook installed successfully"
    
    # Install pre-push hook
    cp .githooks/pre-push .git/hooks/pre-push
    chmod +x .git/hooks/pre-push
    echo "✓ Pre-push hook installed successfully"
else
    echo "Error: .git directory not found. Make sure you're in a git repository."
    exit 1
fi

echo "Git hooks setup complete!"

