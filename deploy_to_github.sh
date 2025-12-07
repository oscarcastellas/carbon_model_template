#!/bin/bash
# Helper script to deploy to GitHub

echo "=== Carbon Model Template - GitHub Deployment ==="
echo ""

# Check if remote already exists
if git remote get-url origin > /dev/null 2>&1; then
    echo "Remote 'origin' already exists:"
    git remote get-url origin
    echo ""
    read -p "Do you want to update it? (y/n) " -n 1 -r
    echo ""
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        echo "Deployment cancelled."
        exit 0
    fi
    git remote remove origin
fi

# Get GitHub username
read -p "Enter your GitHub username: " GITHUB_USERNAME

# Get repository name (default to carbon_model_template)
read -p "Enter repository name [carbon_model_template]: " REPO_NAME
REPO_NAME=${REPO_NAME:-carbon_model_template}

# Set remote
GITHUB_URL="https://github.com/${GITHUB_USERNAME}/${REPO_NAME}.git"
echo ""
echo "Setting remote to: ${GITHUB_URL}"
git remote add origin "${GITHUB_URL}"

# Set branch to main
git branch -M main

# Push to GitHub
echo ""
echo "Pushing to GitHub..."
git push -u origin main

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Successfully deployed to GitHub!"
    echo ""
    echo "Repository URL: https://github.com/${GITHUB_USERNAME}/${REPO_NAME}"
    echo ""
    echo "Next steps:"
    echo "1. Visit your repository on GitHub"
    echo "2. Add a description and topics"
    echo "3. Consider adding screenshots of the Excel output"
else
    echo ""
    echo "❌ Push failed. Make sure:"
    echo "1. The repository exists on GitHub"
    echo "2. You have the correct permissions"
    echo "3. You're authenticated (use 'gh auth login' or SSH keys)"
fi

