#!/bin/bash

# Script to create an independent (orphan) branch in Git
# This creates a branch with no history from the main branch

set -e  # Exit on error

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${GREEN}=== Independent Branch Creator ===${NC}"
echo ""

# Check if branch name is provided
if [ -z "$1" ]; then
    echo -e "${YELLOW}Usage: $0 <branch-name>${NC}"
    echo "Example: $0 my-new-feature"
    exit 1
fi

BRANCH_NAME=$1

# Confirm with user
echo -e "${YELLOW}This will create a new orphan branch called '${BRANCH_NAME}'${NC}"
echo "An orphan branch has no parent commits and no files from other branches."
echo ""
read -p "Do you want to continue? (y/n) " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo -e "${RED}Operation cancelled.${NC}"
    exit 1
fi

# Check if we have uncommitted changes
if [[ -n $(git status -s) ]]; then
    echo -e "${YELLOW}Warning: You have uncommitted changes.${NC}"
    echo "Please commit or stash them before creating an orphan branch."
    echo ""
    echo "Current status:"
    git status -s
    echo ""
    read -p "Do you want to stash your changes and continue? (y/n) " -n 1 -r
    echo ""
    
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo -e "${GREEN}Stashing changes...${NC}"
        git stash push -m "Auto-stash before creating orphan branch ${BRANCH_NAME}"
    else
        echo -e "${RED}Operation cancelled. Please handle your changes first.${NC}"
        exit 1
    fi
fi

# Create the orphan branch
echo -e "${GREEN}Creating orphan branch '${BRANCH_NAME}'...${NC}"
git checkout --orphan "$BRANCH_NAME"

# Remove all files from staging area
echo -e "${GREEN}Clearing staging area...${NC}"
git rm -rf . 2>/dev/null || true

# Remove all files from working directory (except .git)
echo -e "${GREEN}Cleaning working directory...${NC}"
find . -maxdepth 1 ! -name '.git' ! -name '.' ! -name '..' -exec rm -rf {} + 2>/dev/null || true

echo ""
echo -e "${GREEN}✓ Independent branch '${BRANCH_NAME}' created successfully!${NC}"
echo ""
echo "Your branch is now empty and independent from all other branches."
echo ""
echo -e "${YELLOW}Next steps:${NC}"
echo "1. Add your new files: touch README.md main.py"
echo "2. Stage your files: git add ."
echo "3. Make initial commit: git commit -m 'Initial commit'"
echo "4. Push to remote: git push -u origin ${BRANCH_NAME}"
echo ""
echo "To switch back to your previous branch:"
echo "  git checkout -"
echo ""
echo "To see all branches:"
echo "  git branch -a"
