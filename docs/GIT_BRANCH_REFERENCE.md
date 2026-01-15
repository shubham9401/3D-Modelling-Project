# Git Branch Quick Reference

## Creating Independent Branches

### Create an Orphan Branch (No History from Main)
```bash
# Create orphan branch
git checkout --orphan new-branch-name

# Remove all files
git rm -rf .

# Start fresh - add your new files
echo "# New Project" > README.md
git add .
git commit -m "Initial commit"

# Push to remote
git push -u origin new-branch-name
```

### Using the Utility Script
```bash
./create_independent_branch.sh new-branch-name
```

## Working with Branches

### List All Branches
```bash
# Local branches
git branch

# All branches (local + remote)
git branch -a

# Remote branches only
git branch -r
```

### Switch Between Branches
```bash
# Switch to another branch
git checkout branch-name

# Switch to previous branch
git checkout -

# Create and switch to new branch (from current branch)
git checkout -b new-branch-name
```

### View Branch History
```bash
# View commit history with graph
git log --oneline --graph --all

# View branch relationships
git log --oneline --graph --decorate --all
```

## Comparing Branches

### Check Differences Between Branches
```bash
# See files that differ
git diff --name-only branch1 branch2

# See actual differences
git diff branch1 branch2

# See commits in branch2 not in branch1
git log branch1..branch2
```

## Managing Branches

### Delete a Branch
```bash
# Delete local branch (must be on different branch)
git branch -d branch-name

# Force delete (if not merged)
git branch -D branch-name

# Delete remote branch
git push origin --delete branch-name
```

### Rename a Branch
```bash
# Rename current branch
git branch -m new-name

# Rename a different branch
git branch -m old-name new-name
```

## Push and Pull

### Push to Remote
```bash
# Push current branch
git push

# Push and set upstream
git push -u origin branch-name

# Push all branches
git push --all
```

### Pull from Remote
```bash
# Pull current branch
git pull

# Fetch all branches
git fetch --all
```

## Checking Status

### Current Branch Status
```bash
# See current branch and changes
git status

# See current branch name only
git branch --show-current

# See which branch you're on
git branch
```

## Tips for Independent Branches

1. **Verify you're on the right branch** before making changes:
   ```bash
   git branch --show-current
   ```

2. **Check if branch is orphan** (no parent commits):
   ```bash
   git log --oneline
   # If it shows no common history with main, it's independent
   ```

3. **Keep branches organized** with clear naming:
   - `feature/new-implementation`
   - `redesign/complete-rewrite`
   - `experiment/alternative-approach`

4. **Before switching branches**, commit or stash changes:
   ```bash
   git stash push -m "WIP: description"
   git checkout other-branch
   # Later, come back and restore
   git stash pop
   ```

## Common Workflows

### Start Fresh Implementation
```bash
git checkout --orphan fresh-start
git rm -rf .
# Add your new code
git add .
git commit -m "Fresh implementation"
git push -u origin fresh-start
```

### Work on Independent Feature
```bash
./create_independent_branch.sh feature-x
# Add code
git add .
git commit -m "Implement feature X"
git push -u origin feature-x
```

### Switch Between Independent Projects
```bash
# Work on project A
git checkout project-a-branch
# ... make changes ...
git add . && git commit -m "Update project A"

# Switch to project B
git checkout project-b-branch
# ... make changes ...
git add . && git commit -m "Update project B"
```

## Troubleshooting

### "Already exists" error when creating orphan branch
```bash
# The branch name already exists, use a different name or delete the old one
git branch -D existing-branch-name
```

### Files from other branch appearing
```bash
# You may have uncommitted changes
git status
git stash  # Save changes
git checkout --orphan new-branch
git rm -rf .
```

### Lost on which branch you're working
```bash
# Always check current branch
git branch --show-current
git status
```
