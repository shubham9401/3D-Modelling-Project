# Example: Creating an Independent Branch for New Implementation

This example demonstrates how to create an independent branch that has no files or history from the main branch.

## Scenario
You want to implement a completely new version of your 3D modeling project without being influenced by the existing code structure.

## Method 1: Using the Utility Script

### Step 1: Run the Script
```bash
./create_independent_branch.sh new-implementation
```

**Output:**
```
=== Independent Branch Creator ===

This will create a new orphan branch called 'new-implementation'
An orphan branch has no parent commits and no files from other branches.

Do you want to continue? (y/n) y
Creating orphan branch 'new-implementation'...
Clearing staging area...
Cleaning working directory...

✓ Independent branch 'new-implementation' created successfully!

Your branch is now empty and independent from all other branches.

Next steps:
1. Add your new files: touch README.md main.py
2. Stage your files: git add .
3. Make initial commit: git commit -m 'Initial commit'
4. Push to remote: git push -u origin new-implementation
```

### Step 2: Add Your New Code
```bash
# Create a new project structure
echo "# New 3D Modeling Implementation" > README.md

# Create your new main file
cat > main.py << 'EOF'
"""
New implementation of 3D modeling project
Clean slate - no dependencies on old code
"""

def main():
    print("New implementation started!")
    # Your new code here
    pass

if __name__ == "__main__":
    main()
EOF

# Create a new configuration file
cat > config.json << 'EOF'
{
  "project": "3D Modeling - New Implementation",
  "version": "2.0.0",
  "description": "Fresh start with improved architecture"
}
EOF
```

### Step 3: Commit and Push
```bash
# Add all new files
git add .

# Check what will be committed
git status

# Commit your new implementation
git commit -m "Initial commit - new implementation with clean architecture"

# Push to remote
git push -u origin new-implementation
```

### Step 4: Verify Independence
```bash
# Check commit history - should only show your new commits
git log --oneline

# Check all branches and their relationships
git log --oneline --graph --all --decorate
```

**Expected Output:**
```
* a1b2c3d (HEAD -> new-implementation) Initial commit - new implementation
```

Notice there's no connection to other branches!

## Method 2: Manual Creation

### Step 1: Create Orphan Branch
```bash
git checkout --orphan new-implementation
```

### Step 2: Remove All Existing Files
```bash
git rm -rf .
```

### Step 3: Verify Clean State
```bash
ls -la
# Should only show .git directory
```

### Step 4: Add Your New Files
```bash
echo "# New Implementation" > README.md
mkdir src
touch src/main.py
```

### Step 5: Commit and Push
```bash
git add .
git commit -m "Initial commit"
git push -u origin new-implementation
```

## Switching Between Branches

### View Current Branch
```bash
git branch --show-current
```

### Switch to Main Branch
```bash
git checkout main
```

**Notice:** All the files from your new-implementation branch disappear, and the main branch files appear!

### Switch Back to Your Independent Branch
```bash
git checkout new-implementation
```

**Notice:** Main branch files disappear, and your new implementation files appear!

## Verifying Your Setup

### Check That Branches Are Independent
```bash
# View all branches with graph
git log --oneline --graph --all

# Expected output shows separate, unconnected branches:
# * a1b2c3d (new-implementation) Initial commit - new implementation
# * x7y8z9a (main) Previous commits from main branch
```

### List All Branches
```bash
git branch -a

# Expected output:
#   main
# * new-implementation
#   remotes/origin/main
#   remotes/origin/new-implementation
```

## Working on Your Independent Branch

### Daily Workflow
```bash
# 1. Make sure you're on the right branch
git checkout new-implementation

# 2. Make changes to your files
nano src/main.py

# 3. Stage and commit
git add .
git commit -m "Add new feature"

# 4. Push to remote
git push
```

### Checking Status
```bash
# Always verify which branch you're on
git status

# View recent commits
git log --oneline -5
```

## Common Questions

**Q: Can I see what's in the main branch while on new-implementation?**  
A: No, not in your working directory. But you can check with git commands:
```bash
git show main:README.md  # Show a specific file from main
git ls-tree -r main      # List all files in main
```

**Q: What if I want to copy one file from main?**  
A: You can selectively copy:
```bash
# While on new-implementation branch
git checkout main -- path/to/specific/file.py
git add path/to/specific/file.py
git commit -m "Import specific file from main"
```

**Q: Can I merge this back to main later?**  
A: Yes, but it will likely have many conflicts:
```bash
git checkout main
git merge new-implementation
# Resolve conflicts manually
```

**Q: How do I delete the independent branch if I don't need it?**  
A:
```bash
# Switch to a different branch first
git checkout main

# Delete local branch
git branch -D new-implementation

# Delete remote branch
git push origin --delete new-implementation
```

## Full Example Session

```bash
# Start from main branch
git checkout main
git status

# Create independent branch
./create_independent_branch.sh redesign-v2

# Add new code
echo "# Project Redesign v2" > README.md
mkdir app
echo "print('v2')" > app/main.py

# Commit
git add .
git commit -m "Initial redesign commit"

# Push
git push -u origin redesign-v2

# Work on some features
echo "# New feature" >> app/main.py
git add .
git commit -m "Add feature A"
git push

# Switch back to main to check something
git checkout main
# (redesign-v2 files are gone, main files are here)

# Switch back to continue work
git checkout redesign-v2
# (main files are gone, redesign-v2 files are back)

# Continue development
echo "# Another feature" >> app/main.py
git add .
git commit -m "Add feature B"
git push
```

## Summary

✅ **What You Learned:**
- How to create an independent branch with no history from main
- How to verify the branch is truly independent
- How to switch between branches
- How to work on the independent branch
- Common operations and troubleshooting

✅ **Key Takeaway:**
An orphan branch lets you start completely fresh while still being in the same repository. No files or history from other branches will appear on your orphan branch.
