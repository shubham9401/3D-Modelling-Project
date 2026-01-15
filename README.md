# 3D-Modelling-Project
UG PROJECT

## Git Workflow: Creating Independent Branches

### Can I implement new code in a separate branch without seeing the content of main branch?

**Yes!** You can create an **orphan branch** that is completely independent from the main branch. An orphan branch has no parent commits and starts with a clean slate - no files or history from other branches.

### Quick Start

Use the provided utility script:
```bash
./create_independent_branch.sh my-new-implementation
```

Or manually:
```bash
git checkout --orphan new-branch
git rm -rf .
# Now add your new files and commit
```

### Documentation

For detailed instructions, examples, and best practices, see:
- [Independent Branch Guide](docs/INDEPENDENT_BRANCH_GUIDE.md)

### What You Get
- ✅ No files from main branch
- ✅ No commit history from main branch  
- ✅ Complete freedom to start fresh
- ✅ Can still push to the same repository

This is useful for:
- Complete project redesigns
- Alternative implementations
- Documentation branches (like gh-pages)
- Managing multiple unrelated projects in one repository
