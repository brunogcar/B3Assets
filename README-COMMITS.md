# B3Assets Commit Message Patterns & Templates

This repository demonstrates meaningful commit message practices for Apps Script projects. Based on 100+ git amendments analyzing actual code changes.

## 🎯 Philosophy

Good commit messages answer: **What changed?** **Why did it change?** **How does it impact the codebase?**

### ❌ Bad Examples
```bash
"misc fixes"
"small improvements"  
"minor updates"
"bump to version X.X.X"
```

### ✅ Good Examples (from B3Assets)
```bash
fix(- Edit - Process.gs): 
- processEditGeneric now reads rows 1–5 safely with Math.min(5, LR)
- A5 and Row5 now check if row exists before slicing to avoid errors

refactor(performance): 
- Import pipeline optimized: doImportBasics and doImportFinancials now open
  the source spreadsheet once and pass reference (ss_sr) to child functions,
  eliminating repeated SpreadsheetApp.openById() calls (~20 → 1 per batch)

feat(console.log): 
- improved console.log() replacement to Logger.log() in FUNCTIONS.gs
- enhanced Logging documentation with examples
```

## 📊 Commit Categories & Examples

### `fix()` - Bug Fixes

| Pattern | Example |
|---------|---------|
| **File-specific** | `fix(- Save - Process.gs): removed empty lines between processSaveExtra` |
| **Function-specific** | `fix(doExportExtra): Added exportExtraConfig.target_co check before exporting` |
| **Logic corrections** | `fix(save/edit): Changed IsEqual from every() to some() for OR logic` |

### `refactor()` - Code Improvements

| Pattern | Example |
|---------|---------|
| **Performance** | `refactor(performance): Import pipeline optimized (~20 → 1 API call)` |
| **Code cleanup** | `refactor(LogDebug): remove unused safe wrappers and simplify function` |
| **Architecture** | `refactor(doCheckDATA): Remove _CONFIG_VALUE_CACHE - values change during run` |

### `feat()` - New Functionality

| Pattern | Example |
|---------|---------|
| **New features** | `feat(doExportProventos): added doExportProventos() function for exports` |
| **Enhancements** | `feat(save/edit/import): Data_Header and Data_Body standards established` |
| **Configuration** | `feat(config): applied getConfigValue to functions where new functionality can be used` |

### `bump()` - Version Updates

```bash
bump: 
- version bump 17.0.5 - significant changes for new release

commit(source):
- just committing changes to the source
- no changelog prepared, bump to trigger version update
```

## 📝 Writing Your Commit Message Template

### 1. Start with a Category

Choose from: `fix`, `refactor`, `feat`, `perf`, `clean`, `bump`, `revert`

### 2. Mention Affected Components

- File names: `- Edit - Process.gs`, `FUNCTIONS.gs`
- Function names: `processEditGeneric`, `doCheckDATA`
- Code patterns: `=== → ==null`, `forEach() → for loop`

### 3. Describe the Change Impact

- What problem was solved?
- What performance gain achieved?
- What configuration changed?

### 4. Use Bullet Points (Optional but Recommended)

For complex changes, use multiple bullet points:

```bash
refactor(functions): 
- removing unused functions across all modules
- reorganizing functions into proper files (Save, Export, Edit)
- added JS Docs for better code documentation
```

## 🎨 Real Examples from B3Assets History

### Spreadsheet File Changes

```bash
fix(- Data - Process.gs): 
- small fix to complete move code to helper functions

feat(dates): 
- created helper functions for date handling
- extractAndValidateDates validates and extracts dates from ranges
```

### Multi-File Refactoring

```bash
refactor(save/edit/export): 
- remove dead checks and reduce spreadsheet API calls
- Removed unreachable isValidDate() checks in doSaveFinancial / doEditFinancial
```

### Performance Optimizations

```bash
refactor(performance): 
- Import pipeline optimized: doImportBasics and doImportFinancials now open
  the source spreadsheet once and pass reference (ss_sr) to child functions,
  eliminating repeated SpreadsheetApp.openById() calls (~20 → 1 per batch)
```

### Configuration Updates

```bash
feat(console.log): 
- improved console.log() replacement to Logger.log() in FUNCTIONS.gs
- enhanced Logging documentation with examples
```

## 📈 Viewing Your Commit History

```bash
# Recent commits (showing meaningful messages)
git log --oneline -20

# See commit subjects only
git log --format="%h %s" | Select-Object -First 30

# Detailed view of latest commit
git show HEAD
```

Expected output after amendments:
```bash
$ git log --oneline -5

6e443e4 feat(console.log): improved console.log() replacement to Logger.log()
7a16b1e fix(- Edit - Process.gs): processEditGeneric now reads rows safely
2c4d625 refactor(LogDebug): Added _DBG_FETCHING flag to prevent race condition
ba92270 feat(- Save - Process.gs): retry after failed upload attempt
edf2be9 refactor(functions): removing unused functions across all modules
```

## 🚀 Quick Command Reference

### Common Git Commands

```bash
# View recent commits
git log --oneline -30

# See what changed in last commit
git show HEAD --stat

# Check current branch
git branch --show-current

# Stage changes for commit
git add .

# Create commit with meaningful message
git commit -m "refactor(LogDebug):
- Added _DBG_FETCHING flag to prevent race condition
- Updated LogDebug calls for spreadsheet/sheet cache hits"

# Push to remote (with force for amended history)
git push --force-with-lease origin main
```

## 📚 Related Documentation

- **Style Guide**: See `COMMITS-STYLE-GUIDE.md` for detailed patterns
- **Amended Commands**: See `COMMIT-AMEND-COMMANDS.txt` (100 meaningful messages)
- **Comparison**: See `COMMIT-MESSAGE-COMPARISON.txt` for old vs new examples

## 🌐 Template Repository Reference

For inspiration from other projects using similar commit message styles:

- [Agent repo commits](https://github.com/brunogcar/agent/commits/main/)
- Check their latest commits for examples of meaningful descriptions

---

**Generated from**: B3Assets git history analysis (100+ amended commits)
**Last updated**: $(date +%Y-%m-%d)
