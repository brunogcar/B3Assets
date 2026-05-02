# B3Assets Commit Message Style Guide

Based on 100 meaningful git amendments, here are the patterns and templates for writing commit messages.

## 📋 Commit Message Format

```
category(FILE or FUNCTION): 
- bullet point describing specific change
- second bullet point with details (optional)
- code pattern used (if applicable)
```

## ✅ Examples from Production Use

### Bug Fixes (`fix()`)

```bash
fix(- Edit - Process.gs):
- processEditGeneric now reads rows 1–5 safely with Math.min(5, LR)
- A5 and Row5 now check if row exists before slicing to avoid errors

fix(doExportExtra):
- Added exportExtraConfig.target_co check before exporting
- Logs "⏭️ EXPORT skipped: ${SheetName} - not exportable" if not configured
```

### Refactoring (`refactor()`)

```bash
refactor(LogDebug):
- Added _DBG_FETCHING flag to prevent race condition when fetching DBG config
- Changed ORDER array from 3 levels to 4 levels (added 'ALL')

refactor(performance):
- Import pipeline optimized: doImportBasics and doImportFinancials now open
  the source spreadsheet once and pass reference (ss_sr) to child functions,
  eliminating repeated SpreadsheetApp.openById() calls (~20 → 1 per batch)
```

### New Features (`feat()`)

```bash
feat(console.log):
- improved console.log() replacement to Logger.log() in FUNCTIONS.gs
- enhanced Logging documentation with examples

feat(doExportProventos):
- added doExportProventos() function
- export proventos to target spreadsheet
```

### Version Bumps (`bump()`)

```bash
bump: 
- version bump 17.0.5 - significant changes for new release

commit(source):
- just committing changes to the source
- no changelog prepared, bump to trigger version update
```

## 🎯 File Naming Patterns

### Single File Fixes

```
fix(- FILENAME.gs):
- specific change description
```

### Multi-File Changes

```
refactor(file1, file2):
- description of changes across files
```

### Function-Specific Changes

```
fix(functionName):
- describes what was fixed in the function
```

## 📊 Code Patterns to Describe

### Before/After Transformations

```
refactor(export pipeline):
- Refactor Basic save/edit/export pipeline with config-driven structure
- Introduced basicExportLookup configuration map
```

### Performance Improvements

```
perf(save):
- Cached sheet metadata (LR, LC) once instead of recalculating multiple times
- Reduced API calls from ~9 per execution to 1–2
```

### Dead Code Removal

```
clean(functions):
- Delete getValuesSafe — defined but never called anywhere
- Inline getConfigValue(DBG) directly in LogDebug, removing the runSafely wrapper
```

### Configuration Changes

```
feat(config):
- Applied getConfigValue to functions where new functionality can be used
- Improved parameter validation with fallback defaults
```

## 🔄 Common Change Categories

| Category | When to Use | Example |
|----------|-------------|---------|
| `fix()` | Bug fixes, null safety, logic corrections | `fix(- Save - Process.gs)` |
| `refactor()` | Code reorganization, cleanup, optimization | `refactor(LogDebug)` |
| `feat()` | New functionality added | `feat(doExportProventos)` |
| `perf()` | Performance improvements | `perf(import pipeline)` |
| `clean()` | Dead code removal | `clean(functions)` |
| `bump()` | Version updates | `bump: version 17.0.5` |
| `revert()` | Reverting changes | `revert(save/edit)` |

## 📝 Bullet Point Best Practices

### ✅ Good:
- Shows **what** changed in the file
- Includes **function names** modified
- Mentions **code patterns** used
- States **intent** and **impact**
- Uses **specific numbers** (e.g., `~20 → 1 API call`)

### ❌ Avoid:
- Generic "misc fixes" 
- Vague "small changes"
- Without file/function context
- Unclear intent ("improve code")

## 🎨 Examples Organized by File Type

### Spreadsheet Files (.gs)

```bash
fix(- Data - Process.gs):
- small fix to complete move code to helper functions

feat(dates):
- created helper functions for date handling
- extractAndValidateDates validates and extracts dates from ranges
```

### Template Files

```bash
fix(EXPORT & Template.gs):
- Removed FUTURE_1, FUTURE_2, FUTURE_3 from basicExportMap third entry
- export only for sheets with C2/E2/G2 checks (FUND) not B2/C2
```

### Helper Functions

```bash
refactor(doCheckDATA):
- Remove _CONFIG_VALUE_CACHE entirely — config/settings values change
  during a run making caching unsafe and causing stale reads
- Reduce unnecessary getSheet() calls from ~9 per execution to 1–2
```

## 📈 Real Commit History Examples (Recent Commits)

```bash
$ git log --oneline -10

feat(console.log): improved console.log() replacement to Logger.log()
fix(- Edit - Process.gs): processEditGeneric now reads rows safely
refactor(LogDebug): Added _DBG_FETCHING flag to prevent race condition
bump: version bump 17.0.5 - significant changes for new release
fix(save/export): removed dead checks and reduce spreadsheet API calls
```

## 🚀 Quick Reference Commands

### Viewing Your Commit History

```bash
# Recent commits with meaningful messages
git log --oneline | Select-Object -First 20

# See commit subjects
git log --format="%h %s" | Select-Object -First 30

# Full diff of latest commit
git show HEAD
```

### Checking Current Branch

```bash
git branch --show-current

# Check if on main or develop
```

## 📋 Template for Writing New Commit Messages

Use this template when making changes:

```
category(FILE): 
- one-line summary of what changed
- bullet point explaining the change
- code pattern used (if applicable)
- impact or improvement achieved (optional)
```

### Example Usage

```bash
git commit -m "refactor(import):
- Changed import function to use new data structure
- Improved error handling with try-catch blocks
- Reduced memory usage by 15% on large imports"
```

## 🎯 Key Takeaways from B3Assets Commit History

### What We've Learned:

1. **Be Specific** - Mention exact files and functions changed
2. **Show Impact** - Include performance metrics, API call counts
3. **Use Code Patterns** - Describe transformations (e.g., `=== → ==null`)
4. **Preserve Context** - Keep version tags meaningful (e.g., 17.1.x)
5. **Document Intent** - Explain WHY changes were made, not just WHAT

### Avoid These Common Mistakes:

- ❌ "misc fixes" → Use `fix(FILE):` with specifics
- ❌ "small improvements" → Describe what was improved
- ❌ "minor updates" → Specify what changed
- ❌ "apply fixes" → Show what was fixed and why
- ❌ Unclear version bumps → Include bump reason/context

## 📚 Additional Resources

- See [COMMIT-MESSAGE-COMPARISON.txt](COMMIT-MESSAGE-COMPARISON.txt) for side-by-side comparisons
- Check [COMMIT-AMEND-COMMANDS.txt](COMMIT-AMEND-COMMANDS.txt) for all 100 amended messages
- Reference agent repo: https://github.com/brunogcar/agent/commits/main/

---

Generated from analysis of B3Assets git history (100+ commits amended with meaningful messages)
Date: $(date +%Y-%m-%d)
