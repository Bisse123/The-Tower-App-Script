# Deployment Archive Scripts

This directory contains several scripts to help you automatically archive all current deployments for your Google Apps Script project.

## Available Scripts

### 1. PowerShell Script (`archive-deployments.ps1`)
**Best for:** Windows users who want a comprehensive, interactive archival process.

**Features:**
- Lists all current deployments
- Archives (deletes) all deployments one by one

**Usage:**
```powershell
# Run from the project root directory
.\archive-deployments.ps1
```

### 2. Batch Script (`archive-deployments.bat`)
**Best for:** Windows users who want a simple double-click solution.

**Features:**
- Simple wrapper for the PowerShell script
- Double-click to run
- Automatically handles execution policy

**Usage:**
- Double-click the file, or
- Run from command prompt: `archive-deployments.bat`

## Prerequisites

### For PowerShell/Batch Scripts:
1. **Google Apps Script CLI (clasp)** must be installed:
   ```bash
   npm install -g @google/clasp
   ```

2. **Authentication** - You must be logged in to clasp:
   ```bash
   clasp login
   ```

3. **Project Setup** - Run from the directory containing `.clasp.json`

### For Apps Script Function:
1. **Apps Script API** must be enabled in your Google Cloud Console
2. **Required OAuth Scopes** - Add these to your `appsscript.json`:
   ```json
   {
     "oauthScopes": [
       "https://www.googleapis.com/auth/script.projects",
       "https://www.googleapis.com/auth/script.deployments"
     ]
   }
   ```

## What Gets Archived

When you "archive" deployments, you're actually **deleting** them. This means:
- The deployment URLs will no longer work
- Users won't be able to access the web app through old URLs
- The deployment history is removed

**Important:** The `@HEAD` deployment (development version) is automatically excluded from archiving and will remain untouched.

## Recommended Workflow

1. **Before archiving:**
   - Document current deployment URLs if needed
   - Notify users of any upcoming changes
   - Test your latest code thoroughly

2. **Archive process:**
   - Use one of the provided scripts
   - Review the logs to ensure all deployments were archived successfully

3. **After archiving:**
   - Create a new deployment with a descriptive name
   - Update any bookmarks or links to use the new URL
   - Test the new deployment thoroughly

## Troubleshooting

### Common Issues:

1. **"clasp not found"**
   - Install clasp: `npm install -g @google/clasp`

2. **"Not logged in"**
   - Run: `clasp login`

3. **"Permission denied"**
   - Make sure you have owner/editor access to the Apps Script project

4. **"Script not found"**
   - Check that `.clasp.json` exists and contains the correct `scriptId`

### PowerShell Execution Policy Issues:
If you get execution policy errors, run PowerShell as administrator and execute:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## File Structure After Running

After running the archive script, you'll find:
```
archived-deployments/
├── archive-2025-07-04_14-30-15/
│   ├── deployments-before-archive.txt
│   ├── deployments-after-archive.txt
│   ├── archived-deployments.log
│   └── archive-errors.log (if any errors occurred)
```

## Security Notes

- These scripts only work with projects you own or have editor access to
- The Apps Script API requires appropriate OAuth permissions
- Always test in a non-production environment first
- Keep backups of your code before making deployment changes

## Integration with CI/CD

You can integrate these scripts into your continuous deployment pipeline:
```yaml
# Example GitHub Actions step
- name: Archive old deployments
  run: |
    npm install -g @google/clasp
    echo "${{ secrets.CLASP_CREDENTIALS }}" > ~/.clasprc.json
    ./archive-deployments.ps1
```
