# PowerShell Script Manager

This is a desktop application built with Python and tkinter that helps manage and organize PowerShell scripts across multiple folders.

<img width="1127" height="860" alt="image" src="https://github.com/user-attachments/assets/fe9f47d9-bcfe-4bbe-8362-59402807bafe" />

### Todo

- [ ] Dark Mode
- [ ] Module View 
  - [ ] List installed modules
  - [ ] Check if updates are available and provide update option

### In Progress

- [ ] Add Modern UI

### Done ✓

- [x] Create MVP


## Requirements

- Python 3.x (tkinter is included in the standard Python distribution)

## Running the Application

To run the application, simply execute:

```powershell
python main.py
```

## Features

- Three-tab interface: Home (Scripts), PowerShell, and Folders
- Browse and manage folders containing PowerShell scripts
- Recursive script discovery in selected folders
- Persistent storage of folder selections
- List view of all discovered PowerShell scripts (.ps1 files)
- Automatic refresh when adding or removing folders

## Usage

1. Go to the "Folders" tab and click "Add Folder" to select folders containing PowerShell scripts
2. The application will automatically scan these folders for .ps1 files
3. View all discovered scripts in the "Home (Scripts)" tab
4. Use the "Refresh Scripts" button to update the list
5. Remove folders using the "Remove Selected Folder" button in the Folders tab

