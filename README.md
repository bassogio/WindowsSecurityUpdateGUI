
# Windows Security Update GUI

A PyQt5-based desktop application for managing and scheduling Windows security updates on remote machines.

## Features

-   Display and edit a list of machines to update
    
-   Detect remote operating systems and disk space usage via PowerShell
    
-   Schedule updates with recurrence options (weekly/monthly/yearly)
    
-   Visual help center for user guidance
    
-   Progress bar for disk space display
    
-   Built-in table editor with add/remove/reorder functionality
    

----------

## Requirements

-   Python 3.x
    
-   Windows OS with PowerShell
    
-   Remote PowerShell access enabled on target machines
    
-   Required Python packages:
    
    ```bash
    pip install pyqt5
    ```

----------

## Running the Script Locally

You can run the script directly using Python:

```bash
python gui.py
```

Ensure the following UI files are in the same directory as `gui.py`:

-   `WindowsSecurityUpdateMainWindow.ui`
    
-   `ScheduleUpdates.ui`
    

----------

## Building the Executable with PyInstaller

### Step 1: Install PyInstaller

Open a terminal or PowerShell window and run:

```bash
pip install pyinstaller
```

### Step 2: Build the Executable

Use the following command to generate a standalone `.exe`:

```bash
pyinstaller gui.py --noconfirm --onefile --windowed `
--add-data "WindowsSecurityUpdateMainWindow.ui;." `
--add-data "ScheduleUpdates.ui;."
```
**Note for CMD users**: If you're using CMD (instead of PowerShell), replace backticks with carets (`^`) and use colons (`:`) instead of semicolons (`;`) for` --add-data`:

```cmd
pyinstaller gui.py --noconfirm --onefile --windowed ^
--add-data "WindowsSecurityUpdateMainWindow.ui:." ^
--add-data "ScheduleUpdates.ui:."
```

### Step 3: Post-Build Instructions

Once the executable is generated, PyInstaller will place it in the `dist/` folder. **You must manually copy the following `.ui` files into the same folder as the `.exe`** (e.g. `dist/` or wherever you move the `.exe`):

-   `WindowsSecurityUpdateMainWindow.ui`
    
-   `ScheduleUpdates.ui`
    
