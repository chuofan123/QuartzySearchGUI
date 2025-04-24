# Lab Inventory Search GUI

A simple desktop application to search exported lab inventory spreadsheets (e.g., from Quartzy).

## Features

*   Select your Excel inventory file.
*   Choose which sheet to search.
*   Select columns to search within (using 'contains' logic).
*   Select columns to display in the results table.
*   Simple graphical user interface (GUI).

## How to Use

There are two ways to use this application:

**Option 1: Windows Executable (Recommended - Easiest)**

1.  Go to the **[Releases Page](https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME/releases)** of this repository. (<- **IMPORTANT: Replace YOUR_USERNAME and YOUR_REPOSITORY_NAME with your actual details!**)
2.  Download the `InventorySearchGUI.exe` file from the latest release.
3.  **Important:** Your browser or Windows might warn you about downloading executable files. You may need to choose "Keep" or allow the download. Antivirus software might also flag the file (this is common for apps packaged with PyInstaller - it's likely a false positive if you downloaded directly from the link above).
4.  Double-click the downloaded `.exe` file to run the application. No installation is needed.
5.  The application will ask you to select your lab's Excel inventory file.

**Option 2: Running from Source Code (Requires Python)**

This method works on Windows, macOS, and Linux but requires you to have Python installed.

1.  **Prerequisites:** Make sure you have Python 3.8 or newer installed. You can download it from [python.org](https://www.python.org/).
2.  **Download Code:** Download the code from this repository (click the green "Code" button -> "Download ZIP") or clone it using Git.
3.  **Install Dependencies:** Open a terminal or command prompt, navigate (`cd`) into the downloaded code directory, and run:
    ```bash
    pip install -r requirements.txt
    ```
4.  **Run the App:** While still in the same directory in your terminal, run:
    ```bash
    python search_inventory_gui.py
    ```
5.  The application window should open, and it will ask you to select your lab's Excel inventory file.
