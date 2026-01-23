# Microsoft Office 2024 LTSC — Install Guide

This document provides instructions for downloading, configuring, and installing Microsoft Office 2024 LTSC using either a customizable batch installer or the official Microsoft deployment method.

---

## Requirements

Before proceeding, download the official **Office Deployment Tool (ODT)**:

**Download:**  
https://learn.microsoft.com/es-es/office/ltsc/2024/deploy#download-the-office-deployment-tool-from-the-microsoft-download-center

Running the tool will extract the following files:

- `setup.exe`
- `configuration.xml`

Both are required for installation.

---

## Option 1 — Customizable Batch Installer

This method allows you to interactively select which Office applications to install.

1. Place **Setup Office 2024 LTSC (Run As Admin).bat** in the same directory as ***Setup.exe***.
2. Run the batch file **as Administrator**.
3. Select the applications to install by entering the corresponding number.  
   - Each application indicates its status as  **=1 (Enabled)** or  **=0 (Disabled).**
4. Enter **S** to begin the installation.

This option is recommended for users who want a simple, menu‑based installation process.

---

## Option 2 — Official Microsoft Method (Advanced)

This method uses Microsoft’s official configuration and deployment workflow.

1. Generate a custom XML configuration using the online tool:
   **Configuration Tool:**  
   https://config.office.com/deploymentsettings

2. Save the generated XML file in the same folder as `setup.exe`.

3. Open **Command Prompt as Administrator** in that folder.

4. Download the Office installation files:  
   `setup.exe /download configuration.xml`

5. Once the download completes, install Office:  
   `setup.exe /configure configuration.xml`

This option is recommended for IT administrators or advanced users who need full control over deployment settings.

---
