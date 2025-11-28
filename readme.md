
# 🚀 PDF-Previewa - Fix your Windows PDF-Preview

A lightweight, portable Windows utility that repairs broken PDF previews in Windows Explorer.

----------

## ❗ The Problem

Downloaded PDF files often contain the _Mark of the Web_ (MotW), causing Windows to block the Preview Pane.

----------

## ✅ The Solution

This tool adds a context menu entry that:

-   Removes the MotW
        
-   Forces Windows Explorer to refresh the preview
    

All **without** renaming the file or timestamp.

----------

## ✨ How to Use

1.  Right-click a PDF file that shows no preview.
    
2.  Select **“Enable Preview”** (or your custom text from `config.ini`).
    
3.  The preview appears instantly.
    

----------

## 📦 Installation Options

The tool consists of just two files: `PDF-Previewa.exe` and `config.ini`. It acts as a portable installer/uninstaller.

### **1. Graphical User Interface (GUI)**

Simply double-click `PDF-Previewa.exe`.

-   **Install Current User** – Registers the context menu for the logged-in user (no admin rights required)
    
-   **Install All Users** – Registers system-wide (admin rights required)
    
-   **Uninstall** – Removes all related registry entries
    

### **2. Silent Deployment (CMD / SCCM / Intune)**

Ideal for system administrators.

**Examples:**
```dos
PDF-Previewa.exe /install_global "lang_en" /silent
```

```dos
PDF-Previewa.exe /install_user "lang_de" /silent
```

```dos
PDF-Previewa.exe /uninstall /silent
```

----------

## ⚙️ Parameter Reference

| Parameter              | Description                               | Admin Required? |
|------------------------|-------------------------------------------|------------------|
| `/install_user "ID"`   | Installs for the current user (HKCU).     | ❌ No            |
| `/install_global "ID"` | Installs system-wide (HKLM/HKCR).         | ✔️ Yes           |
| `/uninstall`           | Removes entries from HKCU and HKCR.       | ✔️ Partially*    |
| `/silent`              | Suppresses all success messages.          | –                |

\* Admin rights are required to fully remove global entries.


----------

## 📝 Configuration: config.ini

All text strings and languages can be freely defined.

```ini
; --- GERMAN CONFIGURATION ---
[lang_de]
Name=German
MenuText=Vorschau aktivieren
MsgSuccess=Erfolgreich installiert.
MsgUninstall=Erfolgreich entfernt.

; --- ENGLISH CONFIGURATION ---
[lang_en]
Name=English - United Kingdom
MenuText=Enable Preview (Fix)
MsgSuccess=Installed successfully.
MsgUninstall=Removed successfully.

; --- EXAMPLE: ADDING FRENCH ---
[lang_fr]
Name=French
MenuText=Activer l'aperçu
MsgSuccess=Installation réussie!
MsgUninstall=Supprimé avec succès.

```

----------

## 🔧 How It Works

When you click the context menu entry:

1.  **MotW Removal** – Removes the NTFS `Zone.Identifier` stream.
    
2.  **Shell Notification** – Sends `SHChangeNotify` to refresh the file preview.
    

Steps 2 invalidates the thumbnail cache → Explorer re-renders the preview immediately.

----------

## ⚠️ Compilation Note

Windows Explorer is 64-bit → the tool **must be compiled as x64**.

If compiled as x86, registry paths are redirected to `Wow6432Node` and the context menu entry won't appear.

----------
