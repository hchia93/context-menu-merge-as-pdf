# Context Menu Merge as PDF

A lightweight Windows utility that adds right-click context menu options to merge PDF files and images into a single PDF document. Supports selective page extraction and batch conversion from folders.

Supported format: `pdf`, `png`, `jpeg`, `jpg`.

## Feature

- **Right-click Context Menu Integration** : Merge files directly from Windows Explorer
- **Common format Support** : Works with `PDF`, `PNG`, `JPEG`, and `JPG` files
- **Selective Page Extraction** : Choose specific pages or page ranges from PDF files
- **Batch Processing** : Convert entire folders of documents into a single PDF
- **Smart Auto-detection**: Handles mixed selections of files and folders intelligently
- **Zero Configuration** : Auto-generates output in sensible locations

## Installation

- Install Python 3.6+
- Install dependencies by running this file in terminal.

### Setup

1. Clone or download this repository

    ```bash
    git clone https://github.com/yourusername/context-menu-merge-as-pdf.git
    cd context-menu-merge-as-pdf
    ```

2. Install dependencies by running the following in terminal. This should include registration to context menu.

    ```bash
    install.bat
    ```

3. Repair context menu related with `Admin` right if necessarily.

    ```bash
    repair-context.bat
    ```

### Verification

After installation, you should see:


```bash
"Merge PDF (files)" when right-clicking on files
"Merge PDF (folder)" when right-clicking on folders
```

## Command Line

### Usage

#### 1. Merge with Page Selection

```bash
python merge-pdf.py -merge "document.pdf [1,3:5]" "image.png" -o "output.pdf"
```

This merges:

Page 1 from `document.pdf` + Pages 3-5 from `document.pdf` + `image.png` into `"output.pdf"`

#### 2. Batch Merge

```bash
python merge-pdf.py -blob "C:\Documents\MyFolder" -o "output.pdf"
```

Merges all supported files in the folder alphabetically.

#### 3. Auto

```bash
python merge-pdf.py -auto "C:\path\to\folder"
python merge-pdf.py -auto "file1.pdf" "file2.png"
```

Automatically detects whether inputs are files or folders and handles them appropriately. This is the mode mainly used by the context menu.

### Reference

| Command | Description | Requires -o |
|---------|--|--|
| `-merge`  | Merge specific files with optional page selection | Yes |
| `-blob`  | Merge all files in a directory alphabetically | Yes |
| `-auto` | Auto-detect and merge files/folders | No |

> All of the above uses absolute file path or absolute directory path.

## Troubleshooting

1. Context menu not appearing
    - Ensure you ran `repair-context.bat` as Administrator
    - Restart Windows Explorer (Task Manager → Windows Explorer → Restart)

2. Script crashes immediately
    - Check merge-pdf-debug.log in the script directory for error details
    - Ensure Python and all dependencies are properly installed
