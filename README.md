# context-menu-merge-as-pdf

Allow user to merge documents or convert folder into a single output PDF.

Supported format: `pdf`, `png`, `jpeg`, `jpg`.

# Feature
For non-auto mode:
- Merge multiple files into one `pdf`
- Extract certain pages from source `pdf`

## Installation

- Install Python 3.6+
- Install dependencies by running this file in terminal.

```bat
install.bat
```

## Context Menu Installation
- To register, execute this file with Admin right
```bash
register.bat
```

- To unregister, execute this file with Admin right
```bash
unregister.bat
```


## Command

| Command | Description | 
|---------|--|
| `-merge`  | Merge multiple files into one pdf. Optionally support page selection and page range selection for `pdf` files. |
| `-blob `  | Merge every supported files into one pdf by specifiying the source directory. Merged content are ordered alphabetically. | 
| `-auto` | Convert or Create selected files / folder as pdf. |

> All of the above uses absolute file path or absolute directory path.

## Example
`-merge`  

Format:
### 
```cmd
python merge-pdf.py -merge <file1> <file2> -o <outputfile>
```
Use case:
```cmd
python merge-pdf.py -merge "example.png [1, 3:5]" -o "../output.pdf"
```
The target file will have page 1, page 3 to page5.

---

`-blob` 

Use case:
```cmd
python merge-pdf.py -blob <directory> -o <outputfile>
```

`-auto`

```cmd
python merge-pdf.py -auto "E:\work\myfolder"
```
To merge several specific files into a single PDF (output will be 'merged.pdf' in the same directory):
```cmd
python merge-pdf.py -auto "E:\a.pdf" "E:\b.pdf"
```
