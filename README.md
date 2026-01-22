# docx-deshittify

A Python script to fix `.docx` files with phantom, irremovable blank pages, tables, paragraph and space formatting issues, particularly after export/import between google docs and office365.
Edits XML to remove problematic elements that graphical editors gaslight you into thinking aren't there. SPOILER ALERT: they are.


## Why?

- **Phantom blank pages** - Extra pages at the beginning or end of doc, that cannot be deleted
- **Excessive paragraph spacing** - Invisible spacing that pushes content to new pages
- **Forced page breaks** - Unwanted page break elements that persist after deletion in GUI editors
- **Table formatting issues** - Table properties that force unwanted page breaks
- **Metadata persistence** - Hidden formatting data that survives GUI deletion



## I no read. I want use. NOW:

```bash
python3 docx-deshittify.py input.docx output.docx
```
command works with relative and abs paths. desloptimization has never been easier.



## Requirements

- **Python 3.6+**
- **python-docx** lib



## Req Installation

#### Fedora/RHEL:

```bash
sudo dnf install python3 python3-pip
pip3 install --user python-docx
```

#### Debian-like:

```bash
sudo apt update
sudo apt install python3 python3-pip
pip3 install --user python-docx
```

#### Arch:

```bash
sudo pacman -S python python-pip
pip install --user python-docx
```

### Permission errors

If it's not working it's probably mandatory access control, edit the file perms:

```bash
chmod +x docx-deshittify.py

```

