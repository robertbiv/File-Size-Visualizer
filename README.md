# File Size Filter

A simple Windows desktop app to select a folder and visualize the size distribution of files and subfolders using a pie chart. Includes filtering by minimum file size and optionally applies to subfolders.

## Features
- Select a folder
- Compute sizes of files and folders (recursive)
- Filter items by minimum size (e.g., show only items larger than X MB)
- Toggle whether filtering applies within subfolders
- Pie chart visualization of size distribution

## Requirements
- Python 3.9+
- Packages: matplotlib, pillow, humanize (optional)

## Quick Start

```pwsh
# From the project folder
python -m pip install -r requirements.txt
python app.py
```

## Notes
- Large folders may take time to scan; progress is shown in the status bar.
- System or permission-restricted paths may be skipped.
