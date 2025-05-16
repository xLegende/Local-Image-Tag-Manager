# Local Image Tag Manager

A browser-based interface for managing image datasets with corresponding tag files (e.g., `.txt` files), commonly used for preparing AI/ML training data. This tool runs entirely **locally** in your browser using the File System Access API, meaning your data stays on your machine.

It focuses on providing efficient ways to view, filter, search, and edit the tags associated with your images.

## Features

*   **Local Operation:** Runs entirely in the browser. No server needed. Requires a browser supporting the File System Access API (like Chrome, Edge).
*   **Dataset Loading:** Load image/tag pairs from a local folder, including subdirectories recursively.
*   **Image Grid View:** Displays images with a preview of their tags.
*   **Tag Management:**
    *   **Single Image Editor:** Edit tags in a modal view with image preview and metadata.
    *   **Quick Tag Delete:** Remove tags directly from the grid preview (for non-trashed items).
    *   **Tag Suggestions:** Autocomplete suggestions based on existing tags in the dataset.
*   **Filtering & Searching:**
    *   Filter by tags to include (contains logic, AND/OR modes).
    *   Filter by tags to exclude (exact match).
    *   Filter by status (All Active, Tagged, Untagged, Modified, **Trashed**).
    *   Search tag frequency list (based on active images).
*   **Bulk Operations (on Active/Non-Trashed Images):**
    *   Rename tags across the entire active dataset.
    *   Remove specific tags from all active files.
    *   Add tags to all active files.
    *   Remove duplicate tags within each active file.
*   **Multi-Select Batch Editing (for Active/Non-Trashed Items):**
    *   Select multiple images (Ctrl/Cmd+Click, Shift+Click).
    *   Batch add/remove tags to the selection.
    *   View common/differing tags in the selection.
    *   **Move selected items to trash.**
*   **Trash Can System:**
    *   Move images (and their tag files) to a temporary "trash" state instead of immediate deletion.
    *   View trashed items separately.
    *   Restore items from trash.
    *   **Permanently empty the trash**, deleting files from disk.
*   **Custom Formatting Rules:**
    *   Define custom find/replace rules (including regex).
    *   Save rules locally.
    *   Apply rules in bulk to the active dataset.
*   **Utilities (primarily on Active/Non-Trashed Images):**
    *   Find files with identical tag sets.
    *   Find orphaned tag files or images missing tags (considers entire loaded structure).
    *   Find tags with inconsistent capitalization.
    *   Export a list of all unique tags from active images.
*   **Display Options:** Sort the image grid by path, filename, tag count, date, size, dimensions.
*   **Usability:**
    *   Dark Mode support.
    *   Keyboard navigation for grid and editor.
    *   Save indicators for tag modifications and a master "Save All Changes" button.
    *   "Empty Trash" button to finalize deletions.
    *   Status messages and loading indicators.

## Getting Started

1.  **Download:** Get the `index.html`, `style.css`, and `script.js` files.
2.  **Open:** Open the `index.html` file in a compatible browser (e.g., Google Chrome, Microsoft Edge).
3.  **Load:** Click the "Load Dataset Folder" button and select the root folder containing your images and tag files. Grant the necessary read/write permissions when prompted.
4.  **Manage:** Use the sidebar controls and the image grid to manage your tags and items.
    *   To remove an item, click its '×' button to move it to the trash.
    *   View trashed items using the filter.
    *   Restore items or click "Empty Trash" in the header to permanently delete them from your disk.
5.  **Save Tag Changes:** Click "Save All Changes" to write tag modifications back to your local tag files for non-trashed items.

## 📜 License

This project is distributed under the Apache License 2.0. See `LICENSE` file for more information.
