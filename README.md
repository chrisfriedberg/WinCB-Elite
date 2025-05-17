# WinCB-Elite
Advanced Windows Clipboard Manager with Merge Buffer, Tray Control, and Text/Image History — WinCB Elite

# WinCB-Elite Help File
*December 5, 2023*

Contact: Chris Friedberg  
312-785-0044  
Elgin, IL

## Application Functionality

WinCB-Elite is a powerful clipboard manager for Windows that offers the following features:

### Core Features
- **Clipboard History**: Automatically tracks and stores text and image clipboard content
- **Tag System**: Add custom color-coded tags to organize clips
- **Search**: Full-text search through your clipboard history
- **Clip Groups**: Save and restore groups of clips for different projects
- **Keyboard Navigation**: Navigate through clips with keyboard shortcuts

### Clip Management
- **Mixed Content Support**: Handle text, images, and combinations
- **Rich Text Editing**: Edit text clips directly within the application
- **Batch Export**: Export clips to readable text files, including tag information
- **Delete Clips**: Remove individual clips or clear entire history
- **Clip Filtering**: Filter clips by tag or search text

### User Interface
- **Custom Theme**: Dark theme with customizable elements
- **System Tray Integration**: Run in background with tray icon access
- **Resizable Interface**: Adjust window size to your needs
- **Tooltips**: Helpful tips on hover
- **Status Indicators**: Shows current status and operation results

### System Integration
- **Auto-Start Option**: Launch with Windows
- **Hotkey Support**: Global shortcuts for quick access
- **Pause/Resume**: Temporarily pause clipboard monitoring
- **Low Resource Usage**: Optimized for minimal system impact

## Compilation Requirements

To compile WinCB-Elite from source, you'll need:

### File Structure
The application consists of the following key files:
- `WinCB-Elite.pyw` - Main application file
- `icon.ico` - Default application icon (must be in same directory as main file)
- `APP_DATA_DIR` - Will be automatically created at `C:\Users\[Username]\WinCB-Elite`

### Custom Icons
- When using the "Change App Icon" button, the selected icon file remains in its original location (not copied)
- The application stores only a reference to the icon file's path
- For non-ICO images (PNG, JPG, etc.), a converted temporary copy is created in the application data directory
- If the original icon file is moved or deleted, the application will revert to the default icon

### Compilation Steps
1. Ensure all dependencies (listed below) are installed
2. Run a Python compiler like PyInstaller:
   ```
   pyinstaller --onefile --windowed --icon=icon.ico WinCB-Elite.pyw
   ```
3. The compiled executable will be in the `dist` directory

## Dependencies

WinCB-Elite requires the following dependencies:

### Python Version
- Python 3.8 or higher

### External Libraries
- `customtkinter` - For the modern UI components
- `Pillow` (PIL) - For image processing 
- `pywin32` - For Windows system integration and clipboard access
- `pathlib` - For path manipulation
- `keyboard` - For global hotkey support

### Installation
Dependencies can be installed via pip:
```
pip install customtkinter Pillow pywin32 pathlib keyboard
```

### Automatic Dependency Installation
The application does NOT automatically install its dependencies. Users or developers must install the required packages manually before running or compiling the application.

## Additional Notes

- Application data is stored in `C:\Users\[Username]\WinCB-Elite`
- Image clips are stored in memory and serialized to JSON for persistence
- The app creates backup files of your clipboard history automatically
- For technical support or feature requests, contact information is provided at the top of this document

### Auto-created Files and Directories
When WinCB-Elite runs for the first time, it automatically creates the following:

- Main application directory: `C:\Users\[Username]\WinCB-Elite`
- Configuration file: `wincb-elite_config.json` - Stores app settings and preferences
- History file: `wincb-elite_history.json` - Stores all clipboard entries
- Batch exports directory: `batchoutputs` - For storing exported clip groups
- Tag colors file: `tag_colors.json` - Maps tags to their display colors

If you manually create the application directory, the app will automatically create these files with default values when it first runs. You do not need to create these files yourself.

---

© 2023 Chris Friedberg
