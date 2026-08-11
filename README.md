# 🔄 Offline Bulk PowerPoint to PDF Converter

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![UI](https://img.shields.io/badge/UI-CustomTkinter-brightgreen)

A modern, fast, and completely offline desktop application to bulk convert Microsoft PowerPoint presentations into PDF documents. Built with Python and `customtkinter`, it offers a sleek user interface while leveraging Windows COM objects for reliable conversion without relying on external APIs or cloud services.

## ✨ Features

- **Bulk Conversion**: Convert entire folders of PowerPoint files to PDF in one go.
- **Modern UI**: Clean, responsive, and system-theme-aware interface built with `customtkinter`.
- **Fast & Reliable**: Uses Microsoft PowerPoint's native COM interface for 100% accurate rendering and formatting.
- **Offline & Secure**: No internet connection required. Your sensitive files never leave your computer.
- **Progress Tracking**: Real-time progress bar and detailed status updates.
- **Multi-threaded**: Background processing ensures the UI never freezes during heavy conversions.
- **Broad Format Support**: Handles `.ppt`, `.pptx`, `.pptm`, `.pps`, `.ppsx`, `.ppsm`, `.pot`, `.potx`, and `.potm`.

## 🛠️ Prerequisites

To run this application, you must have the following installed on your machine:
- **Windows OS**
- **Microsoft PowerPoint** (Must be installed locally, as the app relies on the PowerPoint COM interface).
- **Python 3.8+**

## 🚀 Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/deshanWimalasooriya/Offline-File-Converter.git
   cd Offline-File-Converter
   ```

2. **Set up a virtual environment (optional but recommended):**
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   ```

3. **Install the required dependencies:**
   ```bash
   pip install customtkinter comtypes
   ```
   *(Note: A pre-built executable might be available in the `dist/` folder if you wish to run it without Python installed.)*

## 🎮 How to Use

1. Launch the application:
   ```bash
   python converter.py
   ```
2. Click **1. Select Source Folder** to select the directory containing your PowerPoint presentations.
3. Click **2. Select Output Folder** to choose the destination where the generated PDFs should be saved.
4. Click **3. START CONVERSION** to begin! You can monitor the live progress at the bottom of the window.

## ⚙️ Under the Hood

The application spawns a background, headless instance of Microsoft PowerPoint using the `comtypes` library. It iteratively loads each presentation file from your selected source folder and natively exports them as PDFs (`format type 32`). This approach ensures that the layout, custom fonts, graphics, and transitions are preserved exactly as they appear in PowerPoint.

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to open a pull request or check the issues page if you want to contribute to the project.

## 📝 License

This project is open-source and available under the MIT License.