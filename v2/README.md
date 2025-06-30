# Asset_Info-v2

![Asset Info Banner](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![Python](https://img.shields.io/badge/Python-3.8+-3776AB?logo=python)
![GUI](https://img.shields.io/badge/GUI-Tkinter-ff69b4)
![PDF Export](https://img.shields.io/badge/Export-PDF-4caf50)
![License](https://img.shields.io/badge/License-Custom-lightgrey)

---

<p align="center">
  <img src="Logo.png" width="80%" alt="Asset_Info-v2 Preview"/>
</p>

---

## ✨ Overview

**Asset_Info-v2** is a professional, modern, and easy-to-use Windows application that collects detailed asset information from your PC or laptop. With just a click, retrieve comprehensive hardware and system specs—including RAM, CPU, drives, product keys, and monitor details—and export them as a visually appealing PDF report.  
Built for IT professionals, helpdesks, and asset managers.

---

## 🎨 Features

- **One-Click System Inventory:**  
  Instantly displays PC/laptop specs such as system name, IP, RAM, CPU, OS, serial number, product key, SSD/HDD, and monitors.
- **Modern GUI:**  
  Sleek interface with copyable fields, styled tables, and a consistent color theme.
- **Copy & Share:**  
  All fields, including disk and monitor info, are easily copyable to clipboard.
- **PDF Export:**  
  Generate a professional, branded PDF asset report with custom table styles.
- **No Admin Required:**  
  Works for standard Windows users (most queries).
- **Built-in Error Handling:**  
  Friendly alerts for missing or unavailable data.

---

## 🖥️ Interface Preview

<p align="center">
  <img src="https://user-images.githubusercontent.com/your-repo-asset-info-gui-screenshot.png" width="80%" alt="Asset_Info-v2 UI"/>
</p>

---

## ⚙️ Requirements

- **OS:** Windows 10/11
- **Python:** 3.8 or newer
- **PowerShell:** Required (default on Windows)
- **Dependencies:**
  - [psutil](https://pypi.org/project/psutil/)
  - [reportlab](https://pypi.org/project/reportlab/)

Install required packages:
```sh
pip install psutil reportlab
```

---

## 🚀 Getting Started

1. **Clone or Download** this repository.
2. **Install dependencies** (see above).
3. **Run**:
   ```sh
   python Asset_Info-v2.py
   ```
4. Click **Get System Info**.
5. Use **Copy** buttons or fields to grab data, or **Export as PDF**.

---

## 📝 PDF Export Sample

<p align="center">
  <img src="https://user-images.githubusercontent.com/your-repo-asset-info-gui-pdf.png" width="80%" alt="Asset_Info-v2 PDF Export"/>
</p>

---

## 🛠️ Build as Standalone EXE

You can create a Windows executable using [PyInstaller](https://pyinstaller.org/):

```sh
pyinstaller --onefile --windowed --icon=youricon.ico Asset_Info-v2.py
```
- The `.exe` will be in the `dist` folder.
- **Custom icon**: use the `--icon` flag with a `.ico` file (optional).

---

## 📋 How It Works

- Uses **PowerShell** and **WMI** for advanced Windows-specific queries (disk, product key, monitor serials).
- Uses **psutil** for RAM and resource data.
- Presents results in a **Tkinter** GUI with modern styling.
- Exports data to a stylish PDF via **ReportLab**.

---

## 💡 Notes

- Some info (e.g., Product Key, Serial Number) may require hardware/firmware or OS support.
- Cross-platform use is not supported; Windows only.
- For best results, run on a physical machine (some info may be missing in VMs).

---

## 🧑‍💻 Author & License

**© 2025 System Asset Info | IT Department**  
*See LICENSE for details.*

---

<p align="center">
  <i>Created with ❤️ for IT and Asset Management.</i>
</p>
