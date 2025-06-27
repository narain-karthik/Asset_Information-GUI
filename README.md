
# System Asset Information Tool

A **professional desktop application** built with Python and Tkinter that retrieves detailed hardware and system information of a Windows machine. The application allows users to view system information in a user-friendly interface and export the results to a professionally formatted PDF report.

---

## 🖥️ Features

✅ Retrieve key system information:  
• **System Name**  
• **IP Address**  
• **RAM Size**  
• **CPU Model & Details**  
• **CPU Tag Number**  
• **Disk Information (Physical Drives)**  
• **Monitor Details (Model & Serial)**  
• **Operating System Name & Status**  

✅ **Export system details to a professionally styled PDF report**  

✅ **Simple, user-friendly GUI using Tkinter**  

✅ **No external dependencies apart from standard libraries and `reportlab`**  

---

## 📂 Project Structure

```
SystemAssetInfo/
│
├── system_asset_info.py       # Main application code
├── README.md                  # Project documentation
└── requirements.txt           # Python dependencies
```

---

## ⚡ Requirements

- Python **3.8** or higher  
- **Windows OS** (Uses WMIC and PowerShell commands)  
- Required Python packages:

```
reportlab
psutil
```

You can install dependencies using:

```bash
pip install -r requirements.txt
```

---

## 🛠️ How to Run

1. Make sure you have Python installed.  
2. Install the required libraries using the command above.  
3. Run the application:

```bash
python system_asset_info.py
```

---

## 📑 PDF Report Example

The exported PDF report contains:  
- System name, IP address, RAM, CPU details, and OS info  
- Physical disk model and size  
- Monitor model and serial numbers  
- Clean, professional design with company branding  

---

## ⚙️ Technical Notes

- Disk and CPU details are fetched using **WMIC commands**, hence **requires Windows OS**.  
- Monitor details use **PowerShell WMI commands**, ensure PowerShell is available.  
- The tool filters out virtual disks from the disk information.  
- The GUI is designed using Tkinter's `ttk` themed widgets for a modern look.  

---

## ⚡ Screenshots

![System Asset Info GUI Screenshot](screenshot.png)  
*Example interface of the System Asset Info Tool*

---

## 📄 License

© 2025 System Asset Info | IT Department  
This project is intended for **internal IT asset management use**. Redistribution or commercial use without permission is prohibited.

---

## 🏢 Contact

For feature requests or support, please contact the **IT Department**.

---

*Professional System Information Reporting Made Simple.*  
