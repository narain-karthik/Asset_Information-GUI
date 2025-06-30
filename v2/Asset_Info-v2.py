import tkinter as tk
from tkinter import scrolledtext, ttk, messagebox, filedialog
import platform
import socket
import psutil
import subprocess
import json
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

POWERSHELL_PATH = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

def get_system_name():
    return platform.node()

def get_ip_address():
    try:
        hostname = socket.gethostname()
        return socket.gethostbyname(hostname)
    except:
        return "N/A"

def get_ram():
    ram_bytes = psutil.virtual_memory().total
    ram_gb = ram_bytes / (1024 ** 3)
    return f"{ram_gb:.2f} GB"

def get_serial_number():
    try:
        result = subprocess.check_output(
            [POWERSHELL_PATH, "-Command", "(Get-CimInstance Win32_BIOS).SerialNumber"],
            stderr=subprocess.STDOUT
        )
        return result.decode(errors="ignore").strip()
    except Exception as e:
        return f"Error: {e}"

def get_product_key():
    try:
        ps_cmd = "(Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey"
        result = subprocess.check_output(
            [POWERSHELL_PATH, "-Command", ps_cmd],
            stderr=subprocess.STDOUT
        )
        key = result.decode(errors="ignore").strip()
        if not key:
            return "Not Found"
        return key
    except Exception as e:
        return f"Error: {e}"

def get_disks_physical():
    try:
        ps_cmd = """
        $disks = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size
        if (-not $disks) {
            $disks = Get-CimInstance Win32_DiskDrive | Select-Object Model, Size
            foreach ($disk in $disks) { $disk | Add-Member -NotePropertyName MediaType -NotePropertyValue "Unknown" }
        }
        $disks | ConvertTo-Json
        """
        result = subprocess.check_output(
            [POWERSHELL_PATH, "-Command", ps_cmd],
            stderr=subprocess.STDOUT
        )
        disks = json.loads(result.decode(errors="ignore"))
        if isinstance(disks, dict):  # Only one disk
            disks = [disks]
        output = []
        for d in disks:
            model = d.get("FriendlyName") or d.get("Model") or "Unknown"
            size = d.get("Size")
            size_gb = f"{int(size) / (1024 ** 3):.2f} GB" if size and str(size).isdigit() else "Unknown Size"
            dtype = d.get("MediaType", "Unknown")
            output.append((model, size_gb, dtype))
        return output
    except Exception as e:
        return [("Error", str(e), "Unknown")]

def get_system_model():
    try:
        result = subprocess.check_output(
            [POWERSHELL_PATH, "-Command", "(Get-CimInstance Win32_ComputerSystem).Model"],
            stderr=subprocess.STDOUT
        )
        return result.decode(errors="ignore").strip()
    except Exception:
        return "Unknown Model"

def get_cpu_details():
    try:
        ps_cmd = """
        $cpu = Get-CimInstance Win32_Processor
        "$($cpu.Name), $($cpu.MaxClockSpeed) MHz, $($cpu.NumberOfCores) Core(s), $($cpu.NumberOfLogicalProcessors) Logical Processor(s)"
        """
        result = subprocess.check_output(
            [POWERSHELL_PATH, "-Command", ps_cmd],
            stderr=subprocess.STDOUT
        )
        return result.decode(errors="ignore").strip()
    except Exception as e:
        return f"Error: {e}"

def get_os_name():
    return f"{platform.system()} {platform.release()}"

def get_status():
    return "Active"

def get_monitor_tags():
    powershell_script = r"""
function Decode {
    param($data)
    if ($data -is [System.Array]) {
        return [System.Text.Encoding]::ASCII.GetString($data).Trim([char]0)
    }
    else {
        return "Not Found"
    }
}
$monitors = Get-WmiObject WmiMonitorID -Namespace root\wmi
$result = @()
foreach ($monitor in $monitors) {
    $name = (Decode $monitor.UserFriendlyName).Trim()
    $serial = (Decode $monitor.SerialNumberID).Trim()
    if ($name -and $serial) {
        $result += "Name: $name`nSerial: $serial"
    }
}
$result -join "`n`n"
"""
    try:
        result = subprocess.check_output(
            [POWERSHELL_PATH, "-Command", powershell_script],
            stderr=subprocess.STDOUT
        )
        return result.decode(errors="ignore").strip()
    except Exception as e:
        return f"Error: {e}"

def gather_info():
    info = {}
    info["System Name"] = get_system_name()
    info["IP Address"] = get_ip_address()
    info["RAM"] = get_ram()
    info["CPU Model Name"] = get_system_model()
    info["CPU Details"] = get_cpu_details()
    info["Serial Number"] = get_serial_number()
    info["Product Key"] = get_product_key()
    info["Monitor Details"] = get_monitor_tags()
    info["Disks"] = get_disks_physical()
    info["OS Name"] = get_os_name()
    info["OS Status"] = get_status()
    return info

def show_info():
    try:
        info = gather_info()
        # Main fields
        for key in field_labels:
            entry = field_labels[key]["value"]
            entry.config(state='normal')
            entry.delete(0, tk.END)
            entry.insert(0, info.get(key, ""))
            entry.config(state='readonly')
        # Disk info (showing model, size, type) as copyable text
        disk_text.config(state='normal')
        disk_text.delete(1.0, tk.END)
        disk_lines = []
        for model, size, dtype in info["Disks"]:
            disk_lines.append(f"Model: {model}    Size: {size}    Type: {dtype}")
        disk_text.insert(tk.END, "\n".join(disk_lines))
        disk_text.config(state='disabled')
        # Monitor info as copyable text
        monitor_text.config(state='normal')
        monitor_text.delete(1.0, tk.END)
        monitor_text.insert(tk.END, info["Monitor Details"])
        monitor_text.config(state='disabled')
        # Serial Number
        serial_entry.config(state='normal')
        serial_entry.delete(0, tk.END)
        serial_entry.insert(0, info["Serial Number"])
        serial_entry.config(state='readonly')
        # Product Key
        product_key_entry.config(state='normal')
        product_key_entry.delete(0, tk.END)
        product_key_entry.insert(0, info["Product Key"])
        product_key_entry.config(state='readonly')
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

def export_to_pdf():
    info = gather_info()
    export_fields = [
        ("System Name", info["System Name"]),
        ("IP Address", info["IP Address"]),
        ("RAM Size", info["RAM"]),
        ("CPU Model Name", info["CPU Model Name"]),
        ("CPU Details", info["CPU Details"]),
        ("Serial Number", info["Serial Number"]),
        ("Product Key", info["Product Key"]),
        ("OS Name", info["OS Name"]),
        ("OS Status", info["OS Status"]),
    ]
    default_filename = f"{info['System Name']}-Asset-Info.pdf"
    file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialfile=default_filename,
        filetypes=[("PDF files", "*.pdf")],
        title="Save As PDF"
    )
    if not file_path:
        return
    try:
        doc = SimpleDocTemplate(file_path, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        # Custom styles
        title_style = styles['Title']
        title_style.fontSize = 22
        title_style.alignment = 1  # Center
        section_heading = ParagraphStyle(
            name='SectionHeading',
            parent=styles['Heading3'],
            fontSize=13,
            leading=16,
            spaceBefore=14,
            spaceAfter=6,
            textColor=colors.HexColor("#1a237e"),
        )
        monitor_heading = ParagraphStyle(
            name='MonitorHeading',
            parent=styles['Heading4'],
            fontSize=11,
            leading=13,
            spaceBefore=10,
            italic=True,
            textColor=colors.HexColor("#222")
        )
        footer_style = ParagraphStyle(
            name='Footer',
            parent=styles['Normal'],
            fontSize=8,
            textColor=colors.HexColor("#888"),
            alignment=0,
            spaceBefore=24
        )
        elements = []
        # Title
        elements.append(Paragraph("System Asset Information Report", title_style))
        elements.append(Spacer(1, 20))
        # Info Table
        table_data = []
        for key, value in export_fields:
            table_data.append([
                Paragraph(f"<b>{key}</b>", styles['Normal']), Paragraph(str(value), styles['Normal'])
            ])
        info_table = Table(table_data, colWidths=[175, 305])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8faff")),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#1a237e")),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#b0b6d6")),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 10),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(KeepTogether([info_table, Spacer(1, 12)]))
        # Disk Info
        elements.append(Paragraph("SSD & HDD Info (Physical Drives)", section_heading))
        disk_data = [[Paragraph("<b>Model</b>", styles['Normal']), Paragraph("<b>Size</b>", styles['Normal']), Paragraph("<b>Type</b>", styles['Normal'])]]
        for model, size, dtype in info["Disks"]:
            disk_data.append([Paragraph(model, styles['Normal']), Paragraph(size, styles['Normal']), Paragraph(dtype, styles['Normal'])])
        disk_table = Table(disk_data, colWidths=[220, 90, 100])
        disk_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#e3e6f3")),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#1a237e")),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#b0b6d6")),
            ('FONTSIZE', (0,0), (-1,-1), 10),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#f5f7fa")]),
        ]))
        elements.append(KeepTogether([disk_table, Spacer(1, 16)]))
        # Monitor Details
        elements.append(Paragraph("Monitor Details", monitor_heading))
        monitor_details = info["Monitor Details"].replace('\n', '<br/>')
        elements.append(Paragraph(monitor_details, styles['Normal']))
        elements.append(Spacer(1, 18))
        # Footer
        elements.append(Paragraph("© 2025 System Asset Info | IT Department", footer_style))
        doc.build(elements)
        messagebox.showinfo("Exported", f"PDF exported successfully:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Export Error", f"Failed to export PDF:\n{e}")

# --------- UI Layout ---------
root = tk.Tk()
root.title("Professional System Asset Info")
root.geometry("800x940")
root.configure(bg="#f5f7fa")

style = ttk.Style()
style.configure('TLabel', font=('Segoe UI', 11), background="#f5f7fa")
style.configure('TButton', font=('Segoe UI', 11, 'bold'), padding=6)
style.configure('Treeview', font=('Consolas', 10), rowheight=24)
style.configure('Treeview.Heading', font=('Segoe UI', 11, 'bold'))

main_frame = ttk.Frame(root, padding=18, style='TFrame')
main_frame.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(
    main_frame,
    text="System Asset Information",
    font=('Segoe UI', 18, 'bold'),
    background="#f5f7fa",
    foreground="#1a237e"
)
header.pack(pady=(0,15))

# Info fields (all copyable now)
fields = ["System Name", "IP Address", "RAM", "CPU Model Name", "CPU Details", "OS Name", "OS Status"]
field_labels = {}
for f in fields:
    frame = ttk.Frame(main_frame)
    label = ttk.Label(frame, text=f + ":", width=18, font=('Segoe UI', 11, 'bold'))
    label.pack(side=tk.LEFT)
    value = ttk.Entry(frame, width=70, font=('Segoe UI', 11), foreground="#222")
    value.pack(side=tk.LEFT, fill=tk.X, expand=True)
    value.config(state='readonly')
    field_labels[f] = {"frame": frame, "label": label, "value": value}
    frame.pack(anchor='w', fill=tk.X, pady=2)

# Serial Number (as a highlighted, copyable field)
serial_frame = ttk.Frame(main_frame)
serial_label = ttk.Label(serial_frame, text="Serial Number:", width=18, font=('Segoe UI', 11, 'bold'))
serial_label.pack(side=tk.LEFT)
serial_entry = ttk.Entry(serial_frame, width=60, font=('Segoe UI', 11))
serial_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
serial_entry.config(state='readonly')
serial_frame.pack(anchor='w', fill=tk.X, pady=10)

# Product Key (as a highlighted, copyable field)
product_key_frame = ttk.Frame(main_frame)
product_key_label = ttk.Label(product_key_frame, text="Product Key:", width=18, font=('Segoe UI', 11, 'bold'))
product_key_label.pack(side=tk.LEFT)
product_key_entry = ttk.Entry(product_key_frame, width=60, font=('Segoe UI', 11))
product_key_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
product_key_entry.config(state='readonly')
product_key_frame.pack(anchor='w', fill=tk.X, pady=10)

# Disk Info - now as copyable text
disk_label = ttk.Label(main_frame, text="SSD & HDD Info (Physical Drives):", font=('Segoe UI', 12, 'bold'))
disk_label.pack(anchor='w', pady=(15,3))
disk_text_frame = ttk.Frame(main_frame)
disk_text_frame.pack(fill=tk.X)
disk_text = tk.Text(disk_text_frame, width=95, height=5, font=('Consolas', 10), wrap=tk.WORD, bg="#f8fafc", relief=tk.FLAT)
disk_text.config(state='disabled')
disk_text.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,2))
btn_copy_disk = ttk.Button(
    disk_text_frame,
    text="Copy",
    command=lambda: (disk_text.config(state='normal'), root.clipboard_clear(), root.clipboard_append(disk_text.get(1.0, tk.END).strip()), disk_text.config(state='disabled'))
)
btn_copy_disk.pack(side=tk.LEFT, padx=(6,0), pady=(0,3))

# Monitor Info - copyable text
monitor_label = ttk.Label(main_frame, text="Monitor Details:", font=('Segoe UI', 12, 'bold'))
monitor_label.pack(anchor='w', pady=(15,3))
monitor_text_frame = ttk.Frame(main_frame)
monitor_text_frame.pack(fill=tk.X)
monitor_text = tk.Text(monitor_text_frame, width=95, height=5, font=('Consolas', 10), wrap=tk.WORD, bg="#f8fafc", relief=tk.FLAT)
monitor_text.config(state='disabled')
monitor_text.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,2))
btn_copy_monitor = ttk.Button(
    monitor_text_frame,
    text="Copy",
    command=lambda: (monitor_text.config(state='normal'), root.clipboard_clear(), root.clipboard_append(monitor_text.get(1.0, tk.END).strip()), monitor_text.config(state='disabled'))
)
btn_copy_monitor.pack(side=tk.LEFT, padx=(6,0), pady=(0,3))

# Buttons
btn_frame = ttk.Frame(main_frame)
btn_frame.pack(pady=22)
btn_info = ttk.Button(btn_frame, text="Get System Info", command=show_info)
btn_info.pack(side=tk.LEFT, padx=(0,10))
btn_pdf = ttk.Button(btn_frame, text="Export as PDF", command=export_to_pdf)
btn_pdf.pack(side=tk.LEFT)

# Footer
footer = ttk.Label(
    main_frame,
    text="© 2025 System Asset Info | IT Department",
    font=('Segoe UI', 9),
    background="#f5f7fa",
    foreground="#999"
)
footer.pack(side=tk.BOTTOM, pady=(18,0))

root.mainloop()