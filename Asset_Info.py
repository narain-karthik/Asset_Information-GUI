import tkinter as tk
from tkinter import scrolledtext, ttk, messagebox, filedialog
import platform
import socket
import psutil
import subprocess
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

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

def get_disks_physical():
    try:
        result = subprocess.check_output(
            'wmic diskdrive get Model,Size /format:csv',
            shell=True
        ).decode(errors="ignore").split('\n')
        disks = []
        for line in result:
            if line.strip() and not line.startswith('Node,'):
                parts = line.strip().split(',')
                if len(parts) == 3:
                    _, model, size = parts
                    if "virtual" in model.lower():
                        continue
                    if size.isdigit():
                        size_gb = f"{int(size) / (1024 ** 3):.2f} GB"
                    else:
                        size_gb = "Unknown Size"
                    disks.append((model, size_gb))
        return disks
    except Exception as e:
        return [("Error", str(e))]

def get_system_model():
    try:
        result = subprocess.check_output(
            'wmic computersystem get model', shell=True
        ).decode(errors="ignore").strip().split('\n')
        for line in result[1:]:
            if line.strip():
                return line.strip()
    except Exception:
        pass
    return "Unknown Model"

def get_cpu_details():
    try:
        proc_info = subprocess.check_output(
            'wmic cpu get Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed /format:csv',
            shell=True
        ).decode(errors="ignore").split('\n')
        for line in proc_info:
            if line.strip() and not line.startswith('Node,'):
                parts = line.strip().split(',')
                if len(parts) == 5:
                    _, name, cores, logical, freq = parts
                    return f"{name}, {freq} MHz, {cores} Core(s), {logical} Logical Processor(s)"
        return "Unknown CPU"
    except Exception as e:
        return f"Error: {e}"

def get_os_name():
    return f"{platform.system()} {platform.release()}"

def get_status():
    return "Active"

def get_cpu_tag():
    try:
        result = subprocess.check_output("wmic bios get serialnumber", shell=True)
        lines = result.decode(errors="ignore").strip().split('\n')
        for line in lines[1:]:
            if line.strip():
                return line.strip()
    except Exception as e:
        return f"Error: {e}"
    return "Not Found"

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
            ["powershell", "-Command", powershell_script],
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
    info["CPU Tag Number"] = get_cpu_tag()
    info["Monitor Details"] = get_monitor_tags()
    info["Disks"] = get_disks_physical()
    info["OS Name"] = get_os_name()
    info["OS Status"] = get_status()
    return info

def show_info():
    try:
        info = gather_info()
        for key in field_labels:
            field_labels[key]["value"].config(text=info[key] if key not in ["Disks", "Monitor Details", "CPU Tag Number"] else "")
        disk_tree.delete(*disk_tree.get_children())
        for model, size in info["Disks"]:
            disk_tree.insert('', 'end', values=(model, size))
        monitor_text.config(state='normal')
        monitor_text.delete(1.0, tk.END)
        monitor_text.insert(tk.END, info["Monitor Details"])
        monitor_text.config(state='disabled')
        cpu_tag_entry.config(state='normal')
        cpu_tag_entry.delete(0, tk.END)
        cpu_tag_entry.insert(0, info["CPU Tag Number"])
        cpu_tag_entry.config(state='readonly')
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

def export_to_pdf():
    info = gather_info()
    default_filename = f"{info['System Name']}-Info.pdf"
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

        # Info Table (no HTML tags, just bold text for keys)
        table_data = []
        field_names = [
            ("System Name", info["System Name"]),
            ("IP Address", info["IP Address"]),
            ("RAM", info["RAM"]),
            ("CPU Model Name", info["CPU Model Name"]),
            ("CPU Details", info["CPU Details"]),
            ("CPU Tag Number", info["CPU Tag Number"]),
            ("OS Name", info["OS Name"]),
            ("OS Status", info["OS Status"]),
        ]
        for key, value in field_names:
            table_data.append(
                [Paragraph(f"<b>{key}</b>", styles['Normal']), Paragraph(str(value), styles['Normal'])]
            )
        info_table = Table(table_data, colWidths=[155, 325])
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
        disk_data = [[Paragraph("<b>Model</b>", styles['Normal']), Paragraph("<b>Size</b>", styles['Normal'])]]
        for model, size in info["Disks"]:
            disk_data.append([Paragraph(model, styles['Normal']), Paragraph(size, styles['Normal'])])
        disk_table = Table(disk_data, colWidths=[300, 90])
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
root.geometry("760x760")
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

# Info fields
fields = [
    "System Name", "IP Address", "RAM",
    "CPU Model Name", "CPU Details",
    "OS Name", "OS Status"
]
field_labels = {}
for f in fields:
    frame = ttk.Frame(main_frame)
    frame.pack(anchor='w', fill=tk.X, pady=2)
    label = ttk.Label(frame, text=f + ":", width=18, font=('Segoe UI', 11, 'bold'))
    label.pack(side=tk.LEFT)
    value = ttk.Label(frame, text="", width=70, font=('Segoe UI', 11), foreground="#222")
    value.pack(side=tk.LEFT, fill=tk.X, expand=True)
    field_labels[f] = {"label": label, "value": value}

# CPU Tag Number (as a highlighted, copyable field)
cpu_tag_frame = ttk.Frame(main_frame)
cpu_tag_frame.pack(anchor='w', fill=tk.X, pady=10)
cpu_tag_label = ttk.Label(cpu_tag_frame, text="CPU Tag Number:", width=18, font=('Segoe UI', 11, 'bold'))
cpu_tag_label.pack(side=tk.LEFT)
cpu_tag_entry = ttk.Entry(cpu_tag_frame, width=60, font=('Segoe UI', 11))
cpu_tag_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
cpu_tag_entry.config(state='readonly')

# Disk Info
disk_label = ttk.Label(main_frame, text="SSD & HDD Info (Physical Drives):", font=('Segoe UI', 12, 'bold'))
disk_label.pack(anchor='w', pady=(15,3))
disk_tree = ttk.Treeview(main_frame, columns=("Model", "Size"), show='headings', height=4)
disk_tree.heading("Model", text="Model")
disk_tree.heading("Size", text="Size")
disk_tree.column("Model", width=400, anchor='w')
disk_tree.column("Size", width=100, anchor='center')
disk_tree.pack(fill=tk.X, padx=4)

# Monitor Info
monitor_label = ttk.Label(main_frame, text="Monitor Details:", font=('Segoe UI', 12, 'bold'))
monitor_label.pack(anchor='w', pady=(15,3))
monitor_text = scrolledtext.ScrolledText(main_frame, width=80, height=4, font=('Consolas', 10), wrap=tk.WORD)
monitor_text.pack(fill=tk.X)
monitor_text.config(state='disabled', background="#f8fafc")

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