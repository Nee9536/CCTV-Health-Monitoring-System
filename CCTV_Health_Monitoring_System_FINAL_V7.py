import os
import re
import sys
import time
import queue
import ipaddress
import subprocess
import threading
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit(
        "Missing dependency: openpyxl\n"
        "Install with: py -m pip install openpyxl"
    )


# ============================================================
# CCTV HEALTH MONITORING SYSTEM
# Daily Camera & NVR Health Monitoring
# ============================================================

APP_TITLE = "CCTV HEALTH MONITORING SYSTEM"
SUBTITLE = "Daily Camera & NVR Health Monitoring"
AUTHOR = "Mr. Neeraj Kumar"
ROLE = "Systems Administration & IT Operations"

MAX_REPORTS_PER_DAY = 5
PING_TIMEOUT_MS = 1500
PING_ATTEMPTS = 2
SCAN_WORKERS = 16

APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)

DEFAULT_MASTER_NAMES = ("IPS_SHEET.xlsx", "MASTER SHEET.xlsx", "MASTER_SHEET.xlsx")
REPORT_ROOT_NAME = "CCTV_Reports"


# ----------------------------- Utility -----------------------------

def clean(value):
    return "" if value is None else str(value).strip()


def excel_safe_text(value):
    """Return text safe for XLSX XML (remove illegal control characters)."""
    if value is None:
        return ""
    text = str(value)
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)


def normalize_header(value):
    s = clean(value).replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()


def valid_ip(value):
    try:
        ipaddress.ip_address(clean(value))
        return True
    except Exception:
        return False


def normalize_mac(mac):
    mac = clean(mac).replace("-", ":").replace(".", "").upper()
    if re.fullmatch(r"(?:[0-9A-F]{2}:){5}[0-9A-F]{2}", mac):
        return mac
    if re.fullmatch(r"[0-9A-F]{12}", mac):
        return ":".join(mac[i:i + 2] for i in range(0, 12, 2))
    return ""


def is_nvr(device_name):
    return clean(device_name).upper().startswith("NVR")


def windows_ping(ip):
    """Return (online, response_ms, remarks).

    Uses two ICMP attempts to avoid a false OFFLINE result caused by a
    single transient packet loss. A device is WORKING only when Windows
    receives at least one valid ping reply.
    """
    if not valid_ip(ip):
        return False, None, "Invalid IP Address"

    cmd = ["ping", "-n", str(PING_ATTEMPTS), "-w", str(PING_TIMEOUT_MS), ip]
    start = time.perf_counter()

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="cp850",
            errors="ignore",
            timeout=5.5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

        elapsed = round((time.perf_counter() - start) * 1000, 1)
        output = (result.stdout or "") + "\n" + (result.stderr or "")

        # Windows ping returns code 0 when at least one echo reply is received.
        if result.returncode == 0:
            # Prefer the first/lowest reported successful reply time.
            matches = re.findall(r"time\s*[=<]\s*(\d+)\s*ms", output, flags=re.IGNORECASE)
            if matches:
                response_ms = min(float(x) for x in matches)
            else:
                response_ms = elapsed
            return True, response_ms, "Ping Reply Received"

        return False, None, "Request Timed Out"

    except subprocess.TimeoutExpired:
        return False, None, "Ping Timeout"
    except Exception as exc:
        return False, None, f"Ping Error: {exc}"


def fetch_mac(ip):
    """
    Resolve MAC for a local/reachable IP.
    ARP is used first; Get-NetNeighbor is the Windows fallback.

    Important networking rule:
    A real remote Internet MAC cannot normally be obtained from an IP.
    For a camera/NVR on the same LAN/VLAN as the monitoring PC,
    Windows ARP/neighbor information can provide the device MAC.
    """
    if os.name != "nt" or not valid_ip(ip):
        return ""

    # 1) Windows ARP table
    try:
        result = subprocess.run(
            ["arp", "-a", ip],
            capture_output=True,
            text=True,
            encoding="cp850",
            errors="ignore",
            timeout=2.5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        combined = result.stdout + "\n" + result.stderr

        # Example:
        # 192.168.3.4    aa-bb-cc-dd-ee-ff    dynamic
        pattern = (
            rf"(?mi)^\s*{re.escape(ip)}\s+"
            r"([0-9a-f]{2}(?:-[0-9a-f]{2}){5})\s+"
        )
        match = re.search(pattern, combined)
        if match:
            mac = normalize_mac(match.group(1))
            if mac:
                return mac
    except Exception:
        pass

    # 2) PowerShell Get-NetNeighbor fallback
    ps = (
        "Get-NetNeighbor -IPAddress '"
        + ip.replace("'", "''")
        + "' -ErrorAction SilentlyContinue "
        "| Select-Object -First 1 -ExpandProperty LinkLayerAddress"
    )

    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                ps,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=3.0,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

        for line in result.stdout.splitlines():
            mac = normalize_mac(line)
            if mac:
                return mac
    except Exception:
        pass

    return ""


def scan_device(device):
    # IMPORTANT: every row is evaluated independently by its own IP.
    # NVR state never changes a camera's state and camera state never changes NVR state.
    online, response_ms, ping_remarks = windows_ping(device["ip"])

    if online:
        mac = fetch_mac(device["ip"])
        status = "WORKING"

        if mac:
            remarks = ping_remarks
        else:
            remarks = "Online - MAC Not Available"

        return {
            **device,
            "mac": mac,
            "status": status,
            "response_ms": response_ms,
            "remarks": remarks,
        }

    return {
        **device,
        "mac": "",
        "status": "OFFLINE",
        "response_ms": None,
        "remarks": ping_remarks,
    }



# ----------------------------- Excel Master Reader -----------------------------

def find_master_columns(headers):
    normalized = {normalize_header(v): i for i, v in enumerate(headers)}

    def find(*names):
        for name in names:
            n = normalize_header(name)
            if n in normalized:
                return normalized[n]

        # Flexible fallback
        for key, idx in normalized.items():
            if any(n in key for n in [normalize_header(x) for x in names]):
                return idx
        return None

    return {
        "sr": find("Sr. No", "Sr No"),
        "digital_nvr": find("Digital NVR"),
        "device_name": find(
            "Device Name / Camera & NVR",
            "Device Name",
            "Camera & NVR",
        ),
        "ip": find("IP Address", "IP"),
        "location": find("Location"),
        "offline_date": find("Device Offline Date"),
        "offline_time": find("Device Offline Time"),
        "online_date": find("Device Online Date"),
        "online_time": find("Device Online Time"),
        "status": find("Status"),
        "report_date": find("Report Generate Date"),
        "report_time": find("Report Generate Time"),
    }


def load_master_sheet(path):
    wb = load_workbook(path, read_only=True, data_only=True)

    if "MASTER SHEET" in wb.sheetnames:
        ws = wb["MASTER SHEET"]
    else:
        ws = wb[wb.sheetnames[0]]

    values = list(ws.iter_rows(values_only=True))

    if not values:
        raise ValueError("The selected Excel file is empty.")

    headers = list(values[0])
    cols = find_master_columns(headers)

    required = {
        "Digital NVR": cols["digital_nvr"],
        "Device Name / Camera & NVR": cols["device_name"],
        "IP Address": cols["ip"],
        "Location": cols["location"],
    }

    missing = [name for name, index in required.items() if index is None]

    if missing:
        raise ValueError(
            "Required column(s) not found:\n\n"
            + "\n".join(f"• {x}" for x in missing)
        )

    devices = []

    # IMPORTANT:
    # We intentionally preserve the exact Excel row sequence.
    # No grouping, sorting, or re-ordering is performed.
    for excel_row, row in enumerate(values[1:], start=2):
        def get_value(key):
            idx = cols.get(key)
            if idx is None or idx >= len(row):
                return ""
            return clean(row[idx])

        device_name = get_value("device_name")
        ip = get_value("ip")

        if not device_name and not ip:
            continue

        devices.append(
            {
                "source_index": len(devices),
                "excel_row": excel_row,
                "sr_no": get_value("sr"),
                "digital_nvr": get_value("digital_nvr"),
                "device_name": device_name,
                "ip": ip,
                "location": get_value("location"),
                "source_status": get_value("status"),
            }
        )

    wb.close()

    if not devices:
        raise ValueError("No device records were found in the Master Sheet.")

    return devices


# ----------------------------- Report Generation -----------------------------

def count_existing_reports(report_day_folder):
    """Count only CCTV Health reports in the current date folder."""
    if not report_day_folder.exists():
        return 0
    pattern = re.compile(r"^CCTV_Health_Report_\d{2}-\d{2}-\d{4}_Run-\d{2}_\d{6}\.xlsx$", re.I)
    return sum(1 for p in report_day_folder.iterdir() if p.is_file() and pattern.match(p.name))


def create_report(results, master_path):
    now = datetime.now()

    # Exactly:
    # CCTV_Reports / September / 2026-09-24 / Run-01.xlsx
    report_root = master_path.parent / REPORT_ROOT_NAME
    month_folder = report_root / now.strftime("%B")
    date_folder = month_folder / now.strftime("%Y-%m-%d")
    date_folder.mkdir(parents=True, exist_ok=True)

    existing = count_existing_reports(date_folder)

    if existing >= MAX_REPORTS_PER_DAY:
        return None, (
            f"Daily report limit reached.\n\n"
            f"Maximum {MAX_REPORTS_PER_DAY} reports are allowed per day."
        )

    run_no = existing + 1
    filename = (
        f"CCTV_Health_Report_{now.strftime('%d-%m-%Y')}_"
        f"Run-{run_no:02d}_{now.strftime('%H%M%S')}.xlsx"
    )
    output_path = date_folder / filename

    wb = Workbook()
    ws = wb.active
    ws.title = "Health Report"

    # ---------------- Theme ----------------
    navy = "08233D"
    blue = "155A96"
    cyan = "08A9D6"
    green = "1DB954"
    green_light = "DDF5E7"
    red = "D9364A"
    red_light = "FCE1E5"
    gold = "F4B400"
    gold_light = "FFF2CC"
    purple = "7657E8"
    light_blue = "DCEEFF"
    white = "FFFFFF"
    dark = "17202A"
    border_color = "B9C7D6"
    grey = "EAF0F6"

    thin = Side(style="thin", color=border_color)
    medium = Side(style="medium", color="6B7C93")

    # ---------------- Title ----------------
    ws.merge_cells("A1:N1")
    ws["A1"] = APP_TITLE
    ws["A1"].font = Font(
        name="Segoe UI",
        size=20,
        bold=True,
        color=white,
    )
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    ws.merge_cells("A2:N2")
    ws["A2"] = SUBTITLE
    ws["A2"].font = Font(
        name="Segoe UI",
        size=11,
        italic=True,
        color=white,
    )
    ws["A2"].fill = PatternFill("solid", fgColor=blue)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24

    # ---------------- Summary cards ----------------
    total = len(results)
    working = sum(1 for r in results if r["status"] == "WORKING")
    offline = total - working
    nvr_count = sum(1 for r in results if is_nvr(r["device_name"]))
    camera_count = total - nvr_count

    summary = [
        ("A4:C5", "TOTAL DEVICES", total, blue),
        ("D4:F5", "WORKING", working, green),
        ("G4:I5", "OFFLINE", offline, red),
        ("J4:L5", "TOTAL NVR", nvr_count, purple),
        ("M4:N5", "TOTAL CAMERA", camera_count, cyan),
    ]

    for merge_range, label, value, color in summary:
        ws.merge_cells(merge_range)
        start = ws[merge_range.split(":")[0]]
        start.value = f"{label}\n{value}"
        start.font = Font(
            name="Segoe UI",
            size=13,
            bold=True,
            color=white,
        )
        start.fill = PatternFill("solid", fgColor=color)
        start.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        start.border = Border(
            left=thin, right=thin, top=thin, bottom=thin
        )

    ws.row_dimensions[4].height = 25
    ws.row_dimensions[5].height = 25

    # ---------------- Report information ----------------
    info = [
        ("A7", "Master Sheet", "B7", str(master_path)),
        ("A8", "Report Generate Date", "B8", now.strftime("%d-%m-%Y")),
        ("A9", "Report Generate Time", "B9", now.strftime("%I:%M:%S %p")),
        ("A10", "Run Number", "B10", f"Run-{run_no:02d}"),
        ("A11", "Generated By", "B11", AUTHOR),
    ]

    for label_cell, label, value_cell, value in info:
        ws[label_cell] = label
        ws[label_cell].font = Font(
            name="Segoe UI",
            bold=True,
            color=white,
        )
        ws[label_cell].fill = PatternFill("solid", fgColor=navy)
        ws[label_cell].alignment = Alignment(vertical="center")

        ws.merge_cells(f"{value_cell}:N{value_cell[1:]}")
        ws[value_cell] = value
        ws[value_cell].font = Font(
            name="Segoe UI",
            color=dark,
        )
        ws[value_cell].fill = PatternFill("solid", fgColor="F7FAFC")
        ws[value_cell].alignment = Alignment(vertical="center")

    # ---------------- Detail table ----------------
    table_start = 13

    headers = [
        "Sr. No",
        "Digital NVR",
        "Device Name / Camera & NVR",
        "IP Address",
        "Location",
        "MAC Address",
        "Status",
        "Response Time (ms)",
        "Remarks",
        "Device Offline Date",
        "Device Offline Time",
        "Device Online Date",
        "Device Online Time",
        "Report Generate Date",
        "Report Generate Time",
        "Scan Timestamp",
    ]

    # Use A:P (16 columns), not an artificial Device Type column.
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(table_start, col, header)
        cell.font = Font(
            name="Segoe UI",
            bold=True,
            color=white,
        )
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        cell.border = Border(
            left=thin, right=thin, top=thin, bottom=thin
        )

    ws.row_dimensions[table_start].height = 36

    report_date = now.strftime("%d-%m-%Y")
    report_time = now.strftime("%I:%M:%S %p")
    scan_timestamp = now.strftime("%d-%m-%Y %I:%M:%S %p")

    for index, result in enumerate(results, start=table_start + 1):
        status = result["status"]

        if status == "WORKING":
            status_fill = green_light
            status_font = "137333"
        else:
            status_fill = red_light
            status_font = "B42318"

        # Keep exact source sequence.
        values = [
            result["sr_no"] or index - table_start,
            excel_safe_text(result["digital_nvr"]),
            excel_safe_text(result["device_name"]),
            excel_safe_text(result["ip"]),
            excel_safe_text(result["location"]),
            excel_safe_text(result["mac"] or "-"),
            status,
            (
                f"{result['response_ms']:.1f}"
                if isinstance(result["response_ms"], (int, float))
                else "-"
            ),
            excel_safe_text(result["remarks"]),
            report_date if status == "OFFLINE" else "",
            report_time if status == "OFFLINE" else "",
            report_date if status == "WORKING" else "",
            report_time if status == "WORKING" else "",
            report_date,
            report_time,
            scan_timestamp,
        ]

        for col, value in enumerate(values, start=1):
            cell = ws.cell(index, col, value)
            cell.font = Font(
                name="Segoe UI",
                size=10,
                color=dark,
            )
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center" if col in (1, 6, 7, 8) else "left",
                wrap_text=True,
            )
            cell.border = Border(
                left=thin, right=thin, top=thin, bottom=thin
            )

        # Status cell
        status_cell = ws.cell(index, 7)
        status_cell.font = Font(
            name="Segoe UI",
            bold=True,
            color=status_font,
        )
        status_cell.fill = PatternFill("solid", fgColor=status_fill)
        status_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

        # MAC cell
        ws.cell(index, 6).font = Font(
            name="Consolas",
            size=10,
            bold=True,
            color=dark,
        )

        ws.row_dimensions[index].height = 28

    last_row = table_start + len(results)

    # Excel-safe filtering: use the worksheet AutoFilter instead of an
    # OOXML Table object.  This avoids Excel repair prompts on stricter
    # Excel builds while retaining filters and the same visual formatting.
    table_ref = f"A{table_start}:P{last_row}"
    ws.auto_filter.ref = table_ref

    # Manual alternating row banding keeps the report colourful without
    # creating a table XML part that can trigger workbook-repair dialogs.
    for row_no in range(table_start + 1, last_row + 1):
        if (row_no - table_start) % 2 == 0:
            for col_no in range(1, 17):
                cell = ws.cell(row_no, col_no)
                if col_no != 7:  # preserve status colour
                    cell.fill = PatternFill("solid", fgColor="F5F9FD")

    # ---------------- Widths ----------------
    widths = {
        "A": 9,
        "B": 14,
        "C": 28,
        "D": 17,
        "E": 22,
        "F": 22,
        "G": 14,
        "H": 18,
        "I": 28,
        "J": 20,
        "K": 20,
        "L": 20,
        "M": 20,
        "N": 22,
        "O": 22,
        "P": 27,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Freeze the actual detail header.
    ws.freeze_panes = "A14"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:13"

    # Footer
    footer_row = last_row + 2
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=16)
    ws.cell(footer_row, 1).value = (
        "Generated automatically by CCTV HEALTH MONITORING SYSTEM"
    )
    ws.cell(footer_row, 1).font = Font(
        name="Segoe UI",
        italic=True,
        color="64748B",
    )
    ws.cell(footer_row, 1).alignment = Alignment(horizontal="center")

    wb.save(output_path)
    wb.close()

    return output_path, ""


# ----------------------------- Professional Button -----------------------------

class RoundedButton(tk.Canvas):
    """Stable professional button: fixed rounded-rectangle, no hover.

    Uses Windows symbol fonts for lightweight icons. Text/icon positions are
    calculated from real font metrics so the icon can never overlap the label.
    """

    ICONS = {
        "play": "▶",
        "folder": "▰",
        "sheet": "▤",
        "refresh": "↻",
        "document": "▤",
        "check": "✓",
    }

    def __init__(self, master, text, command, width=190, height=42,
                 bg="#12395E", hover_bg=None, fg="white",
                 font=("Segoe UI", 10, "bold"), radius=6,
                 icon="", disabled_bg="#334155", **kwargs):
        super().__init__(master, width=width, height=height,
                         bg=master.cget("bg"), highlightthickness=0,
                         bd=0, **kwargs)
        self.button_width = width
        self.button_height = height
        self.bg_color = bg
        self.fg_color = fg
        self.disabled_bg = disabled_bg
        self.command = command
        self.radius = min(radius, 7)
        self.text = text
        self.icon = icon
        self.font = font
        self.enabled = True

        # Intentionally no Enter/Leave bindings: hover is disabled.
        self.bind("<Button-1>", self.on_click)
        self.draw(self.bg_color)

    def rounded_rect(self, x1, y1, x2, y2, r, fill, outline=None):
        # Clean rectangular button with restrained corner radius.
        self.create_arc(x1, y1, x1 + 2*r, y1 + 2*r,
                        start=90, extent=90, fill=fill, outline=outline or fill)
        self.create_arc(x2 - 2*r, y1, x2, y1 + 2*r,
                        start=0, extent=90, fill=fill, outline=outline or fill)
        self.create_arc(x1, y2 - 2*r, x1 + 2*r, y2,
                        start=180, extent=90, fill=fill, outline=outline or fill)
        self.create_arc(x2 - 2*r, y2 - 2*r, x2, y2,
                        start=270, extent=90, fill=fill, outline=outline or fill)
        self.create_rectangle(x1+r, y1, x2-r, y2,
                              fill=fill, outline=outline or fill)
        self.create_rectangle(x1, y1+r, x2, y2-r,
                              fill=fill, outline=outline or fill)

    def draw(self, color):
        self.delete("all")
        self.rounded_rect(1, 1, self.button_width-1, self.button_height-1,
                          self.radius, color)

        # Real font metrics prevent icon/text collisions.
        try:
            text_font = tkfont.Font(font=self.font)
        except Exception:
            text_font = tkfont.Font(family="Segoe UI", size=10, weight="bold")

        icon_text = self.ICONS.get(self.icon, "")
        icon_font = tkfont.Font(family="Segoe UI Symbol", size=13, weight="normal")
        text_width = text_font.measure(self.text)
        icon_width = icon_font.measure(icon_text) if icon_text else 0
        gap = 7 if icon_text else 0
        total_width = text_width + icon_width + gap

        # If a label is longer than the requested button, reduce font slightly
        # rather than allowing overlap.
        if total_width > self.button_width - 20:
            try:
                size = int(self.font[1])
                while total_width > self.button_width - 20 and size > 8:
                    size -= 1
                    text_font.configure(size=size)
                    text_width = text_font.measure(self.text)
                    total_width = text_width + icon_width + gap
            except Exception:
                pass

        start_x = max(10, (self.button_width - total_width) / 2)
        center_y = self.button_height / 2

        if icon_text:
            self.create_text(start_x + icon_width/2, center_y,
                             text=icon_text, fill=self.fg_color,
                             font=icon_font)
            text_x = start_x + icon_width + gap + text_width/2
        else:
            text_x = self.button_width/2

        self.create_text(text_x, center_y, text=self.text,
                         fill=self.fg_color, font=text_font)

    def on_click(self, _event=None):
        if self.enabled and self.command:
            self.command()

    def set_enabled(self, enabled):
        self.enabled = enabled
        self.draw(self.bg_color if enabled else self.disabled_bg)


class CCTVHealthApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1440x900")
        self.root.minsize(1180, 760)
        self.root.configure(bg="#061B2E")

        try:
            self.root.state("zoomed")
        except Exception:
            pass

        self.master_path = None
        self.devices = []
        self.results = []
        self.scanning = False
        self.ui_queue = queue.Queue()
        self.scan_started = None
        self.preview_items = {}
        self.latest_report_path = None

        self.colors = {
            "bg": "#061B2E",
            "panel": "#0B2945",
            "panel2": "#0E3355",
            "header": "#08315A",
            "blue": "#2388D9",
            "cyan": "#08B8E8",
            "green": "#19C37D",
            "green_dark": "#0B8F5A",
            "red": "#EF476F",
            "red_dark": "#B4234D",
            "orange": "#F59E0B",
            "purple": "#7C5CFC",
            "text": "#F7FAFC",
            "muted": "#AFC5D8",
            "border": "#1D537B",
            "white": "#FFFFFF",
        }

        self.setup_styles()
        self.build_splash()

    # ---------------- Splash ----------------

    def build_splash(self):
        self.splash = tk.Frame(self.root, bg=self.colors["bg"])
        self.splash.pack(fill="both", expand=True)

        # No circle: clean technical camera icon.
        icon_canvas = tk.Canvas(
            self.splash,
            width=170,
            height=130,
            bg=self.colors["bg"],
            highlightthickness=0,
        )
        icon_canvas.place(relx=0.5, rely=0.17, anchor="center")

        # Camera body
        icon_canvas.create_rounded_rect = None  # marker only

        icon_canvas.create_rectangle(
            45, 42, 125, 88,
            fill="#174E79",
            outline="#2BC7F5",
            width=3,
        )
        icon_canvas.create_rectangle(
            58, 50, 90, 80,
            fill="#0C2C48",
            outline="#8A63FF",
            width=2,
        )
        icon_canvas.create_oval(
            67, 57, 82, 72,
            fill="#1FE2A6",
            outline="#1FE2A6",
        )
        icon_canvas.create_polygon(
            125, 53, 151, 44, 151, 86, 125, 77,
            fill="#17679C",
            outline="#2BC7F5",
            width=2,
        )
        icon_canvas.create_line(
            53, 88, 40, 105,
            fill="#2BC7F5",
            width=4,
        )
        icon_canvas.create_line(
            40, 105, 68, 105,
            fill="#2BC7F5",
            width=4,
        )

        title = tk.Label(
            self.splash,
            text="CCTV HEALTH MONITORING SYSTEM",
            bg=self.colors["bg"],
            fg=self.colors["white"],
            font=("Segoe UI", 27, "bold"),
        )
        title.place(relx=0.5, rely=0.34, anchor="center")

        subtitle = tk.Label(
            self.splash,
            text=SUBTITLE,
            bg=self.colors["bg"],
            fg="#8CCEF4",
            font=("Segoe UI", 13),
        )
        subtitle.place(relx=0.5, rely=0.39, anchor="center")

        self.splash_status = tk.Label(
            self.splash,
            text="Initializing network monitoring engine",
            bg=self.colors["bg"],
            fg=self.colors["white"],
            font=("Segoe UI", 11),
        )
        self.splash_status.place(relx=0.5, rely=0.52, anchor="center")

        self.splash_percent = tk.Label(
            self.splash,
            text="0%",
            bg=self.colors["bg"],
            fg="#19E6A2",
            font=("Segoe UI", 30, "bold"),
        )
        self.splash_percent.place(relx=0.5, rely=0.59, anchor="center")

        progress_bg = tk.Frame(
            self.splash,
            bg="#123A5C",
            highlightbackground="#2A668F",
            highlightthickness=1,
            height=22,
        )
        progress_bg.place(relx=0.5, rely=0.67, relwidth=0.55, anchor="center")

        self.splash_progress = tk.Frame(
            progress_bg,
            bg="#19E6A2",
            height=20,
        )
        self.splash_progress.place(x=1, y=1, relwidth=0, relheight=1)

        tk.Label(
            self.splash,
            text=AUTHOR,
            bg=self.colors["bg"],
            fg=self.colors["white"],
            font=("Segoe UI", 14, "bold"),
        ).place(relx=0.5, rely=0.77, anchor="center")

        tk.Label(
            self.splash,
            text=ROLE,
            bg=self.colors["bg"],
            fg="#A9BBD0",
            font=("Segoe UI", 11, "italic"),
        ).place(relx=0.5, rely=0.815, anchor="center")

        self.splash_bottom = tk.Label(
            self.splash,
            text="Preparing network scanner...",
            bg=self.colors["bg"],
            fg="#2BC7F5",
            font=("Segoe UI", 9),
        )
        self.splash_bottom.place(relx=0.5, rely=0.88, anchor="center")

        self.splash_start = time.perf_counter()
        self.splash_step = 0
        self.root.after(100, self.animate_splash)

    def animate_splash(self):
        self.splash_step += 1
        percent = min(self.splash_step, 100)

        self.splash_percent.configure(text=f"{percent}%")

        # Professional staged messages.
        if percent < 15:
            message = "Loading application components..."
        elif percent < 30:
            message = "Preparing network scanner..."
        elif percent < 50:
            message = "Preparing IP reachability engine..."
        elif percent < 70:
            message = "Preparing MAC address resolver..."
        elif percent < 85:
            message = "Preparing CCTV report engine..."
        elif percent < 96:
            message = "Preparing monitoring dashboard..."
        else:
            message = "CCTV monitoring engine ready."

        self.splash_status.configure(text=message)
        self.splash_bottom.configure(
            text=f"System initialization • {percent}%"
        )

        self.splash_progress.place_configure(
            relwidth=max(0.001, percent / 100)
        )

        if percent < 100:
            self.root.after(100, self.animate_splash)
        else:
            # 100 steps x 100ms = minimum 10 seconds.
            self.root.after(250, self.show_home)

    # ---------------- UI setup ----------------

    def setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Treeview",
            background="#F8FAFC",
            foreground="#102A43",
            fieldbackground="#F8FAFC",
            rowheight=31,
            font=("Segoe UI", 9),
            borderwidth=0,
        )

        style.configure(
            "Treeview.Heading",
            background="#164D78",
            foreground="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padding=(8, 8),
        )

        style.map(
            "Treeview",
            background=[("selected", "#BEE7FF")],
            foreground=[("selected", "#09243A")],
        )

        style.configure(
            "CCTV.Horizontal.TProgressbar",
            troughcolor="#123A5C",
            background="#19C37D",
            bordercolor="#2A668F",
            lightcolor="#19C37D",
            darkcolor="#19C37D",
            thickness=8,
        )

    def clear_root(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def show_home(self):
        self.clear_root()

        self.home = tk.Frame(
            self.root,
            bg=self.colors["bg"],
        )
        self.home.pack(fill="both", expand=True)

        self.build_header()
        self.build_master_bar()
        self.build_summary_cards()
        self.build_action_bar()
        self.build_scan_status()
        self.build_live_preview()

        self.root.after(100, self.process_queue)

        # Auto-detect a master sheet beside the application.
        for name in DEFAULT_MASTER_NAMES:
            candidate = APP_DIR / name
            if candidate.exists():
                try:
                    self.load_master(candidate, silent=True)
                except Exception:
                    pass
                break

    # ---------------- Header ----------------

    def build_header(self):
        header = tk.Frame(self.home, bg=self.colors["header"], height=78)
        header.pack(fill="x", padx=0, pady=0)
        header.pack_propagate(False)

        # Responsive 3-column header prevents overlap on different resolutions.
        left = tk.Frame(header, bg=self.colors["header"])
        left.pack(side="left", fill="y", padx=(18, 10))

        icon = tk.Canvas(left, width=58, height=54, bg=self.colors["header"], highlightthickness=0)
        icon.pack(side="left", pady=11)
        icon.create_rectangle(8, 18, 38, 38, fill="#174E79", outline="#2BC7F5", width=2)
        icon.create_oval(17, 22, 29, 34, fill="#1FE2A6", outline="#1FE2A6")
        icon.create_polygon(38, 20, 52, 15, 52, 41, 38, 36, fill="#17679C", outline="#2BC7F5")
        icon.create_line(17, 38, 12, 48, fill="#2BC7F5", width=3)
        icon.create_line(12, 48, 27, 48, fill="#2BC7F5", width=3)

        title_box = tk.Frame(left, bg=self.colors["header"])
        title_box.pack(side="left", pady=9)
        tk.Label(title_box, text=APP_TITLE, bg=self.colors["header"], fg="white",
                 font=("Segoe UI", 19, "bold"), anchor="w").pack(anchor="w")
        tk.Label(title_box, text=SUBTITLE, bg=self.colors["header"], fg="#9BD8F4",
                 font=("Segoe UI", 9), anchor="w").pack(anchor="w", pady=(1, 0))

        right = tk.Frame(header, bg=self.colors["header"])
        right.pack(side="right", fill="y", padx=(10, 18))

        tk.Label(right, text="Developed By", bg=self.colors["header"], fg="#91AFC6",
                 font=("Segoe UI", 8)).grid(row=0, column=0, sticky="e", padx=(0, 7), pady=(13, 0))
        tk.Label(right, text=AUTHOR, bg=self.colors["header"], fg="white",
                 font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="e", pady=(13, 0))
        tk.Label(right, text=ROLE, bg=self.colors["header"], fg="#A9C5D9",
                 font=("Segoe UI", 8)).grid(row=1, column=1, sticky="e", pady=(1, 0))
        self.clock_label = tk.Label(right, text="", bg=self.colors["header"], fg="#DCEEFF",
                                    font=("Segoe UI", 8, "bold"))
        self.clock_label.grid(row=2, column=1, sticky="e", pady=(2, 8))
        self.update_clock()

    def update_clock(self):
        if not self.root.winfo_exists():
            return
        self.clock_label.configure(text=datetime.now().strftime("%d %B %Y  •  %I:%M:%S %p"))
        self.root.after(1000, self.update_clock)

    # ---------------- Master bar ----------------

    def build_master_bar(self):
        frame = tk.Frame(self.home, bg=self.colors["panel"], highlightbackground=self.colors["border"], highlightthickness=1)
        frame.pack(fill="x", padx=14, pady=(10, 8))

        tk.Label(frame, text="MASTER SHEET", bg=self.colors["panel"], fg="white",
                 font=("Segoe UI", 10, "bold")).pack(side="left", padx=(14, 12), pady=10)
        self.master_path_var = tk.StringVar(value="No Master Sheet selected")
        tk.Label(frame, textvariable=self.master_path_var, bg=self.colors["panel"], fg="#BBD1E4",
                 font=("Segoe UI", 9), anchor="w").pack(side="left", fill="x", expand=True)

        RoundedButton(frame, "Browse Master Sheet", self.browse_master, width=190, height=38,
                      bg="#7657E8", icon="sheet", radius=6).pack(side="right", padx=(8, 10), pady=7)
        RoundedButton(frame, "Open Master Folder", self.open_master_folder, width=185, height=38,
                      bg="#E58A00", icon="folder", radius=6).pack(side="right", padx=8, pady=7)

    # ---------------- Summary ----------------

    def build_summary_cards(self):
        cards = tk.Frame(self.home, bg=self.colors["bg"])
        cards.pack(fill="x", padx=14, pady=3)
        self.card_vars = {}
        card_info = [
            ("total", "TOTAL DEVICES", "▣", "#1765A5"),
            ("working", "WORKING / LIVE", "✓", "#0A9B63"),
            ("offline", "OFFLINE", "×", "#BE2A4D"),
            ("nvr", "TOTAL NVR", "▤", "#7657E8"),
            ("camera", "TOTAL CAMERA", "◉", "#078FBD"),
        ]
        for key, title, icon, color in card_info:
            card = tk.Frame(cards, bg=color, height=88, highlightbackground="#2E6A93", highlightthickness=1)
            card.pack(side="left", fill="both", expand=True, padx=4)
            card.pack_propagate(False)
            tk.Label(card, text=icon, bg=color, fg="white", font=("Segoe UI Symbol", 20, "bold"), width=3).pack(side="left", padx=(8, 3))
            text_frame = tk.Frame(card, bg=color)
            text_frame.pack(side="left", fill="both", expand=True)
            tk.Label(text_frame, text=title, bg=color, fg="#E8F2F9", font=("Segoe UI", 8, "bold"), anchor="w").pack(anchor="w", pady=(12, 0))
            var = tk.StringVar(value="0")
            self.card_vars[key] = var
            tk.Label(text_frame, textvariable=var, bg=color, fg="white", font=("Segoe UI", 22, "bold"), anchor="w").pack(anchor="w")

    # ---------------- Actions ----------------

    def build_action_bar(self):
        bar = tk.Frame(self.home, bg=self.colors["bg"])
        bar.pack(fill="x", padx=14, pady=(9, 6))
        self.scan_button = RoundedButton(bar, "START CCTV SCAN", self.start_scan, width=220, height=42,
                                          bg="#0B8F5A", icon="play", radius=6)
        self.scan_button.pack(side="left", padx=(0, 8))
        RoundedButton(bar, "Open Reports Folder", self.open_reports_folder, width=195, height=42,
                      bg="#E58A00", icon="folder", radius=6).pack(side="left", padx=8)
        RoundedButton(bar, "Open Latest Report", self.open_latest_report, width=190, height=42,
                      bg="#E58A00", icon="sheet", radius=6).pack(side="left", padx=8)
        RoundedButton(bar, "Refresh Dashboard", self.refresh_dashboard, width=190, height=42,
                      bg="#1676C4", icon="refresh", radius=6).pack(side="right", padx=(8, 0))

    # ---------------- Scan status ----------------

    def build_scan_status(self):
        panel = tk.Frame(self.home, bg=self.colors["panel"], highlightbackground=self.colors["border"], highlightthickness=1, height=66)
        panel.pack(fill="x", padx=14, pady=(2, 8))
        panel.pack_propagate(False)
        top = tk.Frame(panel, bg=self.colors["panel"])
        top.pack(fill="x", padx=12, pady=(7, 2))
        self.scan_status_var = tk.StringVar(value="READY • Dashboard cleared • Press START CCTV SCAN")
        tk.Label(top, textvariable=self.scan_status_var, bg=self.colors["panel"], fg="#E3EEF7",
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left", fill="x", expand=True)
        self.scan_percent_var = tk.StringVar(value="0%")
        tk.Label(top, textvariable=self.scan_percent_var, bg=self.colors["panel"], fg="#2BE4A4",
                 font=("Segoe UI", 9, "bold")).pack(side="right")
        self.progress = ttk.Progressbar(panel, style="CCTV.Horizontal.TProgressbar", mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=12, pady=(0, 8))
        self.progress_target = 0.0
        self.progress_display = 0.0

    # ---------------- Live preview (BOTTOM) ----------------

    def build_live_preview(self):
        outer = tk.Frame(self.home, bg=self.colors["panel"], highlightbackground=self.colors["border"], highlightthickness=1)
        outer.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        title_bar = tk.Frame(outer, bg="#0E3A60", height=40)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text="LIVE SCAN PREVIEW", bg="#0E3A60", fg="white",
                 font=("Segoe UI", 10, "bold")).pack(side="left", padx=13)
        self.live_indicator = tk.Label(title_bar, text="● READY", bg="#0E3A60", fg="#91AFC6",
                                       font=("Segoe UI", 9, "bold"))
        self.live_indicator.pack(side="right", padx=13)
        self.scan_animation = tk.Canvas(title_bar, width=150, height=22, bg="#0E3A60", highlightthickness=0)
        self.scan_animation.pack(side="right", padx=5)
        self.scan_animation.create_line(8, 11, 142, 11, fill="#1E557A", width=2)
        self.scan_dot = self.scan_animation.create_oval(5, 8, 11, 14, fill="#19C37D", outline="")
        self.scan_line_x = 8

        tree_frame = tk.Frame(outer, bg=self.colors["panel"])
        tree_frame.pack(fill="both", expand=True, padx=8, pady=8)
        columns = ("device", "ip", "location", "mac", "status", "response", "remarks")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        headings = {"device":"Device Name / Camera & NVR", "ip":"IP Address", "location":"Location", "mac":"MAC Address",
                    "status":"Status", "response":"Response (ms)", "remarks":"Remarks"}
        widths = {"device":245, "ip":150, "location":205, "mac":190, "status":120, "response":125, "remarks":300}
        for col in columns:
            self.tree.heading(col, text=headings[col], anchor="center")
            self.tree.column(col, width=widths[col], minwidth=90, anchor="center" if col in ("status", "response") else "w", stretch=True)
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.tag_configure("working", background="#DDF7E9", foreground="#096B42")
        self.tree.tag_configure("offline", background="#FDE4E9", foreground="#A31E3C")
        self.tree.tag_configure("scanning", background="#FFF6D8", foreground="#735A00")
        self.animate_scan_line()

    def animate_scan_line(self):
        if not hasattr(self, "scan_animation") or not self.scan_animation.winfo_exists():
            return
        self.scan_line_x += 3
        if self.scan_line_x > 142:
            self.scan_line_x = 8
        self.scan_animation.coords(self.scan_dot, self.scan_line_x - 3, 8, self.scan_line_x + 3, 14)
        self.root.after(55, self.animate_scan_line)

    # ---------------- Smooth progress animation ----------------

    def animate_progress(self):
        if not hasattr(self, "progress") or not self.progress.winfo_exists():
            return
        diff = self.progress_target - self.progress_display
        if abs(diff) > 0.15:
            self.progress_display += max(-1.8, min(1.8, diff * 0.22))
        else:
            self.progress_display = self.progress_target
        self.progress["value"] = self.progress_display
        self.scan_percent_var.set(f"{int(round(self.progress_display))}%")
        if self.scanning or self.progress_display < self.progress_target:
            self.root.after(35, self.animate_progress)

    # ---------------- Master actions ----------------

    def browse_master(self):
        path = filedialog.askopenfilename(
            title="Select CCTV Master Sheet",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )

        if not path:
            return

        try:
            self.load_master(Path(path))
        except Exception as exc:
            messagebox.showerror(
                "Master Sheet Error",
                str(exc),
            )

    def load_master(self, path, silent=False):
        devices = load_master_sheet(path)

        self.master_path = Path(path)
        self.devices = devices

        self.master_path_var.set(str(self.master_path))

        if not silent:
            messagebox.showinfo(
                "Master Sheet Loaded",
                f"Master Sheet loaded successfully.\n\n"
                f"Devices found: {len(devices)}\n"
                f"Sequence preserved exactly as Excel.",
            )

    def open_master_folder(self):
        folder = (
            self.master_path.parent
            if self.master_path
            else APP_DIR
        )
        self.open_path(folder)

    # ---------------- Scan ----------------

    def start_scan(self):
        if self.scanning:
            return

        if not self.master_path:
            path = filedialog.askopenfilename(
                title="Select CCTV Master Sheet",
                filetypes=[("Excel files", "*.xlsx *.xlsm")],
            )
            if not path:
                return

            try:
                self.load_master(Path(path), silent=True)
            except Exception as exc:
                messagebox.showerror(
                    "Master Sheet Error",
                    str(exc),
                )
                return

        if not self.devices:
            messagebox.showwarning(
                "No Devices",
                "No devices are available in the Master Sheet.",
            )
            return

        self.scanning = True
        self.results = []
        self.scan_started = datetime.now()

        self.scan_button.set_enabled(False)

        self.progress_target = 0.0
        self.progress_display = 0.0
        self.progress["value"] = 0
        self.scan_percent_var.set("0%")
        self.scan_status_var.set(
            f"SCANNING • 0/{len(self.devices)} devices checked"
        )
        self.live_indicator.configure(
            text="● SCANNING",
            fg="#19E6A2",
        )

        self.clear_preview()
        self.populate_preview_in_master_sequence()

        # Reset cards before scan begins.
        self.set_card_values(0, 0, 0, 0, 0)

        self.root.after(35, self.animate_progress)
        worker = threading.Thread(
            target=self.run_scan_worker,
            daemon=True,
        )
        worker.start()

    def run_scan_worker(self):
        total = len(self.devices)
        completed = 0
        result_map = {}

        with ThreadPoolExecutor(max_workers=SCAN_WORKERS) as executor:
            future_map = {
                executor.submit(scan_device, device): index
                for index, device in enumerate(self.devices)
            }

            for future in as_completed(future_map):
                source_index = future_map[future]

                try:
                    result = future.result()
                except Exception as exc:
                    device = self.devices[source_index]
                    result = {
                        **device,
                        "mac": "",
                        "status": "OFFLINE",
                        "response_ms": None,
                        "remarks": f"Scan Error: {exc}",
                    }

                result_map[source_index] = result
                completed += 1

                self.ui_queue.put(
                    ("device", source_index, result, completed, total)
                )

        # Rebuild the exact Excel sequence.
        # IMPORTANT: NVR and Camera statuses are independent.
        # A failed NVR must NOT force its cameras offline.
        ordered = [
            result_map[i]
            for i in range(total)
            if i in result_map
        ]
        self.ui_queue.put(("complete", ordered))

    def process_queue(self):
        try:
            while True:
                item = self.ui_queue.get_nowait()

                if item[0] == "device":
                    _, source_index, result, completed, total = item
                    self.add_live_result(result, completed, total)

                elif item[0] == "complete":
                    self.finish_scan(item[1])

        except queue.Empty:
            pass

        if self.root.winfo_exists():
            self.root.after(80, self.process_queue)

    def populate_preview_in_master_sequence(self):
        """Create the live preview in the exact Master Sheet sequence."""
        self.preview_items = {}
        for index, device in enumerate(self.devices):
            item_id = self.tree.insert(
                "",
                "end",
                values=(
                    device["device_name"],
                    device["ip"],
                    device["location"],
                    "-",
                    "SCANNING",
                    "-",
                    "Waiting for scan...",
                ),
                tags=("scanning",),
            )
            self.preview_items[index] = item_id

    def add_live_result(self, result, completed, total):
        # Update the existing row so parallel scanning never changes
        # the Master Sheet sequence shown in the preview.
        source_index = result.get("source_index", result.get("excel_row", 2) - 2)
        item_id = self.preview_items.get(source_index)

        if item_id:
            response = (
                f"{result['response_ms']:.1f}"
                if isinstance(result["response_ms"], (int, float))
                else "-"
            )
            tag = "working" if result["status"] == "WORKING" else "offline"

            self.tree.item(
                item_id,
                values=(
                    excel_safe_text(result["device_name"]),
                    excel_safe_text(result["ip"]),
                    excel_safe_text(result["location"]),
                    excel_safe_text(result["mac"] or "-"),
                    result["status"],
                    response,
                    excel_safe_text(result["remarks"]),
                ),
                tags=(tag,),
            )

        completed_results = []
        for child in self.tree.get_children():
            values = self.tree.item(child, "values")
            if values and values[4] in ("WORKING", "OFFLINE"):
                completed_results.append(values)

        working = sum(1 for values in completed_results if values[4] == "WORKING")
        offline = sum(1 for values in completed_results if values[4] == "OFFLINE")
        completed_count = len(completed_results)

        nvr = sum(
            1 for values in completed_results
            if clean(values[0]).upper().startswith("NVR")
        )
        camera = completed_count - nvr

        self.set_card_values(
            completed_count,
            working,
            offline,
            nvr,
            camera,
        )

        self.progress_target = (completed / total) * 100
        self.scan_status_var.set(
            f"SCANNING • {completed}/{total} devices checked • "
            f"{result['device_name']} • {result['ip']} • {result['status']}"
        )

        if item_id:
            self.tree.see(item_id)

    def refresh_preview_from_results(self, ordered_results):
        """Refresh preview rows after final status/dependency calculation."""
        for result in ordered_results:
            source_index = result.get("source_index")
            item_id = self.preview_items.get(source_index)
            if not item_id:
                continue
            response = (
                f"{result['response_ms']:.1f}"
                if isinstance(result.get("response_ms"), (int, float))
                else "-"
            )
            tag = "working" if result["status"] == "WORKING" else "offline"
            self.tree.item(
                item_id,
                values=(
                    result["device_name"], result["ip"], result["location"],
                    result["mac"] or "-", result["status"], response, result["remarks"]
                ),
                tags=(tag,),
            )

    def finish_scan(self, ordered_results):
        self.results = ordered_results
        self.scanning = False
        self.refresh_preview_from_results(ordered_results)

        total = len(ordered_results)
        working = sum(1 for r in ordered_results if r["status"] == "WORKING")
        offline = total - working
        nvr = sum(1 for r in ordered_results if is_nvr(r["device_name"]))
        camera = total - nvr

        self.set_card_values(total, working, offline, nvr, camera)
        self.progress_target = 100.0
        self.progress_display = 100.0
        self.progress["value"] = 100
        self.scan_percent_var.set("100%")
        self.scan_button.set_enabled(True)
        self.live_indicator.configure(text="● SCAN COMPLETE", fg="#19E6A2")
        self.scan_status_var.set(
            f"SCAN COMPLETE • {total} devices • {working} WORKING • {offline} OFFLINE"
        )

        try:
            report_path, error = create_report(ordered_results, self.master_path)

            if report_path:
                self.latest_report_path = Path(report_path)
                match = re.search(r"Run-(\d{2})_", report_path.name)
                run_no = int(match.group(1)) if match else 0
                self.scan_status_var.set(
                    f"SCAN COMPLETE • REPORT GENERATED • Run-{run_no:02d} • {report_path.name}"
                )
                self.show_scan_complete_dialog(
                    total, working, offline, nvr, camera, report_path
                )
            else:
                self.show_report_limit_dialog(error)

        except Exception as exc:
            messagebox.showerror(
                "Report Generation Error",
                f"Scan completed, but report generation failed.\n\n{exc}",
            )

    def show_scan_complete_dialog(self, total, working, offline, nvr, camera, report_path):
        """Professional, colourful report-complete dialog with direct report opening."""
        dialog = tk.Toplevel(self.root)
        dialog.title("CCTV Scan Complete")
        dialog.configure(bg="#F5F9FD")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        width, height = 820, 560
        self.center_window(dialog, width, height)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)

        header = tk.Frame(dialog, bg="#0A66C2", height=78)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="✓", bg="#0A66C2", fg="white",
                 font=("Segoe UI Symbol", 34, "bold")).pack(side="left", padx=(26, 10))
        tk.Label(header, text="CCTV Scan Complete", bg="#0A66C2", fg="white",
                 font=("Segoe UI", 22, "bold")).pack(side="left")

        success = tk.Frame(dialog, bg="#E6F7EA", highlightbackground="#A9E6B7", highlightthickness=1)
        success.pack(fill="x", padx=22, pady=(18, 12))
        tk.Label(success, text="SCAN COMPLETED SUCCESSFULLY", bg="#E6F7EA", fg="#137333",
                 font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=18, pady=(12, 2))
        tk.Label(success, text="All devices were checked and the report has been generated.",
                 bg="#E6F7EA", fg="#245B38", font=("Segoe UI", 10)).pack(anchor="w", padx=18, pady=(0, 12))

        cards = tk.Frame(dialog, bg="#F5F9FD")
        cards.pack(fill="x", padx=22, pady=4)
        card_data = [
            ("TOTAL DEVICES", total, "#EAF3FF", "#1765A5"),
            ("WORKING / LIVE", working, "#E7F8EE", "#0B8F5A"),
            ("OFFLINE", offline, "#FCE7EB", "#B4234D"),
            ("NVR", nvr, "#F0EAFF", "#6941C6"),
            ("CAMERA", camera, "#E6F7FC", "#078FBD"),
        ]
        for title, value, bg, fg in card_data:
            card = tk.Frame(cards, bg=bg, height=78, highlightbackground="#C8D7E6", highlightthickness=1)
            card.pack(side="left", fill="both", expand=True, padx=4)
            card.pack_propagate(False)
            tk.Label(card, text=str(value), bg=bg, fg=fg, font=("Segoe UI", 22, "bold")).pack(pady=(9, 0))
            tk.Label(card, text=title, bg=bg, fg="#34495E", font=("Segoe UI", 8, "bold")).pack()

        report_box = tk.Frame(dialog, bg="#EEF5FC", highlightbackground="#B8D3EE", highlightthickness=1)
        report_box.pack(fill="x", padx=22, pady=14)
        tk.Label(report_box, text="REPORT LOCATION", bg="#EEF5FC", fg="#145A8D",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
        tk.Label(report_box, text=str(report_path), bg="#EEF5FC", fg="#26384A",
                 font=("Consolas", 9), wraplength=690, justify="left").pack(anchor="w", padx=14, pady=(0, 10))

        buttons = tk.Frame(dialog, bg="#F5F9FD")
        buttons.pack(fill="x", padx=22, pady=(0, 18))

        def open_report_and_close():
            self.open_path(report_path)
            dialog.destroy()

        RoundedButton(buttons, "Open Report", open_report_and_close, width=170, height=40,
                      bg="#0B8F5A", icon="sheet", radius=7).pack(side="left")
        RoundedButton(buttons, "Open Report Folder", lambda: self.open_path(report_path.parent), width=200, height=40,
                      bg="#E58A00", icon="folder", radius=7).pack(side="left", padx=10)
        RoundedButton(buttons, "OK", dialog.destroy, width=110, height=40,
                      bg="#1676C4", icon="✓", radius=7).pack(side="right")

    def show_report_limit_dialog(self, error):
        dialog = tk.Toplevel(self.root)
        dialog.title("Daily Report Limit")
        dialog.configure(bg="#FFF8E8")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 560, 250)
        tk.Label(dialog, text="REPORT LIMIT REACHED", bg="#F59E0B", fg="white",
                 font=("Segoe UI", 17, "bold"), pady=14).pack(fill="x")
        tk.Label(dialog, text="", bg="#FFF8E8", height=1).pack()
        tk.Label(dialog, text=error, bg="#FFF8E8", fg="#5B4500",
                 font=("Segoe UI", 11), justify="center").pack(pady=12)
        RoundedButton(dialog, "OK", dialog.destroy, width=110, height=40,
                      bg="#E58A00", icon="✓", radius=7).pack(pady=15)

    @staticmethod
    def center_window(window, width, height):
        window.update_idletasks()
        screen_w = window.winfo_screenwidth()
        screen_h = window.winfo_screenheight()
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    # ---------------- Dashboard controls ----------------

    def set_card_values(
        self,
        total,
        working,
        offline,
        nvr,
        camera,
    ):
        self.card_vars["total"].set(str(total))
        self.card_vars["working"].set(str(working))
        self.card_vars["offline"].set(str(offline))
        self.card_vars["nvr"].set(str(nvr))
        self.card_vars["camera"].set(str(camera))

    def clear_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.preview_items = {}

    def refresh_dashboard(self):
        if self.scanning:
            messagebox.showwarning(
                "Scan Running",
                "Please wait until the current scan is complete.",
            )
            return

        self.results = []
        self.clear_preview()
        self.set_card_values(0, 0, 0, 0, 0)

        self.progress_target = 0.0
        self.progress_display = 0.0
        self.progress["value"] = 0
        self.scan_percent_var.set("0%")
        self.scan_status_var.set(
            "READY • Dashboard cleared • Press START CCTV SCAN"
        )
        self.live_indicator.configure(
            text="● READY",
            fg="#8FA8BA",
        )

    # ---------------- Report folders ----------------

    def get_report_root(self):
        if self.master_path:
            return self.master_path.parent / REPORT_ROOT_NAME
        return APP_DIR / REPORT_ROOT_NAME

    def open_reports_folder(self):
        root = self.get_report_root()
        root.mkdir(parents=True, exist_ok=True)
        self.open_path(root)

    def open_latest_report(self):
        # First use the exact report generated in this application session.
        if self.latest_report_path and Path(self.latest_report_path).is_file():
            self.open_path(self.latest_report_path)
            return

        root = self.get_report_root()
        candidates = []
        if root.exists():
            candidates.extend(root.rglob("CCTV_Health_Report_*.xlsx"))

        # Fallback: search the master-sheet parent for the report folder.
        if not candidates and self.master_path:
            parent = self.master_path.parent
            candidates.extend(parent.rglob("CCTV_Health_Report_*.xlsx"))

        candidates = [p for p in candidates if p.is_file()]
        if not candidates:
            messagebox.showinfo(
                "No Reports",
                "No CCTV reports have been generated yet.\n\nRun START CCTV SCAN first.",
            )
            return

        latest = max(candidates, key=lambda p: p.stat().st_mtime)
        self.latest_report_path = latest
        self.open_path(latest)

    @staticmethod
    def open_path(path):
        path = Path(path)

        try:
            if os.name == "nt":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror(
                "Open Error",
                f"Unable to open:\n{path}\n\n{exc}",
            )


# ----------------------------- Main -----------------------------

def main():
    if os.name != "nt":
        # The monitoring engine is designed for Windows because
        # the target environment uses Windows ping/ARP/PowerShell.
        pass

    root = tk.Tk()

    # Windows taskbar/title icon is intentionally kept dependency-free.
    try:
        root.iconname(APP_TITLE)
    except Exception:
        pass

    app = CCTVHealthApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
