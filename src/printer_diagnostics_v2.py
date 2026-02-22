"""
=============================================================
  PrinterDiag - Ferramenta de Diagnóstico de Impressoras
  Desenvolvida para Analistas de Suporte
=============================================================
  Desenvolvido por: Thomaz Arthur
  Versão: 2.0
=============================================================
Requisitos: Python 3.8+, pywin32
Instale com: python -m pip install pywin32
=============================================================
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import threading
import ipaddress
import re
import json
import time
from datetime import datetime


# ─────────────────────────────────────────────
#  CORES E ESTILO
# ─────────────────────────────────────────────
THEME = {
    "bg":         "#0f1117",
    "surface":    "#1a1d27",
    "surface2":   "#252836",
    "accent":     "#4f8ef7",
    "accent2":    "#6c63ff",
    "success":    "#22c55e",
    "warning":    "#f59e0b",
    "error":      "#ef4444",
    "text":       "#e2e8f0",
    "text_muted": "#64748b",
    "border":     "#2d3148",
    "info":       "#7dd3fc",
}

# ─────────────────────────────────────────────
#  BASE DE DRIVERS CONHECIDOS POR FABRICANTE
# ─────────────────────────────────────────────
DRIVER_HINTS = {
    "hp":      "HP Universal Printing PCL 6",
    "epson":   "EPSON Universal Print Driver",
    "brother": "Brother Printer Driver",
    "canon":   "Canon Generic Plus PCL6",
    "zebra":   "ZDesigner",
    "samsung": "Samsung Universal Print Driver",
    "ricoh":   "RICOH PCL6 UniversalDriver",
    "xerox":   "Xerox Global Print Driver PCL6",
    "lexmark": "Lexmark Universal v2",
    "kyocera": "KYOCERA Classic Universal Printer",
    "oki":     "OKI Universal PCL6",
}


# ─────────────────────────────────────────────
#  FUNÇÕES UTILITÁRIAS
# ─────────────────────────────────────────────

def run_ps(command: str) -> str:
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True, text=True, timeout=30
        )
        return result.stdout.strip()
    except Exception as e:
        return f"ERRO: {e}"


def is_spooler_running() -> bool:
    out = run_ps("(Get-Service -Name Spooler).Status")
    return "Running" in out


def get_server_ip() -> str:
    ps = ("Get-WmiObject Win32_NetworkAdapterConfiguration "
          "| Where-Object {$_.IPAddress} "
          "| Select-Object -ExpandProperty IPAddress -First 1")
    return run_ps(ps).strip()


# ─────────────────────────────────────────────
#  COLETA DE IMPRESSORAS
# ─────────────────────────────────────────────

def get_all_printers() -> list:
    ps = """
    Get-WmiObject -Class Win32_Printer | Select-Object Name, PortName, DriverName,
        WorkOffline, PrinterStatus, DetectedErrorState, ExtendedDetectedErrorState,
        Shared, ShareName, PrinterState |
    ConvertTo-Json -Depth 3
    """
    output = run_ps(ps)
    if not output or "ERRO" in output:
        return []
    try:
        data = json.loads(output)
        if isinstance(data, dict):
            data = [data]
        return data
    except json.JSONDecodeError:
        return []


def get_printer_jobs(printer_name: str) -> list:
    ps = f"""
    Get-PrintJob -PrinterName '{printer_name}' -ErrorAction SilentlyContinue |
    Select-Object Id, DocumentName, JobStatus, SubmittedTime |
    ConvertTo-Json -Depth 2
    """
    output = run_ps(ps)
    if not output or "ERRO" in output:
        return []
    try:
        data = json.loads(output)
        return [data] if isinstance(data, dict) else data
    except:
        return []


# ─────────────────────────────────────────────
#  VERIFICAÇÕES ESPECÍFICAS
# ─────────────────────────────────────────────

def check_driver_validity(printer: dict) -> dict:
    result = {
        "installed_driver": printer.get("DriverName", ""),
        "suggested_driver": None,
        "driver_ok": True,
        "details": [],
        "fix_available": False,
    }

    name   = printer.get("Name", "").lower()
    driver = printer.get("DriverName", "")

    # Verifica se o driver existe no sistema
    ps = f"Get-PrinterDriver -Name '{driver}' -ErrorAction SilentlyContinue | Select-Object Name | ConvertTo-Json"
    out = run_ps(ps)
    if not out or "ERRO" in out or "{" not in out:
        result["driver_ok"] = False
        result["details"].append(f"✗ Driver '{driver}' NÃO está instalado no sistema.")
        result["fix_available"] = True
    
    # Identifica fabricante pelo nome da impressora
    detected_brand = None
    for brand in DRIVER_HINTS:
        if brand in name:
            detected_brand = brand
            break

    if detected_brand:
        suggested = DRIVER_HINTS[detected_brand]
        result["suggested_driver"] = suggested
        if detected_brand not in driver.lower():
            if result["driver_ok"]:  # só adiciona se não adicionou erro antes
                result["details"].append(f"⚠ Driver instalado: '{driver}'")
            result["details"].append(f"✦ Driver sugerido para {detected_brand.upper()}: '{suggested}'")
            result["driver_ok"] = False
            result["fix_available"] = True
        else:
            if result["driver_ok"]:
                result["details"].append(f"✓ Driver '{driver}' compatível com o fabricante.")
    else:
        if result["driver_ok"]:
            result["details"].append(f"✓ Driver '{driver}' instalado.")

    return result


def check_port_validity(printer: dict) -> dict:
    result = {
        "port": printer.get("PortName", ""),
        "port_ok": True,
        "suggested_port": None,
        "details": [],
        "fix_available": False,
        "is_usb": False,
        "is_network": False,
    }

    port = printer.get("PortName", "")

    # ── USB ──────────────────────────────────────
    if "USB" in port.upper():
        result["is_usb"] = True
        ps_check = f"Get-PrinterPort -Name '{port}' -ErrorAction SilentlyContinue | Select-Object Name | ConvertTo-Json"
        out = run_ps(ps_check)

        if not out or "{" not in out:
            result["port_ok"] = False
            result["details"].append(f"✗ Porta '{port}' não encontrada no sistema.")
            # Busca portas USB ativas
            ps_list = "Get-PrinterPort | Where-Object {$_.Name -like 'USB*'} | Select-Object Name | ConvertTo-Json"
            out_list = run_ps(ps_list)
            try:
                pd = json.loads(out_list)
                if isinstance(pd, dict):
                    pd = [pd]
                available = [p.get("Name", "") for p in pd if p.get("Name")]
                if available:
                    result["suggested_port"] = available[0]
                    result["details"].append(f"✦ Porta sugerida: '{available[0]}'")
                    result["fix_available"] = True
                else:
                    result["details"].append("✗ Nenhuma porta USB ativa encontrada.")
            except:
                result["details"].append("✗ Não foi possível listar portas USB.")
        else:
            result["details"].append(f"✓ Porta USB '{port}' registrada.")
            # Verifica se a porta ainda está entre as ativas
            ps_list = "Get-PrinterPort | Where-Object {$_.Name -like 'USB*'} | Select-Object Name | ConvertTo-Json"
            out_list = run_ps(ps_list)
            try:
                pd = json.loads(out_list)
                if isinstance(pd, dict):
                    pd = [pd]
                available = [p.get("Name", "") for p in pd if p.get("Name")]
                if available and port not in available:
                    result["port_ok"] = False
                    result["suggested_port"] = available[-1]
                    result["details"].append(f"⚠ Porta '{port}' não está entre as portas USB ativas.")
                    result["details"].append(f"✦ Porta sugerida: '{available[-1]}'")
                    result["fix_available"] = True
            except:
                pass

    # ── REDE ─────────────────────────────────────
    elif re.search(r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', port) or "TCP" in port.upper():
        result["is_network"] = True
        ip_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', port)
        if not ip_match:
            ps_ip = (f"Get-WmiObject Win32_TCPIPPrinterPort "
                     f"| Where-Object {{$_.Name -eq '{port}'}} "
                     f"| Select-Object HostAddress | ConvertTo-Json")
            out_ip = run_ps(ps_ip)
            try:
                d = json.loads(out_ip)
                ip_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', str(d))
            except:
                pass

        if ip_match:
            printer_ip = ip_match.group(1)
            result["printer_ip"] = printer_ip
            ping = subprocess.run(["ping", "-n", "1", "-w", "1500", printer_ip],
                                  capture_output=True, text=True)
            if "TTL=" in ping.stdout:
                result["details"].append(f"✓ Porta '{port}' → {printer_ip} responde ao ping.")
            else:
                result["port_ok"] = False
                result["details"].append(f"✗ Porta '{port}' → {printer_ip} NÃO responde.")
                result["details"].append("✦ Verifique o IP atual da impressora e atualize a porta.")
                result["fix_available"] = True
        else:
            result["details"].append(f"⚠ Não foi possível determinar o IP da porta '{port}'.")
    else:
        result["details"].append(f"✓ Porta local '{port}'.")

    return result


def check_network_range(printer: dict) -> dict:
    result = {
        "has_issue": False, "details": [],
        "printer_ip": None, "server_ip": None,
        "fix_available": False, "subnet_mismatch": False,
        "ping_ok": True,
    }

    port = printer.get("PortName", "")
    ip_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', port)
    if not ip_match:
        ps = (f"Get-WmiObject Win32_TCPIPPrinterPort "
              f"| Where-Object {{$_.Name -eq '{port}'}} "
              f"| Select-Object HostAddress | ConvertTo-Json")
        output = run_ps(ps)
        try:
            data = json.loads(output)
            ip_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', str(data))
        except:
            pass

    if not ip_match:
        result["details"].append("Porta não é de rede TCP/IP.")
        return result

    printer_ip = ip_match.group(1)
    result["printer_ip"] = printer_ip
    server_ip = get_server_ip()
    result["server_ip"] = server_ip

    try:
        p_net = ipaddress.IPv4Network(printer_ip + "/24", strict=False)
        s_net = ipaddress.IPv4Network(server_ip + "/24", strict=False)

        if p_net != s_net:
            result["has_issue"] = True
            result["subnet_mismatch"] = True
            result["fix_available"] = True
            result["details"].append(
                f"⚠ Sub-rede diferente: impressora {printer_ip} / servidor {server_ip}.")
        else:
            result["details"].append(f"✓ Mesma sub-rede: {printer_ip} / {server_ip}.")

        ping = subprocess.run(["ping", "-n", "1", "-w", "1500", printer_ip],
                              capture_output=True, text=True)
        if "TTL=" in ping.stdout:
            result["details"].append(f"✓ Responde ao ping ({printer_ip}).")
            result["ping_ok"] = True
        else:
            result["has_issue"] = True
            result["ping_ok"] = False
            result["details"].append(f"✗ Sem resposta ao ping ({printer_ip}).")
    except Exception as e:
        result["details"].append(f"Erro ao verificar rede: {e}")

    return result


def check_usb_driver(printer: dict) -> dict:
    result = {"has_issue": False, "details": [], "fix_available": False}
    port   = printer.get("PortName", "")
    driver = printer.get("DriverName", "")

    ps = (f"Get-PrinterDriver -Name '{driver}' -ErrorAction SilentlyContinue "
          f"| Select-Object Name | ConvertTo-Json")
    output = run_ps(ps)
    if not output or "ERRO" in output or "{" not in output:
        result["has_issue"] = True
        result["details"].append(f"Driver '{driver}' não encontrado.")
        result["fix_available"] = True

    ps_port = f"Get-PrinterPort -Name '{port}' -ErrorAction SilentlyContinue"
    if not run_ps(ps_port) or "ERRO" in run_ps(ps_port):
        result["has_issue"] = True
        result["details"].append(f"Porta USB '{port}' não responde.")
        result["fix_available"] = True

    return result


def get_error_description(printer: dict) -> str:
    status_map = {
        1: "Outro",          2: "✓ Normal",
        3: "⚠ Inativa",     4: "✗ Erro",
        5: "✗ Desconhecida", 6: "✗ Sem impressão",
        7: "✗ Offline",     8: "✗ I/O ativo",
        9: "⚠ Ocupada",     10: "⚠ Imprimindo",
        11: "⚠ Aquecendo",  12: "⚠ Sem seleção",
        13: "⚠ Sem papel",
    }
    error_map = {
        0: "Nenhum",          2: "✗ Papel atolado",
        4: "✗ Sem papel",     8: "✗ Sem toner/tinta",
        16: "⚠ Toner baixo",  32: "✗ Erro saída",
        64: "⚠ Offline",     128: "✗ Requer técnico",
        256: "✗ Erro I/O",   512: "✗ Erro envelope",
        1024: "✗ Papel personalizado", 2048: "✗ Erro bandeja",
        4096: "⚠ Tampa aberta", 8192: "✗ Erro referência",
    }
    state_bits = {
        1: "⚠ Pausada", 2: "⚠ Erro", 8: "⚠ Papel atolado",
        16: "✗ Sem papel", 64: "✗ Offline", 512: "⚠ Imprimindo",
    }

    sc = printer.get("PrinterStatus", 2)
    ec = printer.get("DetectedErrorState", 0)
    st = printer.get("PrinterState", 0) or 0

    status_txt = status_map.get(sc, f"Código {sc}")
    error_txt  = error_map.get(ec, f"Erro {ec}")
    state_parts = [v for k, v in state_bits.items() if k > 0 and (st & k)]
    state_txt  = " | ".join(state_parts) if state_parts else ""

    out = f"Status: {status_txt} | Erro: {error_txt}"
    if state_txt:
        out += f" | Estado: {state_txt}"
    return out


# ─────────────────────────────────────────────
#  DIAGNÓSTICO COMPLETO
# ─────────────────────────────────────────────

def diagnose_printer(printer: dict) -> dict:
    port       = printer.get("PortName", "")
    is_usb     = "USB" in port.upper()
    is_network = bool(re.search(r'\d{1,3}\.\d{1,3}', port)) or "TCP" in port.upper()

    diagnosis = {
        "name":          printer.get("Name", "Desconhecida"),
        "port":          port,
        "driver":        printer.get("DriverName", ""),
        "offline":       printer.get("WorkOffline", False),
        "paused":        False,
        "status_text":   get_error_description(printer),
        "type":          "USB" if is_usb else ("Rede" if is_network else "Local"),
        "issues":        [],
        "fixes":         [],
        "usb_check":     None,
        "network_check": None,
        "driver_check":  None,
        "port_check":    None,
    }

    # Spooler
    if not is_spooler_running():
        diagnosis["issues"].append("Serviço Spooler NÃO está em execução!")
        diagnosis["fixes"].append("restart_spooler")

    # Offline
    if printer.get("WorkOffline"):
        diagnosis["issues"].append("Impressora definida como OFFLINE no Windows.")
        diagnosis["fixes"].append("set_online")

    # Pausada (PrinterState bit 1)
    state = printer.get("PrinterState", 0) or 0
    if state & 1:
        diagnosis["paused"] = True
        diagnosis["issues"].append("Impressora PAUSADA — trabalhos não são enviados.")
        diagnosis["fixes"].append("resume_printer")

    # Status geral
    status = printer.get("PrinterStatus", 2)
    if status == 7:
        diagnosis["issues"].append("Windows reporta impressora OFFLINE (código 7).")
        if "set_online" not in diagnosis["fixes"]:
            diagnosis["fixes"].append("set_online")
    elif status == 4:
        diagnosis["issues"].append("Windows reporta ERRO genérico na impressora.")
    elif status not in (2, 10, 9):
        diagnosis["issues"].append(f"Status anormal detectado (código {status}).")

    # Fila de impressão
    jobs = get_printer_jobs(printer.get("Name", ""))
    if jobs:
        diagnosis["issues"].append(f"{len(jobs)} trabalho(s) travado(s) na fila.")
        diagnosis["fixes"].append("clear_queue")

    # Driver
    drv = check_driver_validity(printer)
    diagnosis["driver_check"] = drv
    if not drv["driver_ok"]:
        diagnosis["issues"].append(
            next((d for d in drv["details"] if "✗" in d or "⚠" in d), "Problema no driver."))
        if drv["fix_available"]:
            diagnosis["fixes"].append("fix_driver")

    # Porta
    pc = check_port_validity(printer)
    diagnosis["port_check"] = pc
    if not pc["port_ok"]:
        msg = f"Porta '{port}' com problema."
        if pc.get("suggested_port"):
            msg += f" Sugerida: '{pc['suggested_port']}'"
        diagnosis["issues"].append(msg)
        if pc["fix_available"]:
            diagnosis["fixes"].append("fix_port")

    # USB
    if is_usb:
        usb = check_usb_driver(printer)
        diagnosis["usb_check"] = usb
        if usb["has_issue"]:
            for d in usb["details"]:
                if d not in diagnosis["issues"]:
                    diagnosis["issues"].append(d)
            if usb["fix_available"] and "fix_driver" not in diagnosis["fixes"]:
                diagnosis["fixes"].append("fix_driver")

    # Rede
    elif is_network:
        net = check_network_range(printer)
        diagnosis["network_check"] = net
        if net["has_issue"]:
            for d in net["details"]:
                if "✗" in d and d not in diagnosis["issues"]:
                    diagnosis["issues"].append(d)
            if net.get("subnet_mismatch"):
                diagnosis["fixes"].append("subnet_guide")
            elif not net.get("ping_ok", True):
                diagnosis["fixes"].append("ping_fail_guide")

    # Remove duplicatas mantendo ordem
    seen = set()
    diagnosis["fixes"] = [f for f in diagnosis["fixes"] if not (f in seen or seen.add(f))]

    return diagnosis


# ─────────────────────────────────────────────
#  CORREÇÕES AUTOMÁTICAS
# ─────────────────────────────────────────────

def fix_set_online(printer_name: str) -> str:
    # Método 1 — cmdlet
    run_ps(f"Set-Printer -Name '{printer_name}' -WorkOffline $false -ErrorAction SilentlyContinue")
    time.sleep(1)
    # Método 2 — WMI (fallback robusto)
    run_ps(f"""
$p = Get-WmiObject -Class Win32_Printer | Where-Object {{$_.Name -eq '{printer_name}'}}
if ($p) {{ $p.WorkOffline = $false; $p.Put() }}
""")
    time.sleep(0.5)
    # Resume fila também
    run_ps(f"Resume-PrintQueue -Name '{printer_name}' -ErrorAction SilentlyContinue")
    return f"Impressora '{printer_name}' definida como ONLINE e fila retomada."


def fix_resume_printer(printer_name: str) -> str:
    run_ps(f"Resume-PrintQueue -Name '{printer_name}' -ErrorAction SilentlyContinue")
    run_ps(f"""
$p = Get-WmiObject -Class Win32_Printer | Where-Object {{$_.Name -eq '{printer_name}'}}
if ($p) {{ $p.ResumePrinter() }}
""")
    return f"Impressora '{printer_name}' retomada."


def fix_clear_queue(printer_name: str) -> str:
    run_ps(f"""
$jobs = Get-PrintJob -PrinterName '{printer_name}' -ErrorAction SilentlyContinue
if ($jobs) {{ $jobs | Remove-PrintJob -ErrorAction SilentlyContinue }}
""")
    # Limpa arquivos spool como fallback
    sp = r"C:\Windows\System32\spool\PRINTERS"
    run_ps(f"""
Stop-Service Spooler -Force -ErrorAction SilentlyContinue
Remove-Item '{sp}\\*.SHD' -ErrorAction SilentlyContinue
Remove-Item '{sp}\\*.SPL' -ErrorAction SilentlyContinue
Start-Service Spooler -ErrorAction SilentlyContinue
""")
    return f"Fila de '{printer_name}' limpa e Spooler reiniciado."


def fix_restart_spooler() -> str:
    sp = r"C:\Windows\System32\spool\PRINTERS"
    run_ps(f"""
Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Remove-Item '{sp}\\*.SHD' -ErrorAction SilentlyContinue
Remove-Item '{sp}\\*.SPL' -ErrorAction SilentlyContinue
Start-Service -Name Spooler -ErrorAction SilentlyContinue
""")
    return "Spooler reiniciado e temporários limpos."


def fix_driver(printer_name: str, driver_name: str, port: str) -> str:
    run_ps(f"Remove-Printer -Name '{printer_name}' -ErrorAction SilentlyContinue")
    time.sleep(1)
    run_ps(f"""
Add-Printer -Name '{printer_name}' -DriverName '{driver_name}' -PortName '{port}' -ErrorAction SilentlyContinue
""")
    return f"Driver '{driver_name}' reinstalado para '{printer_name}'."


def fix_port(printer_name: str, new_port: str) -> str:
    run_ps(f"Set-Printer -Name '{printer_name}' -PortName '{new_port}' -ErrorAction SilentlyContinue")
    run_ps(f"""
$p = Get-WmiObject -Class Win32_Printer | Where-Object {{$_.Name -eq '{printer_name}'}}
if ($p) {{ $p.PortName = '{new_port}'; $p.Put() }}
""")
    return f"Porta de '{printer_name}' alterada para '{new_port}'."


# ─────────────────────────────────────────────
#  INTERFACE GRÁFICA
# ─────────────────────────────────────────────

class PrinterDiagApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PrinterDiag — Diagnóstico de Impressoras  v2.0")
        self.geometry("1160x780")
        self.configure(bg=THEME["bg"])
        self.resizable(True, True)
        self.printers_data    = []
        self.diagnoses        = {}
        self.selected_printer = None
        self._build_ui()

    # ── BUILD UI ──────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=THEME["surface"], height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="🖨  PrinterDiag", font=("Segoe UI", 16, "bold"),
                 bg=THEME["surface"], fg=THEME["accent"]).pack(side="left", padx=20, pady=12)
        tk.Label(hdr, text="Diagnóstico Automático de Impressoras  •  v2.0",
                 font=("Segoe UI", 10), bg=THEME["surface"],
                 fg=THEME["text_muted"]).pack(side="left", pady=12)

        tk.Button(hdr, text="⟳  Escanear", command=self._scan_printers,
                  bg=THEME["accent"], fg="white", font=("Segoe UI", 10, "bold"),
                  relief="flat", cursor="hand2", padx=14, pady=6
                  ).pack(side="right", padx=16, pady=10)
        tk.Button(hdr, text="↺  Reiniciar Spooler", command=self._restart_spooler,
                  bg=THEME["surface2"], fg=THEME["text"], font=("Segoe UI", 10),
                  relief="flat", cursor="hand2", padx=12, pady=6
                  ).pack(side="right", padx=4, pady=10)

        # Body
        body = tk.Frame(self, bg=THEME["bg"])
        body.pack(fill="both", expand=True, padx=14, pady=10)

        # Painel esquerdo
        left = tk.Frame(body, bg=THEME["surface"], width=320)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        tk.Label(left, text="IMPRESSORAS DETECTADAS",
                 font=("Segoe UI", 9, "bold"), bg=THEME["surface"],
                 fg=THEME["text_muted"]).pack(anchor="w", padx=14, pady=(14, 6))

        sty = ttk.Style()
        sty.theme_use("default")
        sty.configure("Diag.Treeview",
                      background=THEME["surface2"], foreground=THEME["text"],
                      fieldbackground=THEME["surface2"], rowheight=34,
                      font=("Segoe UI", 10), borderwidth=0)
        sty.configure("Diag.Treeview.Heading",
                      background=THEME["border"], foreground=THEME["text_muted"],
                      font=("Segoe UI", 9, "bold"), relief="flat")
        sty.map("Diag.Treeview", background=[("selected", THEME["accent2"])])

        self.tree = ttk.Treeview(left, columns=("S", "Nome", "Tipo"),
                                 show="headings", style="Diag.Treeview",
                                 selectmode="browse")
        self.tree.heading("S", text="")
        self.tree.heading("Nome", text="Nome")
        self.tree.heading("Tipo", text="Tipo")
        self.tree.column("S", width=28, anchor="center")
        self.tree.column("Nome", width=195)
        self.tree.column("Tipo", width=65, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        tk.Button(left, text="🔍  Diagnosticar Selecionada",
                  command=self._diagnose_selected,
                  bg=THEME["accent2"], fg="white", font=("Segoe UI", 10, "bold"),
                  relief="flat", cursor="hand2", padx=12, pady=8
                  ).pack(fill="x", padx=8, pady=(0, 6))
        tk.Button(left, text="⚡  Diagnosticar Todas",
                  command=self._diagnose_all,
                  bg=THEME["surface2"], fg=THEME["text"], font=("Segoe UI", 10),
                  relief="flat", cursor="hand2", padx=12, pady=6
                  ).pack(fill="x", padx=8, pady=(0, 14))

        # Divisor
        tk.Frame(body, bg=THEME["border"], width=2).pack(side="left", fill="y", padx=8)

        # Painel direito
        right = tk.Frame(body, bg=THEME["bg"])
        right.pack(side="left", fill="both", expand=True)

        nb_sty = ttk.Style()
        nb_sty.configure("TNotebook", background=THEME["bg"], borderwidth=0)
        nb_sty.configure("TNotebook.Tab",
                         background=THEME["surface"], foreground=THEME["text_muted"],
                         font=("Segoe UI", 10), padding=[14, 8])
        nb_sty.map("TNotebook.Tab",
                   background=[("selected", THEME["accent"])],
                   foreground=[("selected", "white")])

        self.tabs = ttk.Notebook(right, style="TNotebook")
        self.tabs.pack(fill="both", expand=True)

        # Aba 1 — Diagnóstico
        t1 = tk.Frame(self.tabs, bg=THEME["bg"])
        self.tabs.add(t1, text="  Diagnóstico  ")
        self.diag_text = self._make_text(t1)
        for tag, color, bold in [
            ("header", THEME["accent"], True), ("ok", THEME["success"], False),
            ("warning", THEME["warning"], False), ("error", THEME["error"], False),
            ("muted", THEME["text_muted"], False), ("info", THEME["info"], False),
            ("bold", THEME["text"], True),
        ]:
            font = ("Consolas", 12 if tag == "header" else 11, "bold" if bold else "normal")
            self.diag_text.tag_configure(tag, foreground=color, font=font)

        # Aba 2 — Correções (scrollable)
        t2 = tk.Frame(self.tabs, bg=THEME["bg"])
        self.tabs.add(t2, text="  Correções  ")
        canvas = tk.Canvas(t2, bg=THEME["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(t2, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self.fix_frame = tk.Frame(canvas, bg=THEME["bg"])
        cw = canvas.create_window((0, 0), window=self.fix_frame, anchor="nw")
        self.fix_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
            lambda e: canvas.itemconfig(cw, width=e.width))
        self._show_fix_placeholder()

        # Aba 3 — Log
        t3 = tk.Frame(self.tabs, bg=THEME["bg"])
        self.tabs.add(t3, text="  Log  ")
        self.log_text = self._make_text(t3, fg=THEME["text_muted"], font_size=10)

        # Aba 4 — Guia de Rede
        t4 = tk.Frame(self.tabs, bg=THEME["bg"])
        self.tabs.add(t4, text="  Guia: Rede  ")
        self.subnet_text = self._make_text(t4)
        for tag, color, bold, bg in [
            ("title", THEME["accent"],  True,  None),
            ("step",  THEME["warning"], True,  None),
            ("cmd",   "#a8ff78",        False, "#0d1f0d"),
            ("ok",    THEME["success"], False, None),
            ("error", THEME["error"],   False, None),
            ("muted", THEME["text_muted"], False, None),
            ("info",  THEME["info"],    False, None),
        ]:
            kw = {"foreground": color,
                  "font": ("Consolas", 13 if tag == "title" else 11,
                            "bold" if bold else "normal")}
            if bg:
                kw["background"] = bg
            self.subnet_text.tag_configure(tag, **kw)
        self._write(self.subnet_text,
                    "\n  Execute o diagnóstico para ver o guia de rede.\n", "muted")

        # Status bar
        sb = tk.Frame(self, bg=THEME["surface"])
        sb.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="Pronto. Clique em 'Escanear' para começar.")
        tk.Label(sb, textvariable=self.status_var,
                 bg=THEME["surface"], fg=THEME["text_muted"],
                 font=("Segoe UI", 9), anchor="w", padx=14, pady=6
                 ).pack(side="left", fill="x", expand=True)
        tk.Label(sb, text="Desenvolvido por  Thomaz Arthur  •  v2.0",
                 bg=THEME["surface"], fg=THEME["border"],
                 font=("Segoe UI", 8), anchor="e", padx=14, pady=6
                 ).pack(side="right")

    # ── HELPERS ──────────────────────────────────

    def _make_text(self, parent, fg=None, font_size=11):
        t = scrolledtext.ScrolledText(
            parent, bg=THEME["surface"], fg=fg or THEME["text"],
            font=("Consolas", font_size), insertbackground=THEME["text"],
            relief="flat", borderwidth=0, padx=16, pady=12,
            wrap="word", state="disabled"
        )
        t.pack(fill="both", expand=True, padx=2, pady=2)
        return t

    def _write(self, widget, text: str, tag: str = ""):
        widget.config(state="normal")
        if tag:
            widget.insert("end", text, tag)
        else:
            widget.insert("end", text)
        widget.see("end")
        widget.config(state="disabled")

    def _clear(self, widget):
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.config(state="disabled")

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self._write(self.log_text, f"[{ts}] {msg}\n")

    def _set_status(self, msg: str):
        self.status_var.set(msg)
        self.update_idletasks()

    def _show_fix_placeholder(self):
        for w in self.fix_frame.winfo_children():
            w.destroy()
        tk.Label(self.fix_frame,
                 text="Selecione uma impressora e execute o diagnóstico\npara ver as correções disponíveis.",
                 bg=THEME["bg"], fg=THEME["text_muted"],
                 font=("Segoe UI", 12)).pack(expand=True, pady=40)

    # ── SCAN ─────────────────────────────────────

    def _scan_printers(self):
        self._set_status("Escaneando impressoras...")
        self._log("Iniciando escaneamento...")
        threading.Thread(target=self._scan_thread, daemon=True).start()

    def _scan_thread(self):
        printers = get_all_printers()
        self.after(0, self._populate_list, printers)

    def _populate_list(self, printers: list):
        self.printers_data = printers
        for row in self.tree.get_children():
            self.tree.delete(row)
        for p in printers:
            status  = p.get("PrinterStatus", 2)
            offline = p.get("WorkOffline", False)
            state   = p.get("PrinterState", 0) or 0
            port    = p.get("PortName", "")
            tipo    = "USB" if "USB" in port.upper() else (
                      "Rede" if re.search(r'\d{1,3}\.\d{1,3}', port) else "Local")

            if offline or status == 7 or (state & 64):
                icon = "🔴"
            elif status == 4 or (state & 2):
                icon = "🟠"
            elif state & 1:
                icon = "🟡"
            elif status == 2:
                icon = "🟢"
            else:
                icon = "🟡"

            self.tree.insert("", "end", values=(icon, p.get("Name", "?"), tipo))

        count = len(printers)
        self._set_status(f"{count} impressora(s) encontrada(s).")
        self._log(f"Escaneamento concluído: {count} impressora(s).")

    # ── SELECT ───────────────────────────────────

    def _on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx < len(self.printers_data):
            self.selected_printer = self.printers_data[idx]

    # ── DIAGNOSE ─────────────────────────────────

    def _diagnose_selected(self):
        if not self.selected_printer:
            messagebox.showwarning("Aviso", "Selecione uma impressora.")
            return
        threading.Thread(target=self._diagnose_thread,
                         args=([self.selected_printer],), daemon=True).start()

    def _diagnose_all(self):
        if not self.printers_data:
            messagebox.showwarning("Aviso", "Escaneie as impressoras primeiro.")
            return
        threading.Thread(target=self._diagnose_thread,
                         args=(self.printers_data,), daemon=True).start()

    def _diagnose_thread(self, printers: list):
        self.after(0, self._clear, self.diag_text)
        self.after(0, self._show_fix_placeholder)
        for p in printers:
            name = p.get("Name", "?")
            self.after(0, self._set_status, f"Diagnosticando: {name}...")
            self.after(0, self._log, f"Diagnosticando '{name}'...")
            diag = diagnose_printer(p)
            self.diagnoses[name] = diag
            self.after(0, self._render_diagnosis, diag)
        self.after(0, self._set_status, "Diagnóstico concluído.")
        self.after(0, self._render_fixes)

    def _render_diagnosis(self, diag: dict):
        sep = "─" * 64 + "\n"
        w = self.diag_text

        self._write(w, f"\n{sep}", "muted")
        self._write(w, f"  🖨  {diag['name']}\n", "header")
        self._write(w, sep, "muted")

        self._write(w, "  Porta   : ", "muted");  self._write(w, f"{diag['port']}\n")
        self._write(w, "  Driver  : ", "muted");  self._write(w, f"{diag['driver']}\n")
        self._write(w, "  Tipo    : ", "muted");  self._write(w, f"{diag['type']}\n")
        self._write(w, "  Status  : ", "muted");  self._write(w, f"{diag['status_text']}\n")

        if diag.get("offline"):
            self._write(w, "  ⚠ OFFLINE — definida como offline no Windows!\n", "warning")
        if diag.get("paused"):
            self._write(w, "  ⚠ PAUSADA — trabalhos não serão enviados!\n", "warning")

        # Driver
        dc = diag.get("driver_check")
        if dc:
            self._write(w, "\n  [ Driver ]\n", "bold")
            for det in dc.get("details", []):
                tag = "ok" if "✓" in det else ("info" if "✦" in det else ("warning" if "⚠" in det else "error"))
                self._write(w, f"    {det}\n", tag)

        # Porta
        pc = diag.get("port_check")
        if pc:
            self._write(w, "\n  [ Porta ]\n", "bold")
            for det in pc.get("details", []):
                tag = "ok" if "✓" in det else ("info" if "✦" in det else ("warning" if "⚠" in det else "error"))
                self._write(w, f"    {det}\n", tag)

        # Rede
        nc = diag.get("network_check")
        if nc:
            self._write(w, "\n  [ Rede ]\n", "bold")
            for det in nc.get("details", []):
                tag = "ok" if "✓" in det else ("error" if "✗" in det else "warning")
                self._write(w, f"    {det}\n", tag)

        # USB
        uc = diag.get("usb_check")
        if uc:
            self._write(w, "\n  [ USB ]\n", "bold")
            for det in uc.get("details", []):
                tag = "ok" if "registrado" in det.lower() else "error"
                self._write(w, f"    {det}\n", tag)

        # Resumo
        if diag["issues"]:
            self._write(w, f"\n  [ {len(diag['issues'])} Problema(s) Encontrado(s) ]\n", "bold")
            for issue in diag["issues"]:
                self._write(w, f"    ✗ {issue}\n", "error")
        else:
            self._write(w, "\n  ✓ Nenhum problema detectado.\n", "ok")

        self._write(w, "\n")

    # ── RENDER FIXES ─────────────────────────────

    def _render_fixes(self):
        for child in self.fix_frame.winfo_children():
            child.destroy()

        all_fixes = []
        for name, diag in self.diagnoses.items():
            for fix in diag.get("fixes", []):
                all_fixes.append((name, fix, diag))

        if not all_fixes:
            tk.Label(self.fix_frame,
                     text="✓ Nenhuma correção necessária!",
                     bg=THEME["bg"], fg=THEME["success"],
                     font=("Segoe UI", 13, "bold")).pack(pady=30)
            return

        FIX_META = {
            "set_online":       ("⚡ Definir como Online",          THEME["success"], False),
            "resume_printer":   ("▶  Retomar Impressão Pausada",    THEME["success"], False),
            "clear_queue":      ("🗑  Limpar Fila de Impressão",     THEME["warning"], False),
            "restart_spooler":  ("↺  Reiniciar Spooler",            THEME["accent"],  False),
            "fix_driver":       ("🔄  Corrigir Driver",              THEME["accent"],  False),
            "fix_port":         ("🔌  Corrigir Porta",              "#a78bfa",         False),
            "subnet_guide":     ("🌐  Ver Guia: Sub-rede",           THEME["info"],    True),
            "ping_fail_guide":  ("📡  Ver Guia: Sem Resposta",       THEME["warning"], True),
        }

        auto_fixes = [(pn, fk, d) for pn, fk, d in all_fixes
                      if not FIX_META.get(fk, ("", "", True))[2]]

        tk.Label(self.fix_frame,
                 text=f"{len(all_fixes)} correção(ões) disponível(is)  •  {len(auto_fixes)} automática(s)",
                 bg=THEME["bg"], fg=THEME["text"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(12, 8))

        for printer_name, fix_key, diag in all_fixes:
            label, color, is_guide = FIX_META.get(fix_key, (fix_key, THEME["text_muted"], False))
            card = tk.Frame(self.fix_frame, bg=THEME["surface"], pady=10)
            card.pack(fill="x", padx=8, pady=3)

            tk.Label(card, text=f"  {label}",
                     bg=THEME["surface"], fg=color,
                     font=("Segoe UI", 11, "bold")).pack(side="left", padx=8)
            tk.Label(card, text=f"→  {printer_name}",
                     bg=THEME["surface"], fg=THEME["text_muted"],
                     font=("Segoe UI", 10)).pack(side="left")

            btn_text = "Ver Guia" if is_guide else "Aplicar"
            tk.Button(card, text=btn_text,
                      command=lambda pn=printer_name, fk=fix_key, d=diag: self._apply_fix(pn, fk, d),
                      bg=color, fg="#0f1117",
                      font=("Segoe UI", 9, "bold"),
                      relief="flat", cursor="hand2", padx=12, pady=4
                      ).pack(side="right", padx=12)

        if auto_fixes:
            tk.Button(self.fix_frame,
                      text="⚡  Aplicar Todas as Correções Automáticas",
                      command=lambda: self._apply_all_fixes(auto_fixes),
                      bg=THEME["accent"], fg="white",
                      font=("Segoe UI", 11, "bold"),
                      relief="flat", cursor="hand2",
                      padx=16, pady=10).pack(fill="x", padx=8, pady=(14, 4))

    # ── FIXES ────────────────────────────────────

    def _apply_fix(self, printer_name: str, fix_key: str, diag: dict):
        self._log(f"Aplicando '{fix_key}' em '{printer_name}'...")
        threading.Thread(target=self._fix_thread,
                         args=(printer_name, fix_key, diag), daemon=True).start()

    def _apply_all_fixes(self, fixes: list):
        for pn, fk, d in fixes:
            threading.Thread(target=self._fix_thread,
                             args=(pn, fk, d), daemon=True).start()

    def _fix_thread(self, printer_name: str, fix_key: str, diag: dict):
        if fix_key == "subnet_guide":
            nc = diag.get("network_check", {})
            self.after(0, self._render_subnet_guide,
                       printer_name, nc.get("printer_ip", "?"), nc.get("server_ip", "?"))
            return
        if fix_key == "ping_fail_guide":
            self.after(0, self._render_ping_fail_guide, printer_name, diag)
            return

        result = ""
        if fix_key == "set_online":
            result = fix_set_online(printer_name)
        elif fix_key == "resume_printer":
            result = fix_resume_printer(printer_name)
        elif fix_key == "clear_queue":
            result = fix_clear_queue(printer_name)
        elif fix_key == "restart_spooler":
            result = fix_restart_spooler()
        elif fix_key == "fix_driver":
            dc = diag.get("driver_check", {})
            suggested = dc.get("suggested_driver") or diag.get("driver", "")
            result = fix_driver(printer_name, suggested, diag.get("port", "USB001"))
        elif fix_key == "fix_port":
            pc = diag.get("port_check", {})
            new_port = pc.get("suggested_port", "")
            result = (fix_port(printer_name, new_port) if new_port
                      else f"Nenhuma porta sugerida para '{printer_name}'.")

        self.after(0, self._log, f"✓ {result}")
        self.after(0, self._set_status, result)

    def _restart_spooler(self):
        if messagebox.askyesno("Confirmar",
                               "Reiniciar o serviço Spooler agora?\n\n"
                               "Impressões em andamento serão interrompidas."):
            threading.Thread(target=self._spooler_thread, daemon=True).start()

    def _spooler_thread(self):
        self.after(0, self._set_status, "Reiniciando Spooler...")
        self.after(0, self._log, "Reiniciando Spooler...")
        result = fix_restart_spooler()
        self.after(0, self._log, f"✓ {result}")
        self.after(0, self._set_status, result)
        self.after(0, messagebox.showinfo, "Spooler", result)

    # ── GUIAS DE REDE ─────────────────────────────

    def _render_subnet_guide(self, printer_name, printer_ip, server_ip):
        try:
            p  = printer_ip.split(".")
            s  = server_ip.split(".")
            printer_net  = ".".join(p[:3]) + ".0/24"
            server_net   = ".".join(s[:3]) + ".0/24"
            suggested_ip = ".".join(s[:3]) + "." + p[3]
            gateway      = ".".join(s[:3]) + ".1"
        except:
            printer_net = server_net = suggested_ip = gateway = "?"

        w = self.subnet_text
        self._clear(w)
        self._write(w, "═" * 64 + "\n", "muted")
        self._write(w, "  🌐  GUIA: Sub-redes Diferentes\n", "title")
        self._write(w, f"  Impressora: {printer_name}\n", "muted")
        self._write(w, "═" * 64 + "\n\n", "muted")
        self._write(w, "  SITUAÇÃO\n", "step")
        self._write(w, f"  ✗ IP impressora : {printer_ip}  (rede {printer_net})\n", "error")
        self._write(w, f"  ✗ IP servidor   : {server_ip}  (rede {server_net})\n\n", "error")

        self._write(w, "─" * 64 + "\n", "muted")
        self._write(w, "  OPÇÃO 1 — Alterar IP da Impressora (Recomendado)\n", "step")
        self._write(w, "─" * 64 + "\n", "muted")
        self._write(w, """
  Acesse o menu de rede da impressora:
  • HP      → Menu > Configuração > Rede > TCP/IP
  • Epson   → Wi-Fi Setup / Rede > Config. IP
  • Brother → Menu > Rede > LAN com fio > Config. TCP/IP
  • Canon   → Menu > Preferências de rede

  Altere:\n""", "muted")
        self._write(w, f"    {printer_ip}  →  {suggested_ip}\n", "info")
        self._write(w, f"  Máscara: 255.255.255.0  |  Gateway: {gateway}\n\n", "muted")

        self._write(w, "─" * 64 + "\n", "muted")
        self._write(w, "  OPÇÃO 2 — Alterar Porta no Servidor Windows\n", "step")
        self._write(w, "─" * 64 + "\n", "muted")
        self._write(w, "\n  PowerShell (Administrador):\n", "muted")
        self._write(w,
            f'  Add-PrinterPort -Name "IP_{printer_ip}" -PrinterHostAddress "{printer_ip}"\n'
            f'  Set-Printer -Name "{printer_name}" -PortName "IP_{printer_ip}"\n', "cmd")
        self._write(w, "\n")

        self._write(w, "─" * 64 + "\n", "muted")
        self._write(w, "  VERIFICAÇÃO FINAL\n", "step")
        self._write(w, f'\n  ping {printer_ip}\n', "cmd")
        self._write(w, "  Se responder → imprima uma página de teste.\n\n", "muted")
        self.tabs.select(3)
        self._log(f"Guia de sub-rede exibido para '{printer_name}'.")

    def _render_ping_fail_guide(self, printer_name, diag):
        nc = diag.get("network_check", {})
        printer_ip = nc.get("printer_ip", "?")
        try:
            net_prefix = ".".join(printer_ip.split(".")[:3])
        except:
            net_prefix = "192.168.1"

        w = self.subnet_text
        self._clear(w)
        self._write(w, "═" * 64 + "\n", "muted")
        self._write(w, "  📡  GUIA: Impressora Sem Resposta (Ping Falhou)\n", "title")
        self._write(w, f"  Impressora: {printer_name}  |  IP: {printer_ip}\n", "muted")
        self._write(w, "═" * 64 + "\n\n", "muted")

        self._write(w, "  POSSÍVEIS CAUSAS\n", "step")
        self._write(w, """
  1. Impressora desligada ou em espera profunda
  2. Cabo de rede desconectado
  3. IP mudou (DHCP atribuiu outro endereço)
  4. Defeito de hardware
  5. Firewall bloqueando ICMP\n\n""", "muted")

        self._write(w, "  PASSO 1 — Verificação física\n", "step")
        self._write(w, "  • Ligada?  LED de rede aceso?  Cabo conectado?\n\n", "muted")

        self._write(w, "  PASSO 2 — Descobrir IP atual\n", "step")
        self._write(w, """
  Imprima a página de config. de rede:
  • HP      → Segure botão ℹ por 3s
  • Epson   → Menu > Configuração > Status da Rede
  • Brother → Menu > Informações > Impr. config. rede
  • Canon   → Menu > Config. dispositivo > Folha de status\n\n""", "muted")

        self._write(w, "  PASSO 3 — Varredura de rede\n", "step")
        self._write(w,
            f'  1..254 | % {{ $ip = "{net_prefix}.$_";\n'
            f'    if (Test-Connection $ip -Count 1 -Quiet) {{ $ip }} }}\n', "cmd")

        self._write(w, "\n  PASSO 4 — Atualizar porta após encontrar novo IP\n", "step")
        self._write(w,
            f'  Set-PrinterPort -Name "IP_{printer_ip}" -PrinterHostAddress "NOVO_IP"\n', "cmd")
        self._write(w, "\n", "muted")
        self.tabs.select(3)
        self._log(f"Guia de ping exibido para '{printer_name}'.")


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    app = PrinterDiagApp()
    app.mainloop()
