"""
=============================================================
  PrinterDiag - Ferramenta de Diagnóstico de Impressoras
  Desenvolvida para Analistas de Suporte
=============================================================
  Desenvolvido por: Thomaz Arthur
  Versão: 1.0
=============================================================
Requisitos: Python 3.8+, pywin32, requests
Instale com: pip install pywin32 requests
=============================================================
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import threading
import socket
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
}


# ─────────────────────────────────────────────
#  FUNÇÕES DE DIAGNÓSTICO (LÓGICA PRINCIPAL)
# ─────────────────────────────────────────────

def run_ps(command: str) -> str:
    """Executa um comando PowerShell e retorna o output."""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True, text=True, timeout=30
        )
        return result.stdout.strip()
    except Exception as e:
        return f"ERRO: {e}"


def get_all_printers() -> list[dict]:
    """Retorna lista de impressoras com status via WMI."""
    ps = """
    Get-WmiObject -Class Win32_Printer | Select-Object Name, PortName, DriverName,
        WorkOffline, PrinterStatus, DetectedErrorState, ExtendedDetectedErrorState |
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


def get_printer_jobs(printer_name: str) -> list[dict]:
    """Retorna fila de impressão de uma impressora."""
    ps = f"""
    Get-PrintJob -PrinterName '{printer_name}' |
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


def get_server_ip() -> str:
    """Obtém o IP do servidor de impressão (spooler)."""
    ps = "Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.IPAddress} | Select-Object -ExpandProperty IPAddress -First 1"
    return run_ps(ps).strip()


def check_usb_driver(printer: dict) -> dict:
    """Verifica se há problemas de driver para impressoras USB."""
    result = {"has_issue": False, "details": [], "fix_available": False}
    port = printer.get("PortName", "")
    driver = printer.get("DriverName", "")

    if "USB" in port.upper():
        # Verifica se o driver está instalado
        ps = f"Get-PrinterDriver -Name '{driver}' -ErrorAction SilentlyContinue | Select-Object Name, PrinterEnvironment | ConvertTo-Json"
        output = run_ps(ps)

        if not output or "ERRO" in output:
            result["has_issue"] = True
            result["details"].append(f"Driver '{driver}' não encontrado ou com problema.")
            result["fix_available"] = True
        else:
            result["details"].append(f"Driver '{driver}' instalado corretamente.")

        # Verifica se a porta USB está ativa
        ps_port = f"Get-PrinterPort -Name '{port}' -ErrorAction SilentlyContinue"
        port_out = run_ps(ps_port)
        if not port_out or "ERRO" in port_out:
            result["has_issue"] = True
            result["details"].append(f"Porta USB '{port}' não responde.")
            result["fix_available"] = True

    return result


def check_network_range(printer: dict) -> dict:
    """Verifica se a impressora de rede está na mesma faixa do servidor."""
    result = {"has_issue": False, "details": [], "printer_ip": None, "server_ip": None, "fix_available": False}

    port = printer.get("PortName", "")

    # Tenta extrair IP da porta (ex: IP_192.168.1.100 ou diretamente)
    ip_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', port)
    if not ip_match:
        # Tenta buscar o IP via WMI na porta TCP/IP
        ps = f"Get-WmiObject Win32_TCPIPPrinterPort | Where-Object {{$_.Name -eq '{port}'}} | Select-Object HostAddress | ConvertTo-Json"
        output = run_ps(ps)
        try:
            data = json.loads(output)
            ip_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', str(data))
        except:
            pass

    if not ip_match:
        result["details"].append("Porta não parece ser de rede TCP/IP.")
        return result

    printer_ip = ip_match.group(1)
    result["printer_ip"] = printer_ip

    # Obtém IP do servidor
    server_ip = get_server_ip()
    result["server_ip"] = server_ip

    try:
        p_net = ipaddress.IPv4Network(printer_ip + "/24", strict=False)
        s_net = ipaddress.IPv4Network(server_ip + "/24", strict=False)

        if p_net != s_net:
            result["has_issue"] = True
            result["subnet_mismatch"] = True
            result["details"].append(
                f"⚠ Impressora ({printer_ip}) está em sub-rede diferente do servidor ({server_ip})."
            )
            result["fix_available"] = True  # Guia de correção disponível
        else:
            result["subnet_mismatch"] = False
            result["details"].append(f"✓ Mesma sub-rede: impressora {printer_ip} / servidor {server_ip}.")

        # Testa ping
        ping = subprocess.run(["ping", "-n", "1", "-w", "1000", printer_ip],
                               capture_output=True, text=True)
        if "TTL=" in ping.stdout:
            result["details"].append(f"✓ Impressora responde ao ping ({printer_ip}).")
            result["ping_ok"] = True
        else:
            result["has_issue"] = True
            result["ping_ok"] = False
            result["details"].append(f"✗ Impressora não responde ao ping ({printer_ip}) — pode estar offline.")

    except Exception as e:
        result["details"].append(f"Erro ao verificar rede: {e}")

    return result


def get_error_description(printer: dict) -> str:
    """Traduz os códigos de erro do Win32_Printer para texto."""
    status_map = {
        1: "Outro",
        2: "✓ Normal",
        3: "⚠ Inativa",
        4: "✗ Erro",
        5: "✗ Desconhecida",
        6: "✗ Impressão não disponível",
        7: "✗ Offline",
        8: "✗ I/O ativo",
        9: "⚠ Ocupada",
        10: "⚠ Imprimindo",
        11: "⚠ Aquecendo",
        12: "⚠ Sem seleção",
        13: "⚠ Fora do papel"
    }
    error_map = {
        0: "Nenhum erro detectado",
        2: "✗ Papel atolado",
        4: "✗ Sem papel",
        8: "✗ Sem toner/tinta",
        16: "✗ Tinta/Toner quase acabando",
        32: "✗ Erro de saída de papel",
        64: "⚠ Modo offline",
        128: "✗ Requer intervenção do serviço",
        256: "✗ Erro de entrada/saída",
        512: "✗ Erro de envelope",
        1024: "✗ Erro de papel personalizado",
        2048: "✗ Erro de bandeja",
        4096: "⚠ Tampa aberta",
        8192: "✗ Referência de ponto"
    }
    status_code = printer.get("PrinterStatus", 2)
    error_code = printer.get("DetectedErrorState", 0)
    status_text = status_map.get(status_code, f"Código {status_code}")
    error_text = error_map.get(error_code, f"Código de erro {error_code}")
    return f"Status: {status_text} | Erro: {error_text}"


def diagnose_printer(printer: dict) -> dict:
    """Diagnóstico completo de uma impressora."""
    port = printer.get("PortName", "")
    is_usb = "USB" in port.upper()
    is_network = re.search(r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', port) or "TCP" in port.upper()

    diagnosis = {
        "name": printer.get("Name", "Desconhecida"),
        "port": port,
        "driver": printer.get("DriverName", ""),
        "offline": printer.get("WorkOffline", False),
        "status_text": get_error_description(printer),
        "type": "USB" if is_usb else ("Rede" if is_network else "Local"),
        "issues": [],
        "fixes": [],
        "usb_check": None,
        "network_check": None,
    }

    # Verifica offline
    if printer.get("WorkOffline"):
        diagnosis["issues"].append("Impressora definida como OFFLINE no Windows.")
        diagnosis["fixes"].append("set_online")

    # Status de erro
    status = printer.get("PrinterStatus", 2)
    if status not in (2, 10):  # Normal ou imprimindo
        diagnosis["issues"].append(f"Status de erro reportado pelo Windows.")

    if is_usb:
        usb = check_usb_driver(printer)
        diagnosis["usb_check"] = usb
        if usb["has_issue"]:
            diagnosis["issues"].extend(usb["details"])
            if usb["fix_available"]:
                diagnosis["fixes"].append("reinstall_driver")
    elif is_network:
        net = check_network_range(printer)
        diagnosis["network_check"] = net
        if net["has_issue"]:
            diagnosis["issues"].extend(net["details"])
            if net.get("subnet_mismatch"):
                diagnosis["fixes"].append("subnet_guide")
            if not net.get("ping_ok", True) and not net.get("subnet_mismatch"):
                diagnosis["fixes"].append("ping_fail_guide")

    # Verifica fila travada
    jobs = get_printer_jobs(printer.get("Name", ""))
    if jobs:
        diagnosis["issues"].append(f"{len(jobs)} trabalho(s) na fila de impressão.")
        diagnosis["fixes"].append("clear_queue")

    return diagnosis


# ─────────────────────────────────────────────
#  FUNÇÕES DE CORREÇÃO AUTOMÁTICA
# ─────────────────────────────────────────────

def fix_set_online(printer_name: str) -> str:
    ps = f"Set-Printer -Name '{printer_name}' -WorkOffline $false"
    run_ps(ps)
    return f"Impressora '{printer_name}' definida como ONLINE."


def fix_clear_queue(printer_name: str) -> str:
    ps = f"""
    $jobs = Get-PrintJob -PrinterName '{printer_name}' -ErrorAction SilentlyContinue
    if ($jobs) {{ $jobs | Remove-PrintJob }}
    """
    run_ps(ps)
    return f"Fila de impressão de '{printer_name}' limpa."


def fix_restart_spooler() -> str:
    run_ps("Stop-Service -Name Spooler -Force")
    time.sleep(2)
    run_ps("Start-Service -Name Spooler")
    return "Serviço Spooler reiniciado com sucesso."


def fix_reinstall_driver(printer_name: str, driver_name: str) -> str:
    # Remove e reinstala o driver
    ps = f"""
    Remove-Printer -Name '{printer_name}' -ErrorAction SilentlyContinue
    Add-Printer -Name '{printer_name}' -DriverName '{driver_name}' -PortName 'USB001' -ErrorAction SilentlyContinue
    """
    run_ps(ps)
    return f"Driver '{driver_name}' reinstalado para '{printer_name}'."


# ─────────────────────────────────────────────
#  INTERFACE GRÁFICA
# ─────────────────────────────────────────────

class PrinterDiagApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PrinterDiag — Diagnóstico de Impressoras")
        self.geometry("1100x750")
        self.configure(bg=THEME["bg"])
        self.resizable(True, True)

        self.printers_data = []
        self.diagnoses = {}
        self.selected_printer = None

        self._build_ui()

    # ── BUILD UI ──────────────────────────────

    def _build_ui(self):
        # ── Header ──
        header = tk.Frame(self, bg=THEME["surface"], height=60)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="🖨  PrinterDiag", font=("Segoe UI", 16, "bold"),
                 bg=THEME["surface"], fg=THEME["accent"]).pack(side="left", padx=20, pady=12)
        tk.Label(header, text="Diagnóstico Automático de Impressoras",
                 font=("Segoe UI", 10), bg=THEME["surface"], fg=THEME["text_muted"]).pack(side="left", pady=12)

        btn_scan = tk.Button(header, text="⟳  Escanear Impressoras",
                             command=self._scan_printers,
                             bg=THEME["accent"], fg="white",
                             font=("Segoe UI", 10, "bold"),
                             relief="flat", cursor="hand2", padx=16, pady=6)
        btn_scan.pack(side="right", padx=20, pady=10)

        btn_spooler = tk.Button(header, text="↺  Reiniciar Spooler",
                                command=self._restart_spooler,
                                bg=THEME["surface2"], fg=THEME["text"],
                                font=("Segoe UI", 10),
                                relief="flat", cursor="hand2", padx=14, pady=6)
        btn_spooler.pack(side="right", padx=4, pady=10)

        # ── Body ──
        body = tk.Frame(self, bg=THEME["bg"])
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # Painel esquerdo — lista de impressoras
        left = tk.Frame(body, bg=THEME["surface"], width=340)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        tk.Label(left, text="IMPRESSORAS DETECTADAS",
                 font=("Segoe UI", 9, "bold"), bg=THEME["surface"],
                 fg=THEME["text_muted"]).pack(anchor="w", padx=14, pady=(14, 6))

        # Tabela de impressoras
        cols = ("Status", "Nome", "Tipo")
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Diag.Treeview",
                        background=THEME["surface2"],
                        foreground=THEME["text"],
                        fieldbackground=THEME["surface2"],
                        rowheight=34,
                        font=("Segoe UI", 10),
                        borderwidth=0)
        style.configure("Diag.Treeview.Heading",
                        background=THEME["border"],
                        foreground=THEME["text_muted"],
                        font=("Segoe UI", 9, "bold"),
                        relief="flat")
        style.map("Diag.Treeview", background=[("selected", THEME["accent2"])])

        self.tree = ttk.Treeview(left, columns=cols, show="headings",
                                 style="Diag.Treeview", selectmode="browse")
        self.tree.heading("Status", text="")
        self.tree.heading("Nome", text="Nome")
        self.tree.heading("Tipo", text="Tipo")
        self.tree.column("Status", width=28, anchor="center")
        self.tree.column("Nome", width=200)
        self.tree.column("Tipo", width=65, anchor="center")

        self.tree.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Botão de diagnóstico
        self.btn_diagnose = tk.Button(left, text="🔍  Diagnosticar Selecionada",
                                      command=self._diagnose_selected,
                                      bg=THEME["accent2"], fg="white",
                                      font=("Segoe UI", 10, "bold"),
                                      relief="flat", cursor="hand2",
                                      padx=12, pady=8)
        self.btn_diagnose.pack(fill="x", padx=8, pady=(0, 8))

        btn_all = tk.Button(left, text="⚡  Diagnosticar Todas",
                            command=self._diagnose_all,
                            bg=THEME["surface2"], fg=THEME["text"],
                            font=("Segoe UI", 10),
                            relief="flat", cursor="hand2",
                            padx=12, pady=6)
        btn_all.pack(fill="x", padx=8, pady=(0, 14))

        # Divisor
        tk.Frame(body, bg=THEME["border"], width=2).pack(side="left", fill="y", padx=8)

        # Painel direito — detalhes e correções
        right = tk.Frame(body, bg=THEME["bg"])
        right.pack(side="left", fill="both", expand=True)

        # Tabs
        tab_style = ttk.Style()
        tab_style.configure("TNotebook", background=THEME["bg"], borderwidth=0)
        tab_style.configure("TNotebook.Tab", background=THEME["surface"],
                            foreground=THEME["text_muted"],
                            font=("Segoe UI", 10),
                            padding=[14, 8])
        tab_style.map("TNotebook.Tab",
                      background=[("selected", THEME["accent"])],
                      foreground=[("selected", "white")])

        self.tabs = ttk.Notebook(right, style="TNotebook")
        self.tabs.pack(fill="both", expand=True)

        # Aba 1 — Diagnóstico
        tab_diag = tk.Frame(self.tabs, bg=THEME["bg"])
        self.tabs.add(tab_diag, text="  Diagnóstico  ")

        self.diag_text = scrolledtext.ScrolledText(
            tab_diag,
            bg=THEME["surface"], fg=THEME["text"],
            font=("Consolas", 11),
            insertbackground=THEME["text"],
            relief="flat", borderwidth=0,
            padx=16, pady=12,
            wrap="word",
            state="disabled"
        )
        self.diag_text.pack(fill="both", expand=True, padx=2, pady=2)

        # Tags de cor para o log
        self.diag_text.tag_configure("header",  foreground=THEME["accent"],  font=("Consolas", 12, "bold"))
        self.diag_text.tag_configure("ok",      foreground=THEME["success"])
        self.diag_text.tag_configure("warning", foreground=THEME["warning"])
        self.diag_text.tag_configure("error",   foreground=THEME["error"])
        self.diag_text.tag_configure("muted",   foreground=THEME["text_muted"])
        self.diag_text.tag_configure("bold",    font=("Consolas", 11, "bold"))

        # Aba 2 — Correções
        tab_fix = tk.Frame(self.tabs, bg=THEME["bg"])
        self.tabs.add(tab_fix, text="  Correções Disponíveis  ")

        self.fix_frame = tk.Frame(tab_fix, bg=THEME["bg"])
        self.fix_frame.pack(fill="both", expand=True, padx=16, pady=16)

        self._show_placeholder()

        # Aba 3 — Log
        tab_log = tk.Frame(self.tabs, bg=THEME["bg"])
        self.tabs.add(tab_log, text="  Log de Ações  ")

        self.log_text = scrolledtext.ScrolledText(
            tab_log,
            bg=THEME["surface"], fg=THEME["text_muted"],
            font=("Consolas", 10),
            relief="flat", borderwidth=0,
            padx=14, pady=10,
            state="disabled"
        )
        self.log_text.pack(fill="both", expand=True, padx=2, pady=2)

        # Aba 4 — Guia Sub-rede
        tab_subnet = tk.Frame(self.tabs, bg=THEME["bg"])
        self.tabs.add(tab_subnet, text="  Guia: Sub-rede  ")

        self.subnet_text = scrolledtext.ScrolledText(
            tab_subnet,
            bg=THEME["surface"], fg=THEME["text"],
            font=("Consolas", 11),
            insertbackground=THEME["text"],
            relief="flat", borderwidth=0,
            padx=16, pady=12,
            wrap="word",
            state="disabled"
        )
        self.subnet_text.pack(fill="both", expand=True, padx=2, pady=2)
        self.subnet_text.tag_configure("title",   foreground=THEME["accent"],  font=("Consolas", 13, "bold"))
        self.subnet_text.tag_configure("step",    foreground=THEME["warning"], font=("Consolas", 11, "bold"))
        self.subnet_text.tag_configure("cmd",     foreground="#a8ff78",        font=("Consolas", 11), background="#0d1f0d")
        self.subnet_text.tag_configure("ok",      foreground=THEME["success"])
        self.subnet_text.tag_configure("error",   foreground=THEME["error"])
        self.subnet_text.tag_configure("muted",   foreground=THEME["text_muted"])
        self.subnet_text.tag_configure("info",    foreground="#7dd3fc")
        self._write_subnet_placeholder()

        # Status bar
        self.status_var = tk.StringVar(value="Pronto. Clique em 'Escanear Impressoras' para começar.")
        status_bar = tk.Frame(self, bg=THEME["surface"])
        status_bar.pack(fill="x", side="bottom")

        tk.Label(status_bar, textvariable=self.status_var,
                 bg=THEME["surface"], fg=THEME["text_muted"],
                 font=("Segoe UI", 9), anchor="w", padx=14, pady=6).pack(side="left", fill="x", expand=True)

        tk.Label(status_bar, text="Desenvolvido por  Thomaz Arthur  •  v1.0",
                 bg=THEME["surface"], fg=THEME["border"],
                 font=("Segoe UI", 8), anchor="e", padx=14, pady=6).pack(side="right")

    def _write_subnet_placeholder(self):
        self.subnet_text.config(state="normal")
        self.subnet_text.delete("1.0", "end")
        self.subnet_text.insert("end", "\n  Execute o diagnóstico para ver o guia de correção de sub-rede.\n", "muted")
        self.subnet_text.config(state="disabled")

    def _write_subnet(self, text: str, tag: str = ""):
        self.subnet_text.config(state="normal")
        if tag:
            self.subnet_text.insert("end", text, tag)
        else:
            self.subnet_text.insert("end", text)
        self.subnet_text.config(state="disabled")

    def _render_subnet_guide(self, printer_name: str, printer_ip: str, server_ip: str):
        """Gera o guia passo a passo para corrigir problema de sub-rede."""
        # Calcula as redes
        try:
            p_parts = printer_ip.split(".")
            s_parts = server_ip.split(".")
            printer_net = ".".join(p_parts[:3]) + ".0/24"
            server_net  = ".".join(s_parts[:3]) + ".0/24"
            # Sugere novo IP para impressora na mesma rede do servidor
            suggested_ip = ".".join(s_parts[:3]) + "." + p_parts[3]
        except:
            printer_net = "?"
            server_net  = "?"
            suggested_ip = "?"

        self.subnet_text.config(state="normal")
        self.subnet_text.delete("1.0", "end")

        self._write_subnet("═" * 62 + "\n", "muted")
        self._write_subnet(f"  🌐  GUIA: Correção de Sub-rede Diferente\n", "title")
        self._write_subnet(f"  Impressora: {printer_name}\n", "muted")
        self._write_subnet("═" * 62 + "\n\n", "muted")

        self._write_subnet("  SITUAÇÃO DETECTADA\n", "step")
        self._write_subnet(f"  ✗ IP da impressora : {printer_ip}  (rede {printer_net})\n", "error")
        self._write_subnet(f"  ✗ IP do servidor   : {server_ip}  (rede {server_net})\n", "error")
        self._write_subnet("  → O servidor de impressão não consegue se comunicar\n    com a impressora porque estão em redes diferentes.\n\n", "muted")

        self._write_subnet("─" * 62 + "\n", "muted")
        self._write_subnet("  OPÇÃO 1 — Alterar o IP da Impressora (Recomendado)\n", "step")
        self._write_subnet("─" * 62 + "\n", "muted")
        self._write_subnet("""
  Acesse a configuração de rede da impressora:
  • Impressoras HP  → Menu > Configuração > Rede > TCP/IP
  • Impressoras Epson → Wi-Fi Setup / Rede > Config. IP
  • Impressoras Brother → Menu > Rede > LAN com fio > Config. TCP/IP
  • Impressoras Canon → Menu > Configuração > Preferências de rede

  Altere o IP de:
""", "muted")
        self._write_subnet(f"    {printer_ip}   →   {suggested_ip}\n", "info")
        self._write_subnet(f"""
  Mantenha:
    • Máscara de sub-rede : 255.255.255.0
    • Gateway padrão      : {".".join(s_parts[:3])}.1  (verifique com o cliente)
    • DNS                 : mesmo do servidor\n\n""", "muted")

        self._write_subnet("─" * 62 + "\n", "muted")
        self._write_subnet("  OPÇÃO 2 — Alterar a Porta no Servidor Windows\n", "step")
        self._write_subnet("─" * 62 + "\n", "muted")
        self._write_subnet("""
  Se não conseguir alterar o IP da impressora, mude
  a porta de impressão no Windows para o IP atual dela:

  Execute no PowerShell como Administrador:\n""", "muted")

        ps_cmd = (
            f'  Add-PrinterPort -Name "IP_{printer_ip}" -PrinterHostAddress "{printer_ip}"\n'
            f'  Set-Printer -Name "{printer_name}" -PortName "IP_{printer_ip}"\n'
        )
        self._write_subnet(ps_cmd, "cmd")
        self._write_subnet("""
  Ou manualmente:
  1. Painel de Controle → Dispositivos e Impressoras
  2. Clique direito na impressora → Propriedades da impressora
  3. Aba "Portas" → Adicionar Porta → Porta TCP/IP padrão
  4. Digite o IP atual da impressora\n\n""", "muted")

        self._write_subnet("─" * 62 + "\n", "muted")
        self._write_subnet("  OPÇÃO 3 — Verificar Roteamento de Rede\n", "step")
        self._write_subnet("─" * 62 + "\n", "muted")
        self._write_subnet(f"""
  Se a impressora PRECISA ficar no IP {printer_ip},
  é necessário configurar roteamento entre as redes.

  Verifique com o responsável de TI/rede se existe
  rota entre {server_net} e {printer_net}.

  Teste de roteamento (rode no servidor):\n""", "muted")
        self._write_subnet(f'  tracert {printer_ip}\n', "cmd")
        self._write_subnet("""
  Se o tracert não chegar à impressora, a rota não existe
  e será necessário configurar no switch/roteador.\n\n""", "muted")

        self._write_subnet("─" * 62 + "\n", "muted")
        self._write_subnet("  VERIFICAÇÃO FINAL\n", "step")
        self._write_subnet("─" * 62 + "\n", "muted")
        self._write_subnet(f"""
  Após qualquer uma das correções, teste:

  1. Ping da impressora:\n""", "muted")
        self._write_subnet(f'     ping {printer_ip}\n', "cmd")
        self._write_subnet("""
  2. Tente imprimir uma página de teste
  3. Clique em "Escanear Impressoras" novamente
     para confirmar que o problema foi resolvido ✓\n\n""", "muted")

        self.subnet_text.config(state="disabled")
        # Muda para a aba do guia automaticamente
        self.tabs.select(3)

    # ── PLACEHOLDERS / HELPERS ─────────────────

    def _show_placeholder(self):
        for w in self.fix_frame.winfo_children():
            w.destroy()
        tk.Label(self.fix_frame,
                 text="Selecione uma impressora e execute o diagnóstico\npara ver as correções disponíveis.",
                 bg=THEME["bg"], fg=THEME["text_muted"],
                 font=("Segoe UI", 12)).pack(expand=True)

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{ts}] {msg}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _write_diag(self, text: str, tag: str = ""):
        self.diag_text.config(state="normal")
        if tag:
            self.diag_text.insert("end", text, tag)
        else:
            self.diag_text.insert("end", text)
        self.diag_text.see("end")
        self.diag_text.config(state="disabled")

    def _clear_diag(self):
        self.diag_text.config(state="normal")
        self.diag_text.delete("1.0", "end")
        self.diag_text.config(state="disabled")

    def _set_status(self, msg: str):
        self.status_var.set(msg)
        self.update_idletasks()

    # ── SCAN ──────────────────────────────────

    def _scan_printers(self):
        self._set_status("Escaneando impressoras...")
        self._log("Iniciando escaneamento de impressoras...")
        threading.Thread(target=self._scan_thread, daemon=True).start()

    def _scan_thread(self):
        printers = get_all_printers()
        self.after(0, self._populate_list, printers)

    def _populate_list(self, printers: list):
        self.printers_data = printers
        for row in self.tree.get_children():
            self.tree.delete(row)

        for p in printers:
            status = p.get("PrinterStatus", 2)
            offline = p.get("WorkOffline", False)
            port = p.get("PortName", "")
            is_usb = "USB" in port.upper()
            is_net = bool(re.search(r'\d{1,3}\.\d{1,3}', port))
            tipo = "USB" if is_usb else ("Rede" if is_net else "Local")

            if offline or status == 7:
                icon = "🔴"
            elif status == 4:
                icon = "🟠"
            elif status == 2:
                icon = "🟢"
            else:
                icon = "🟡"

            self.tree.insert("", "end", values=(icon, p.get("Name", "?"), tipo))

        count = len(printers)
        self._set_status(f"{count} impressora(s) encontrada(s).")
        self._log(f"Escaneamento concluído: {count} impressora(s).")

    # ── SELECT ────────────────────────────────

    def _on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx < len(self.printers_data):
            self.selected_printer = self.printers_data[idx]

    # ── DIAGNOSE ──────────────────────────────

    def _diagnose_selected(self):
        if not self.selected_printer:
            messagebox.showwarning("Aviso", "Selecione uma impressora na lista.")
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
        self.after(0, self._clear_diag)
        self.after(0, self._show_placeholder)

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
        sep = "─" * 60 + "\n"
        self._write_diag(f"\n{sep}", "muted")
        self._write_diag(f"  🖨  {diag['name']}\n", "header")
        self._write_diag(f"{sep}", "muted")

        self._write_diag(f"  Porta   : ", "muted")
        self._write_diag(f"{diag['port']}\n")
        self._write_diag(f"  Driver  : ", "muted")
        self._write_diag(f"{diag['driver']}\n")
        self._write_diag(f"  Tipo    : ", "muted")
        self._write_diag(f"{diag['type']}\n")
        self._write_diag(f"  Status  : ", "muted")
        self._write_diag(f"{diag['status_text']}\n")

        if diag.get("offline"):
            self._write_diag("  ⚠ OFFLINE — impressora definida como offline no Windows!\n", "warning")

        if diag.get("network_check"):
            nc = diag["network_check"]
            self._write_diag("\n  [ Verificação de Rede ]\n", "bold")
            for det in nc.get("details", []):
                tag = "ok" if "✓" in det else ("error" if "✗" in det else "warning")
                self._write_diag(f"    {det}\n", tag)

        if diag.get("usb_check"):
            uc = diag["usb_check"]
            self._write_diag("\n  [ Verificação de Driver USB ]\n", "bold")
            for det in uc.get("details", []):
                tag = "ok" if "corretamente" in det.lower() else "error"
                self._write_diag(f"    {det}\n", tag)

        if diag["issues"]:
            self._write_diag("\n  [ Problemas Encontrados ]\n", "bold")
            for issue in diag["issues"]:
                self._write_diag(f"    ✗ {issue}\n", "error")
        else:
            self._write_diag("\n  ✓ Nenhum problema detectado nesta impressora.\n", "ok")

        self._write_diag("\n")

    def _render_fixes(self):
        for w in self.fix_frame.winfo_children():
            w.destroy()

        # Coleta todas as correções disponíveis
        all_fixes = []
        for name, diag in self.diagnoses.items():
            for fix in diag.get("fixes", []):
                all_fixes.append((name, fix, diag))

        if not all_fixes:
            tk.Label(self.fix_frame,
                     text="✓ Nenhuma correção automática necessária!",
                     bg=THEME["bg"], fg=THEME["success"],
                     font=("Segoe UI", 13, "bold")).pack(pady=30)
            return

        tk.Label(self.fix_frame,
                 text=f"{len(all_fixes)} correção(ões) automática(s) disponível(is):",
                 bg=THEME["bg"], fg=THEME["text"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 12))

        fix_labels = {
            "set_online":       ("⚡ Definir como Online",             THEME["success"]),
            "clear_queue":      ("🗑  Limpar Fila de Impressão",        THEME["warning"]),
            "reinstall_driver": ("🔄  Reinstalar Driver",               THEME["accent"]),
            "subnet_guide":     ("🌐  Ver Guia: Sub-rede Diferente",    "#7dd3fc"),
            "ping_fail_guide":  ("📡  Ver Guia: Impressora Sem Resposta", THEME["warning"]),
        }

        for printer_name, fix_key, diag in all_fixes:
            card = tk.Frame(self.fix_frame, bg=THEME["surface"], pady=10)
            card.pack(fill="x", pady=4)

            label, color = fix_labels.get(fix_key, (fix_key, THEME["text_muted"]))
            is_guide = fix_key in ("subnet_guide", "ping_fail_guide")

            tk.Label(card, text=f"  {label}",
                     bg=THEME["surface"], fg=color,
                     font=("Segoe UI", 11, "bold")).pack(side="left", padx=8)
            tk.Label(card, text=f"→  {printer_name}",
                     bg=THEME["surface"], fg=THEME["text_muted"],
                     font=("Segoe UI", 10)).pack(side="left")

            btn_text = "Ver Guia" if is_guide else "Aplicar"
            btn = tk.Button(card, text=btn_text,
                            command=lambda pn=printer_name, fk=fix_key, d=diag: self._apply_fix(pn, fk, d),
                            bg=color, fg="#0f1117",
                            font=("Segoe UI", 9, "bold"),
                            relief="flat", cursor="hand2", padx=12, pady=4)
            btn.pack(side="right", padx=12)

        # Só mostra "Aplicar Todas" para fixes automáticos reais
        auto_fixes = [(pn, fk, d) for pn, fk, d in all_fixes if fk not in ("subnet_guide", "ping_fail_guide")]
        if auto_fixes:
            tk.Button(self.fix_frame, text="⚡  Aplicar Todas as Correções Automáticas",
                      command=lambda: self._apply_all_fixes(auto_fixes),
                      bg=THEME["accent"], fg="white",
                      font=("Segoe UI", 11, "bold"),
                      relief="flat", cursor="hand2",
                      padx=16, pady=10).pack(fill="x", pady=(16, 0))

    # ── FIXES ─────────────────────────────────

    def _apply_fix(self, printer_name: str, fix_key: str, diag: dict):
        self._log(f"Aplicando correção '{fix_key}' em '{printer_name}'...")
        threading.Thread(target=self._fix_thread,
                         args=(printer_name, fix_key, diag), daemon=True).start()

    def _apply_all_fixes(self, fixes: list):
        for printer_name, fix_key, diag in fixes:
            threading.Thread(target=self._fix_thread,
                             args=(printer_name, fix_key, diag), daemon=True).start()

    def _fix_thread(self, printer_name: str, fix_key: str, diag: dict):
        if fix_key == "subnet_guide":
            nc = diag.get("network_check", {})
            printer_ip = nc.get("printer_ip", "?")
            server_ip  = nc.get("server_ip", "?")
            self.after(0, self._render_subnet_guide, printer_name, printer_ip, server_ip)
            self.after(0, self._log, f"Guia de sub-rede exibido para '{printer_name}'.")
            return
        if fix_key == "ping_fail_guide":
            self.after(0, self._render_ping_fail_guide, printer_name, diag)
            self.after(0, self._log, f"Guia de falha de ping exibido para '{printer_name}'.")
            return
        result = ""
        if fix_key == "set_online":
            result = fix_set_online(printer_name)
        elif fix_key == "clear_queue":
            result = fix_clear_queue(printer_name)
        elif fix_key == "reinstall_driver":
            result = fix_reinstall_driver(printer_name, diag.get("driver", ""))
        self.after(0, self._log, f"✓ {result}")
        self.after(0, self._set_status, result)

    def _render_ping_fail_guide(self, printer_name: str, diag: dict):
        """Guia para quando a impressora não responde ao ping."""
        nc = diag.get("network_check", {})
        printer_ip = nc.get("printer_ip", "?")

        self.subnet_text.config(state="normal")
        self.subnet_text.delete("1.0", "end")

        self._write_subnet("═" * 62 + "\n", "muted")
        self._write_subnet(f"  📡  GUIA: Impressora Sem Resposta (Ping Falhou)\n", "title")
        self._write_subnet(f"  Impressora: {printer_name}  |  IP: {printer_ip}\n", "muted")
        self._write_subnet("═" * 62 + "\n\n", "muted")

        self._write_subnet("  POSSÍVEIS CAUSAS\n", "step")
        self._write_subnet("""
  1. Impressora desligada ou em modo de espera profundo
  2. Cabo de rede desconectado
  3. IP da impressora mudou (DHCP)
  4. Impressora com defeito de hardware
  5. Firewall bloqueando o ping\n\n""", "muted")

        self._write_subnet("  PASSO 1 — Verificar fisicamente\n", "step")
        self._write_subnet("""
  • Confirme com o cliente que a impressora está LIGADA
  • Verifique se o cabo de rede está conectado
  • Veja se o LED de rede da impressora está aceso\n\n""", "muted")

        self._write_subnet("  PASSO 2 — Descobrir o IP atual da impressora\n", "step")
        self._write_subnet("""
  Na impressora, imprima uma página de configuração de rede:
  • HP     → Segure o botão de informações por 3 segundos
  • Epson  → Menu > Configuração > Status da Rede
  • Brother → Menu > Informações da Máquina > Impr. config. rede
  • Canon  → Menu > Config. dispositivo > Imprimir folha de status\n\n""", "muted")

        self._write_subnet("  PASSO 3 — Varredura de rede para encontrar a impressora\n", "step")
        self._write_subnet("  Execute no PowerShell:\n", "muted")
        self._write_subnet(f'  1..254 | % {{ $ip = "192.168.1.$_"; if (Test-Connection $ip -Count 1 -Quiet) {{ $ip }} }}\n', "cmd")
        self._write_subnet("  (Ajuste '192.168.1' para a faixa da rede do cliente)\n\n", "muted")

        self._write_subnet("  PASSO 4 — Atualizar o IP na porta do Windows\n", "step")
        self._write_subnet("  Se o IP mudou, atualize a porta:\n", "muted")
        self._write_subnet(f'  Set-PrinterPort -Name "IP_{printer_ip}" -PrinterHostAddress "NOVO_IP"\n', "cmd")

        self.subnet_text.config(state="disabled")
        self.tabs.select(3)

    def _restart_spooler(self):
        if messagebox.askyesno("Confirmar", "Reiniciar o serviço Spooler agora?\n\nIsso pode interromper impressões em andamento."):
            threading.Thread(target=self._spooler_thread, daemon=True).start()

    def _spooler_thread(self):
        self.after(0, self._set_status, "Reiniciando Spooler...")
        self.after(0, self._log, "Reiniciando serviço Spooler...")
        result = fix_restart_spooler()
        self.after(0, self._log, f"✓ {result}")
        self.after(0, self._set_status, result)
        self.after(0, messagebox.showinfo, "Sucesso", result)


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    app = PrinterDiagApp()
    app.mainloop()
