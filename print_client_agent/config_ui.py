import os
import sys
import threading
import socket
import winreg
import tkinter as tk
from tkinter import ttk, messagebox

import pystray
import win32print
from PIL import Image

from agent import run_agent
from config import AgentConfig, app_base_dir, default_config_path, load_config, save_config


_SINGLETON_SOCK = None


def ensure_single_instance() -> bool:
    global _SINGLETON_SOCK
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(("127.0.0.1", 53991))
        s.listen(1)
        _SINGLETON_SOCK = s
        return True
    except Exception:
        return False


def list_printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = win32print.EnumPrinters(flags)
    return [p[2] for p in printers]


def startup_command() -> str:
    if getattr(sys, "frozen", False):
        return f'"{os.path.abspath(sys.executable)}"'
    return f'"{sys.executable}" "{os.path.abspath(__file__)}"'


def set_startup(enable: bool):
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Run",
        0,
        winreg.KEY_ALL_ACCESS,
    )
    name = "PrintClientAgent"
    if enable:
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, startup_command())
    else:
        try:
            winreg.DeleteValue(key, name)
        except FileNotFoundError:
            pass
    winreg.CloseKey(key)


def main():
    if not ensure_single_instance():
        messagebox.showinfo("Print Client Agent", "O agent ja esta em execucao.")
        return

    root = tk.Tk()
    root.title("Print Client Agent")
    root.geometry("560x360")
    root.resizable(False, False)
    root.configure(bg="#f5f7fb")

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TFrame", background="#f5f7fb")
    style.configure("TLabel", background="#f5f7fb", foreground="#1f2328")
    style.configure("TEntry", fieldbackground="#ffffff", foreground="#1f2328")
    style.configure("TCombobox", fieldbackground="#ffffff", foreground="#1f2328")
    style.configure("TCheckbutton", background="#f5f7fb", foreground="#1f2328")
    style.configure("TButton", padding=6)

    base_dir = app_base_dir()
    cfg_path = default_config_path()
    try:
        cfg = load_config(cfg_path)
    except Exception:
        cfg = AgentConfig(
            server_url="",
            printer_name="",
            poll_interval_sec=5,
            start_with_windows=False,
            monitor_default_printer=False,
        )

    # Compat: builds antigos salvaram use_default_printer
    monitor_default_initial = bool(
        getattr(cfg, "monitor_default_printer", False)
        or getattr(cfg, "use_default_printer", False)
    )

    frm = ttk.Frame(root, padding=16)
    frm.pack(fill="both", expand=True)

    title = ttk.Label(frm, text="Configuracao do Agente", font=("Segoe UI", 12, "bold"))
    title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

    ttk.Label(frm, text="Servidor (URL)").grid(row=1, column=0, sticky="w", pady=(6, 2))
    server_var = tk.StringVar(value=cfg.server_url)
    ttk.Entry(frm, textvariable=server_var, width=52).grid(row=1, column=1, sticky="w", pady=(6, 2))

    ttk.Label(frm, text="Impressora").grid(row=2, column=0, sticky="w", pady=(6, 2))
    printers = list_printers()
    printer_var = tk.StringVar(value=cfg.printer_name if cfg.printer_name in printers else "")
    printer_combo = ttk.Combobox(frm, textvariable=printer_var, values=printers, width=49)
    printer_combo.grid(row=2, column=1, sticky="w", pady=(6, 2))

    monitor_default_var = tk.BooleanVar(value=monitor_default_initial)
    ttk.Checkbutton(
        frm,
        text="Monitorar automaticamente a impressora padrao do Windows",
        variable=monitor_default_var,
    ).grid(row=3, column=1, sticky="w", pady=(2, 2))

    ttk.Label(frm, text="Intervalo (segundos)").grid(row=4, column=0, sticky="w", pady=(6, 2))
    interval_var = tk.IntVar(value=cfg.poll_interval_sec)
    ttk.Entry(frm, textvariable=interval_var, width=10).grid(row=4, column=1, sticky="w", pady=(6, 2))

    start_var = tk.BooleanVar(value=cfg.start_with_windows)
    ttk.Checkbutton(frm, text="Iniciar com Windows", variable=start_var).grid(row=5, column=1, sticky="w", pady=(8, 2))

    status_var = tk.StringVar(value="Coleta: parada")
    ttk.Label(frm, textvariable=status_var).grid(row=6, column=1, sticky="w", pady=(8, 2))

    def refresh_printer_field_state():
        try:
            printer_combo.configure(state="disabled" if monitor_default_var.get() else "readonly")
        except Exception:
            pass

    monitor_default_var.trace_add("write", lambda *_: refresh_printer_field_state())
    refresh_printer_field_state()

    worker_thread = None
    worker_stop = None

    def stop_worker():
        nonlocal worker_thread, worker_stop
        if worker_stop is not None:
            worker_stop.set()
        if worker_thread is not None and worker_thread.is_alive():
            worker_thread.join(timeout=2)
        worker_thread = None
        worker_stop = None
        status_var.set("Coleta: parada")

    def start_worker(show_errors: bool = True) -> bool:
        nonlocal worker_thread, worker_stop

        server_url = server_var.get().strip()
        printer_name = printer_var.get().strip()
        auto_default = bool(monitor_default_var.get())

        if not server_url or (not auto_default and not printer_name):
            status_var.set("Coleta: configure servidor e impressora")
            if show_errors:
                messagebox.showwarning("Coleta", "Configure servidor e impressora para iniciar a coleta.")
            return False

        stop_worker()

        worker_cfg = AgentConfig(
            server_url=server_url,
            printer_name=printer_name,
            poll_interval_sec=int(interval_var.get() or 5),
            start_with_windows=bool(start_var.get()),
            monitor_default_printer=auto_default,
        )
        worker_stop = threading.Event()

        def _runner():
            try:
                run_agent(worker_cfg, stop_event=worker_stop)
            except Exception as e:
                root.after(0, lambda: status_var.set(f"Coleta erro: {e}"))

        worker_thread = threading.Thread(target=_runner, daemon=True)
        worker_thread.start()
        if auto_default:
            status_var.set("Coleta: ativa (impressora padrao)")
        else:
            status_var.set("Coleta: ativa")
        return True

    def on_save(show_success: bool = True):
        new_cfg = AgentConfig(
            server_url=server_var.get().strip(),
            printer_name=printer_var.get().strip(),
            poll_interval_sec=int(interval_var.get() or 5),
            start_with_windows=bool(start_var.get()),
            monitor_default_printer=bool(monitor_default_var.get()),
        )
        save_config(new_cfg, cfg_path)
        set_startup(new_cfg.start_with_windows)
        started = start_worker(show_errors=False)
        if show_success:
            if started:
                messagebox.showinfo("Salvo", "Configuracoes salvas e coleta iniciada.")
            else:
                messagebox.showinfo("Salvo", "Configuracoes salvas. Configure servidor/impressora para iniciar a coleta.")

    ttk.Button(frm, text="Salvar", command=on_save).grid(row=7, column=1, sticky="e", pady=16)

    icon = None

    def show_window():
        root.deiconify()
        root.after(10, root.focus_force)

    def hide_window():
        root.withdraw()

    def on_exit():
        stop_worker()
        try:
            if icon:
                icon.stop()
        except Exception:
            pass
        root.destroy()

    def on_close():
        hide_window()

    def _load_icon():
        icon_path = os.path.join(base_dir, "icon.ico")
        try:
            return Image.open(icon_path)
        except Exception:
            return Image.new("RGB", (64, 64), color="#3aa0ff")

    def setup_tray():
        nonlocal icon
        menu = pystray.Menu(
            pystray.MenuItem("Abrir configuracao", lambda: root.after(0, show_window)),
            pystray.MenuItem("Salvar configuracao", lambda: root.after(0, on_save)),
            pystray.MenuItem("Iniciar coleta", lambda: root.after(0, lambda: start_worker(True))),
            pystray.MenuItem("Parar coleta", lambda: root.after(0, stop_worker)),
            pystray.MenuItem("Sair", lambda: root.after(0, on_exit)),
        )
        icon = pystray.Icon("PrintClientAgent", _load_icon(), "Print Client Agent", menu)
        threading.Thread(target=icon.run, daemon=True).start()

    root.protocol("WM_DELETE_WINDOW", on_close)
    setup_tray()
    start_worker(show_errors=False)
    root.mainloop()


if __name__ == "__main__":
    main()
