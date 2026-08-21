"""
Interface gráfica da Automação COB (Fusion).

Não altera main.py / models.py / utils.py. Replica o fluxo do main.py
e exibe os prints da execução em um console ao vivo.

Para gerar o .exe (na pasta do projeto, com o venv ativo):

    pyinstaller --noconfirm --windowed --name "AutomacaoCOB" ^
        --add-data "assets;assets" gui.py
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from datetime import datetime
from tkinter import filedialog, messagebox
import tkinter as tk


# Mesmos valores usados em main.py
COD_FILIAL = "01MG0014"
COD_UO = "10310"
LOCAL_DESTINO = "C:/Temp/Historico.xlsx"


# ---------------------------------------------------------------------------
# Paleta — tons de capivara / terra
# ---------------------------------------------------------------------------
C = {
    "bg": "#F3EBE1",
    "bg_soft": "#EFE4D6",
    "header": "#2B1F1A",
    "sidebar": "#2F231E",
    "card": "#3B2C26",
    "card_alt": "#44342D",
    "line": "#5A463C",
    "accent": "#C4783A",
    "accent_h": "#D4894A",
    "green": "#5E8F64",
    "green_h": "#6FA375",
    "red": "#B85A4C",
    "red_h": "#C96B5D",
    "text": "#F6EFE6",
    "muted": "#C9B8A8",
    "ink": "#2B1F1A",
    "ink_soft": "#5C4A40",
    "console_bg": "#16110F",
    "console_fg": "#E8DCC8",
    "ok": "#8FCB8A",
    "err": "#E07A6A",
    "info": "#7EB6D4",
    "warn": "#E0B86A",
    "idle": "#8A9E7A",
}


def resource_path(*parts: str) -> str:
    """Caminho de arquivos empacotados no exe ou ao lado do script."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS  # type: ignore[attr-defined]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)


def enable_dpi_awareness() -> None:
    if os.name != "nt":
        return
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            from ctypes import windll

            windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Worker — mesmo fluxo do main.py, sem input() de menu
# ---------------------------------------------------------------------------
def run_worker(opcao: int, caminho: str) -> None:
    """Executa a automação. Roda em processo separado para o Tk da GUI
    não conflitar com os diálogos já existentes em models.py."""
    if sys.stdout is None:
        sys.stdout = sys.__stdout__
    if sys.stderr is None:
        sys.stderr = sys.__stderr__
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:
        pass

    from utils import copiar_para_planilha
    from models import AutomacaoFusion

    print("-----------Automação COB---------")
    print(f"Planilha : {caminho}")
    print(f"Operação : {'Medição Variável' if opcao == 1 else 'Criar Novos COBs'}")
    print(f"Filial   : {COD_FILIAL}    UO: {COD_UO}")
    print("-" * 42)

    destino = os.path.dirname(caminho)
    planilha_destino = destino + "/Historico.xlsx"
    local_destino = LOCAL_DESTINO
    global_instance = None

    try:
        if opcao == 1:
            global_instance = AutomacaoFusion(
                caminho, None, None, None, planilha_destino, local_destino,
                "medicao_vr", cod_filial=COD_FILIAL, cod_uo=COD_UO,
            )
            navegador, chrome_proc, planilha = global_instance.inicializacao("Medição")
            global_instance.navegador = navegador
            global_instance.chrome_proc = chrome_proc
            global_instance.planilha = planilha
            print("Iniciando Medição Variável...")
            global_instance.medicao_vr()
            print("Medição Variável finalizada.")
        elif opcao == 2:
            global_instance = AutomacaoFusion(
                caminho, None, None, None, planilha_destino, local_destino,
                "cob_nv", cod_filial=COD_FILIAL, cod_uo=COD_UO,
            )
            navegador, chrome_proc, planilha = global_instance.inicializacao("Novo")
            global_instance.navegador = navegador
            global_instance.chrome_proc = chrome_proc
            global_instance.planilha = planilha
            print("Iniciando criação de Novos COBs...")
            global_instance.cob_nv()
            print("Criação de Novos COBs finalizada.")
        else:
            print("Escolha inválida.")
    except SystemExit as e:
        print(f"Encerrado: {e}")
    except Exception as e:
        if global_instance:
            tipo = "Medição" if opcao == 1 else "Novo"
            global_instance.tratar_erro_critico(e, tipo)
        else:
            print(f"Ocorreu um erro: {e}")
            try:
                copiar_para_planilha(planilha_destino, local_destino)
            except Exception:
                pass
            import tkinter as _tk
            from tkinter import messagebox as _mb

            root = _tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            _mb.showerror("Erro", f"Ocorreu um erro:\n{e}\n\nChame a T.I.")
            root.destroy()


# ---------------------------------------------------------------------------
# Widgets customizados
# ---------------------------------------------------------------------------
class RoundedButton(tk.Canvas):
    def __init__(
        self,
        master,
        text: str,
        command=None,
        bg_color: str = C["accent"],
        hover_color: str = C["accent_h"],
        fg: str = C["text"],
        radius: int = 12,
        height: int = 44,
        canvas_bg: str | None = None,
        font=("Segoe UI", 11, "bold"),
        **kwargs,
    ):
        super().__init__(
            master,
            height=height,
            highlightthickness=0,
            bd=0,
            bg=canvas_bg or C["sidebar"],
            **kwargs,
        )
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.disabled_color = "#6A564C"
        self.fg = fg
        self.radius = radius
        self.label = text
        self.font = font
        self._enabled = True
        self._current = bg_color
        self.bind("<Configure>", self._redraw)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def set_enabled(self, enabled: bool) -> None:
        self._enabled = enabled
        self._current = self.bg_color if enabled else self.disabled_color
        self.configure(cursor="hand2" if enabled else "arrow")
        self._redraw()

    def _on_enter(self, _=None):
        if self._enabled:
            self._current = self.hover_color
            self.configure(cursor="hand2")
            self._redraw()

    def _on_leave(self, _=None):
        self._current = self.bg_color if self._enabled else self.disabled_color
        self._redraw()

    def _on_click(self, _=None):
        if self._enabled and self.command:
            self.command()

    def _redraw(self, _=None):
        self.delete("all")
        w = max(self.winfo_width(), 10)
        h = max(self.winfo_height(), 10)
        r = min(self.radius, h // 2, w // 2)
        self._round_rect(2, 2, w - 2, h - 2, r, fill=self._current)
        self.create_text(w / 2, h / 2, text=self.label, fill=self.fg, font=self.font)

    def _round_rect(self, x1, y1, x2, y2, r, **kwargs):
        self.create_polygon(
            x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
            x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
            x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
            smooth=True, **kwargs,
        )


class StatusPill(tk.Canvas):
    def __init__(self, master, **kwargs):
        super().__init__(master, width=168, height=32, highlightthickness=0, bd=0, bg=C["header"], **kwargs)
        self._text = "Pronto"
        self._color = C["idle"]
        self._draw()

    def set_status(self, text: str, color: str) -> None:
        self._text = text
        self._color = color
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h = 168, 32
        self.create_oval(4, 4, h - 4, h - 4, fill="#1A1410", outline="")
        self.create_oval(10, 10, 22, 22, fill=self._color, outline="")
        self.create_text(34, h / 2, text=self._text, fill=C["text"], font=("Segoe UI", 9, "bold"), anchor="w")


# ---------------------------------------------------------------------------
# Aplicação
# ---------------------------------------------------------------------------
class AutomacaoCOBGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação COB  ·  Fusion")
        self.configure(bg=C["bg"])
        self.minsize(980, 640)
        self.geometry("1180x740")
        self.caminho = ""
        self.proc: subprocess.Popen | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()
        self._reader_thread: threading.Thread | None = None
        self._photo = None
        self._icon = None
        self._log_lines = 0
        self.footer_label = None

        self._load_icon()
        self._build()
        self._poll_log()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Configure>", self._on_resize)

    # -- assets --------------------------------------------------------------
    def _load_photo(self, max_size: int):
        path = resource_path("assets", "capivara.png")
        if not os.path.isfile(path):
            return None
        try:
            img = tk.PhotoImage(file=path)
            side = max(img.width(), img.height())
            factor = max(1, side // max_size)
            if factor > 1:
                img = img.subsample(factor, factor)
            return img
        except Exception:
            return None

    def _load_icon(self) -> None:
        icon = self._load_photo(64)
        if icon:
            self._icon = icon
            try:
                self.iconphoto(True, icon)
            except Exception:
                pass

    def _mascot_image(self, max_size: int = 96):
        return self._load_photo(max_size)

    # -- layout --------------------------------------------------------------
    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self._build_header()

        body = tk.PanedWindow(
            self, orient=tk.HORIZONTAL, sashwidth=6, sashrelief=tk.FLAT,
            bg=C["line"], bd=0,
        )
        body.grid(row=1, column=0, sticky="nsew")
        self.body = body

        sidebar = tk.Frame(body, bg=C["sidebar"], width=340)
        sidebar.pack_propagate(False)
        self._build_sidebar(sidebar)

        console_wrap = tk.Frame(body, bg=C["bg"])
        self._build_console(console_wrap)

        body.add(sidebar, minsize=280, stretch="never")
        body.add(console_wrap, minsize=480, stretch="always")
        self.after(80, lambda: body.sash_place(0, 340, 0))

        self._build_footer()

    def _build_header(self) -> None:
        header = tk.Frame(self, bg=C["header"], height=104)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.columnconfigure(1, weight=1)
        header.rowconfigure(0, weight=1)

        mascot_size = 72
        mascot_box = tk.Frame(header, bg=C["header"], width=mascot_size, height=mascot_size)
        mascot_box.grid(row=0, column=0, padx=(18, 4), pady=16, sticky="nsw")
        mascot_box.grid_propagate(False)
        mascot_box.pack_propagate(False)

        mascot_canvas = tk.Canvas(
            mascot_box, width=mascot_size, height=mascot_size,
            bg=C["header"], highlightthickness=0, bd=0,
        )
        mascot_canvas.pack()
        self._photo = self._mascot_image(mascot_size)
        if self._photo:
            mascot_canvas.create_image(mascot_size // 2, mascot_size // 2, image=self._photo)
        else:
            mascot_canvas.create_text(
                mascot_size // 2, mascot_size // 2, text="🦫",
                font=("Segoe UI Emoji", 28), fill=C["accent"],
            )

        titles = tk.Frame(header, bg=C["header"])
        titles.grid(row=0, column=1, sticky="w", padx=(18, 8))
        tk.Label(
            titles, text="Automação COB", font=("Segoe UI", 20, "bold"),
            bg=C["header"], fg=C["text"],
        ).pack(anchor="w")
        tk.Label(
            titles, text="Fusion  ·  Medição variável e novos COBs",
            font=("Segoe UI", 10), bg=C["header"], fg=C["muted"],
        ).pack(anchor="w", pady=(6, 0))

        right = tk.Frame(header, bg=C["header"])
        right.grid(row=0, column=2, padx=20, sticky="e")
        self.status_pill = StatusPill(right)
        self.status_pill.pack()

    def _section_title(self, parent, text: str) -> None:
        row = tk.Frame(parent, bg=C["sidebar"])
        row.pack(fill="x", padx=18, pady=(16, 6))
        tk.Label(
            row, text=text.upper(), font=("Segoe UI", 8, "bold"),
            bg=C["sidebar"], fg=C["accent"],
        ).pack(anchor="w")
        tk.Frame(row, bg=C["line"], height=1).pack(fill="x", pady=(6, 0))

    def _build_sidebar(self, parent: tk.Frame) -> None:
        parent.columnconfigure(0, weight=1)

        self._section_title(parent, "Planilha")

        file_card = tk.Frame(parent, bg=C["card"], padx=12, pady=12)
        file_card.pack(fill="x", padx=18)

        self.file_label = tk.Label(
            file_card,
            text="Nenhuma planilha selecionada",
            font=("Segoe UI", 9),
            bg=C["card"],
            fg=C["muted"],
            wraplength=260,
            justify="left",
            anchor="w",
        )
        self.file_label.pack(fill="x", pady=(0, 10))

        self.btn_file = RoundedButton(
            file_card, "Selecionar planilha COB",
            command=self._escolher_planilha,
            bg_color=C["card_alt"], hover_color=C["line"], height=40,
            canvas_bg=C["card"],
        )
        self.btn_file.pack(fill="x")

        self._section_title(parent, "Operações")

        ops = tk.Frame(parent, bg=C["sidebar"])
        ops.pack(fill="x", padx=18)

        self.btn_medicao = RoundedButton(
            ops, "1  ·  Medição Variável",
            command=lambda: self._iniciar(1),
            bg_color=C["green"], hover_color=C["green_h"], height=48,
        )
        self.btn_medicao.pack(fill="x", pady=(0, 10))

        self.btn_novos = RoundedButton(
            ops, "2  ·  Criar Novos COBs",
            command=lambda: self._iniciar(2),
            bg_color=C["accent"], hover_color=C["accent_h"], height=48,
        )
        self.btn_novos.pack(fill="x", pady=(0, 10))

        self.btn_parar = RoundedButton(
            ops, "Parar execução",
            command=self._parar,
            bg_color=C["red"], hover_color=C["red_h"], height=40,
        )
        self.btn_parar.pack(fill="x")
        self.btn_parar.set_enabled(False)

        self._section_title(parent, "Parâmetros")

        params = tk.Frame(parent, bg=C["card"], padx=12, pady=10)
        params.pack(fill="x", padx=18)
        for label, value in (
            ("Filial", COD_FILIAL),
            ("UO", COD_UO),
            ("Histórico", LOCAL_DESTINO),
        ):
            row = tk.Frame(params, bg=C["card"])
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label, font=("Segoe UI", 8), bg=C["card"], fg=C["muted"], width=10, anchor="w").pack(side="left")
            tk.Label(row, text=value, font=("Consolas", 9), bg=C["card"], fg=C["text"], anchor="w").pack(side="left", fill="x", expand=True)

        hint = tk.Label(
            parent,
            text="A capivara cuida do Fusion enquanto\nvocê acompanha o console ao lado.",
            font=("Segoe UI", 8),
            bg=C["sidebar"],
            fg=C["muted"],
            justify="left",
        )
        hint.pack(side="bottom", anchor="w", padx=18, pady=18)

    def _build_console(self, parent: tk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        parent.configure(bg=C["bg"])

        bar = tk.Frame(parent, bg=C["bg"])
        bar.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        tk.Label(
            bar, text="Console de execução", font=("Segoe UI", 12, "bold"),
            bg=C["bg"], fg=C["ink"],
        ).pack(side="left")

        tk.Button(
            bar, text="Limpar", command=self._limpar_console,
            bg=C["bg_soft"], fg=C["ink_soft"], relief="flat",
            font=("Segoe UI", 9), cursor="hand2", padx=10, pady=4,
            activebackground=C["bg"],
        ).pack(side="right")

        card = tk.Frame(parent, bg=C["console_bg"])
        card.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 8))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(0, weight=1)

        self.console = tk.Text(
            card,
            wrap="word",
            bg=C["console_bg"],
            fg=C["console_fg"],
            insertbackground=C["console_fg"],
            relief="flat",
            font=("Consolas", 10),
            padx=14,
            pady=12,
            state="disabled",
            highlightthickness=0,
            bd=0,
        )
        scroll = tk.Scrollbar(card, command=self.console.yview, bg=C["console_bg"], troughcolor=C["console_bg"])
        self.console.configure(yscrollcommand=scroll.set)
        self.console.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        self.console.tag_configure("ts", foreground="#8A7A6C")
        self.console.tag_configure("ok", foreground=C["ok"])
        self.console.tag_configure("err", foreground=C["err"])
        self.console.tag_configure("info", foreground=C["info"])
        self.console.tag_configure("warn", foreground=C["warn"])
        self.console.tag_configure("plain", foreground=C["console_fg"])

        self._append_log("Capivara pronta. Selecione a planilha e escolha uma operação.", "info")

    def _build_footer(self) -> None:
        footer = tk.Frame(self, bg=C["header"], height=32)
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_propagate(False)
        self.footer_label = tk.Label(
            footer,
            text="Pronto  ·  0 linhas no console",
            font=("Segoe UI", 8),
            bg=C["header"],
            fg=C["muted"],
            anchor="w",
        )
        self.footer_label.pack(side="left", padx=16)
        tk.Label(
            footer, text="main.py  ·  sem alterações",
            font=("Segoe UI", 8), bg=C["header"], fg=C["line"],
        ).pack(side="right", padx=16)
        self.footer_label.configure(text=f"Pronto  ·  {self._log_lines} linhas no console")

    # -- ações ---------------------------------------------------------------
    def _escolher_planilha(self) -> None:
        caminho = filedialog.askopenfilename(
            parent=self,
            title="Selecione a Planilha COB!",
            filetypes=[
                ("Planilhas Excel", "*.xlsx *.xls"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not caminho:
            return
        self.caminho = caminho
        nome = os.path.basename(caminho)
        pasta = os.path.dirname(caminho)
        self.file_label.configure(
            text=f"{nome}\n{pasta}",
            fg=C["text"],
        )
        self._append_log(f"Planilha selecionada: {caminho}", "info")
        self._set_status("Planilha ok", C["ok"])

    def _iniciar(self, opcao: int) -> None:
        if self.proc and self.proc.poll() is None:
            messagebox.showinfo("Em execução", "Já existe uma automação em andamento.")
            return
        if not self.caminho:
            messagebox.showwarning("Planilha", "Selecione a planilha COB antes de iniciar.")
            return
        if not os.path.isfile(self.caminho):
            messagebox.showerror("Planilha", "O arquivo selecionado não existe mais.")
            return

        nome = "Medição Variável" if opcao == 1 else "Criar Novos COBs"
        self._append_log(f"Iniciando {nome}...", "warn")
        self._set_running(True, nome)

        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"

        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--worker", "--opcao", str(opcao), "--planilha", self.caminho]
        else:
            cmd = [
                sys.executable, "-u", os.path.abspath(__file__),
                "--worker", "--opcao", str(opcao), "--planilha", self.caminho,
            ]

        kwargs = dict(
            args=cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
            cwd=os.path.dirname(os.path.abspath(__file__)),
        )
        if os.name == "nt":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

        try:
            self.proc = subprocess.Popen(**kwargs)
        except Exception as e:
            self._append_log(f"Falha ao iniciar: {e}", "err")
            self._set_running(False)
            return

        self._reader_thread = threading.Thread(target=self._read_stdout, daemon=True)
        self._reader_thread.start()
        self.after(400, self._watch_proc)

    def _read_stdout(self) -> None:
        if not self.proc or not self.proc.stdout:
            return
        try:
            for line in self.proc.stdout:
                self.log_queue.put(line.rstrip("\r\n"))
        except Exception:
            pass
        self.log_queue.put("__PROC_DONE__")

    def _watch_proc(self) -> None:
        still_running = self.proc and self.proc.poll() is None
        draining = (
            not self.log_queue.empty()
            or (self._reader_thread is not None and self._reader_thread.is_alive())
        )
        if still_running or draining:
            self.after(250, self._watch_proc)
            return
        code = self.proc.returncode if self.proc else 0
        if code == 0:
            self._append_log("Execução concluída.", "ok")
            self._set_status("Concluído", C["ok"])
        else:
            self._append_log(f"Processo finalizado (código {code}).", "err")
            self._set_status("Encerrado", C["err"])
        self._set_running(False)

    def _parar(self) -> None:
        if not self.proc or self.proc.poll() is not None:
            return
        self._append_log("Parando execução a pedido do usuário...", "warn")
        try:
            self.proc.terminate()
        except Exception:
            pass
        self.after(1500, self._kill_if_alive)

    def _kill_if_alive(self) -> None:
        if self.proc and self.proc.poll() is None:
            try:
                self.proc.kill()
            except Exception:
                pass

    def _set_running(self, running: bool, nome: str = "") -> None:
        self.btn_medicao.set_enabled(not running)
        self.btn_novos.set_enabled(not running)
        self.btn_file.set_enabled(not running)
        self.btn_parar.set_enabled(running)
        if running:
            self._set_status("Em execução", C["warn"])
            self.footer_label.configure(text=f"Executando {nome}...")
        else:
            self.footer_label.configure(text=f"Pronto  ·  {self._log_lines} linhas no console")

    def _set_status(self, text: str, color: str) -> None:
        self.status_pill.set_status(text, color)

    def _limpar_console(self) -> None:
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")
        self._log_lines = 0
        self.footer_label.configure(text="Pronto  ·  0 linhas no console")

    def _append_log(self, text: str, kind: str | None = None) -> None:
        if not text:
            return
        ts = datetime.now().strftime("%H:%M:%S")
        tag = kind or self._classify(text)
        self.console.configure(state="normal")
        self.console.insert("end", f"[{ts}]  ", "ts")
        self.console.insert("end", text + "\n", tag)
        self.console.see("end")
        self.console.configure(state="disabled")
        self._log_lines += 1
        if self.footer_label is not None and (self.proc is None or self.proc.poll() is not None):
            self.footer_label.configure(text=f"Pronto  ·  {self._log_lines} linhas no console")

    def _classify(self, text: str) -> str:
        lower = text.lower()
        if "sucesso" in lower or "finalizada" in lower or "conclu" in lower:
            return "ok"
        if "erro" in lower or "falha" in lower or "encerrado" in lower:
            return "err"
        if "iniciando" in lower or "planilha" in lower or "automação cob" in lower:
            return "info"
        if "chame" in lower or "atenção" in lower:
            return "warn"
        return "plain"

    def _poll_log(self) -> None:
        try:
            while True:
                line = self.log_queue.get_nowait()
                if line == "__PROC_DONE__":
                    continue
                self._append_log(line)
        except queue.Empty:
            pass
        self.after(80, self._poll_log)

    def _on_resize(self, event) -> None:
        if event.widget is not self:
            return
        wrap = max(180, min(280, event.width // 5))
        try:
            self.file_label.configure(wraplength=wrap)
        except Exception:
            pass

    def _on_close(self) -> None:
        if self.proc and self.proc.poll() is None:
            if not messagebox.askyesno("Sair", "A automação ainda está em execução. Deseja encerrar?"):
                return
            try:
                self.proc.terminate()
            except Exception:
                pass
        self.destroy()


def parse_worker_args(argv: list[str]) -> tuple[int, str] | None:
    if "--worker" not in argv:
        return None
    opcao = None
    planilha = None
    i = 0
    while i < len(argv):
        if argv[i] == "--opcao" and i + 1 < len(argv):
            opcao = int(argv[i + 1])
            i += 2
            continue
        if argv[i] == "--planilha" and i + 1 < len(argv):
            planilha = argv[i + 1]
            i += 2
            continue
        i += 1
    if opcao is None or not planilha:
        print("Uso: gui.py --worker --opcao 1|2 --planilha CAMINHO")
        sys.exit(2)
    return opcao, planilha


def main() -> None:
    worker = parse_worker_args(sys.argv)
    if worker:
        opcao, planilha = worker
        run_worker(opcao, planilha)
        return
    enable_dpi_awareness()
    app = AutomacaoCOBGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
