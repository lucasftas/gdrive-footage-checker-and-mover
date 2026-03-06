import os
import ctypes
import ctypes.wintypes
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path
from datetime import datetime
from collections import defaultdict


def _send_to_trash(paths) -> bool:
    """
    Envia arquivo(s) para a Lixeira do Windows via SHFileOperationW.
    'paths' pode ser um Path/str único ou uma lista deles.
    Usa uma única chamada batch — o Windows exibe seu próprio dialog de progresso.
    """
    if isinstance(paths, (str, Path)):
        paths = [paths]
    if not paths:
        return True

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [
            ("hwnd",                  ctypes.wintypes.HWND),
            ("wFunc",                 ctypes.wintypes.UINT),
            ("pFrom",                 ctypes.c_void_p),   # ponteiro bruto p/ suportar multi-null
            ("pTo",                   ctypes.c_void_p),
            ("fFlags",                ctypes.wintypes.WORD),
            ("fAnyOperationsAborted", ctypes.wintypes.BOOL),
            ("hNameMappings",         ctypes.c_void_p),
            ("lpszProgressTitle",     ctypes.c_void_p),
        ]

    FO_DELETE     = 0x0003
    FOF_ALLOWUNDO = 0x0040   # Envia pra lixeira (recuperavel)
    FOF_NOCONFIRM = 0x0010   # Sem "tem certeza?" do Windows
    # FOF_SILENT  = 0x0004   # REMOVIDO — Windows mostra o dialog de progresso nativo

    # pFrom: caminhos separados por \0, terminados com \0\0
    from_str = '\0'.join(str(p) for p in paths) + '\0\0'
    from_buf  = ctypes.create_unicode_buffer(from_str, len(from_str))

    op        = SHFILEOPSTRUCTW()
    op.wFunc  = FO_DELETE
    op.pFrom  = ctypes.addressof(from_buf)
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRM

    return ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op)) == 0


def _try_move(src_str: str, dst_str: str,
             max_retries: int = 3, retry_delay: float = 2.0):
    """
    Tenta mover um arquivo com retry automático.
    - Erro 17 (cross-device): usa shutil como fallback
    - Erro 32 (sharing violation) ou 5 (access denied): aguarda e tenta novamente
    Retorna (True, info_str | None) em sucesso ou (False, erro_str) em falha.
    """
    import shutil, time
    RETRY_ON = {32, 5}   # sharing violation, access denied
    err_code, err_msg = 0, ""

    for attempt in range(1, max_retries + 1):
        if ctypes.windll.kernel32.MoveFileW(src_str, dst_str):
            return True, None

        err_code = ctypes.windll.kernel32.GetLastError()
        err_msg  = ctypes.FormatError(err_code)

        if err_code == 17:   # ERROR_NOT_SAME_DEVICE — copia e deleta
            try:
                shutil.move(src_str, dst_str)
                return True, "cross-device"
            except Exception as ex:
                return False, f"[shutil fallback] {ex}"

        if err_code in RETRY_ON and attempt < max_retries:
            time.sleep(retry_delay)
            continue           # próxima tentativa

        break  # erro não-recuóperável ou esgotadas as tentativas

    suffix = f" (após {max_retries} tentativas)" if err_code in RETRY_ON else ""
    return False, f"[{err_code}] {err_msg}{suffix}"

MONTHS_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

# Extensões auxiliares geradas pela Sony FX3 (não vídeo/áudio)
FX3_CLEANUP_EXTENSIONS = {
    '.xml',   # Metadados de clip
    '.bdm',   # Disc/bin management
    '.bin',   # Arquivos de sistema
    '.smi',   # Sub-clip info
    '.cpi',   # Clip info
    '.bns',   # Bin summary
    '.idx',   # Índice
    '.sif',   # System info
    '.mui',   # Menu UI data
    '.ppn',   # Planning metadata
    '.mpf',   # Management plan
    '.cup',   # Cue-up
    '.ppl',   # Playlist
    '.edl',   # Edit decision list
}

# Extensões de foto suportadas (A7IV)
PHOTO_EXTENSIONS = {'.arw', '.jpg', '.jpeg', '.tiff', '.heic', '.png'}


def _read_exif_datetime(path_obj) -> str | None:
    """
    Lê o campo DateTimeOriginal do EXIF de um JPEG sem nenhuma dependência.
    Lê apenas os primeiros 64 KB do arquivo — muito rápido.
    Retorna string 'YYYY:MM:DD HH:MM:SS' ou None.
    """
    if path_obj.suffix.lower() not in ('.jpg', '.jpeg'):
        return None
    try:
        import struct
        with open(path_obj, 'rb') as f:
            data = f.read(65536)

        pos = data.find(b'Exif\x00\x00')
        if pos == -1:
            return None

        tiff = pos + 6
        bo   = data[tiff:tiff + 2]
        end  = '<' if bo == b'II' else '>' if bo == b'MM' else None
        if not end:
            return None

        ifd0 = tiff + struct.unpack(end + 'I', data[tiff + 4:tiff + 8])[0]

        def find_tag(ifd_pos, tag_id, depth=0):
            if depth > 3 or ifd_pos + 2 > len(data):
                return None
            n = struct.unpack(end + 'H', data[ifd_pos:ifd_pos + 2])[0]
            for i in range(min(n, 128)):
                ep  = ifd_pos + 2 + i * 12
                if ep + 12 > len(data):
                    break
                tag = struct.unpack(end + 'H', data[ep:ep + 2])[0]
                if tag == tag_id:
                    off = struct.unpack(end + 'I', data[ep + 8:ep + 12])[0]
                    return data[tiff + off:tiff + off + 19].decode('ascii', errors='ignore').strip()
                elif tag == 0x8769:  # ExifIFD pointer
                    off = struct.unpack(end + 'I', data[ep + 8:ep + 12])[0]
                    r   = find_tag(tiff + off, tag_id, depth + 1)
                    if r:
                        return r
            return None

        return find_tag(ifd0, 0x9003)  # DateTimeOriginal
    except Exception:
        return None


class DriveOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Drive Organizer - Limpeza de Duplicatas MP4")
        self.root.geometry("960x680")

        self.local_var  = tk.StringVar(value=r"E:\ConferirComDrive")
        self.drive_var  = tk.StringVar(value=r"Z:\Shared drives\FA 103 - GRAVAÇÕES BRUTAS 2025")
        self.mirror_var = tk.StringVar(value=r"E:\OK")

        self.modos = [
            "1. Nome exato",
            "2. Nome e tamanho exato",
            "3. Nome, tamanho e data",
            "4. Tamanho exato (Ignorar nome - Ideal p/ Sony A6400)"
        ]
        self.mode_var = tk.StringVar(value=self.modos[1])

        self.pending_moves = []
        self._stop_event = threading.Event()   # Sinaliza interrupção da análise

        # Janela de log (Toplevel) — criada sob demanda
        self._log_win   = None
        self._log_text  = None

        self.create_widgets()

    # ================================================================
    # UI principal
    # ================================================================

    def create_widgets(self):
        # --- Diretórios ---
        tk.Label(self.root, text="Pasta Local (A Verificar):", font=("Arial", 10, "bold")).pack(anchor="w", padx=15, pady=(12, 0))
        frame_local = tk.Frame(self.root)
        frame_local.pack(fill="x", padx=15)
        tk.Entry(frame_local, textvariable=self.local_var, font=("Arial", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(frame_local, text="Procurar", command=lambda: self.select_folder(self.local_var)).pack(side="right", padx=(10, 0))

        tk.Label(self.root, text="Referência no Drive (O que já tem lá):", font=("Arial", 10, "bold")).pack(anchor="w", padx=15, pady=(6, 0))
        frame_drive = tk.Frame(self.root)
        frame_drive.pack(fill="x", padx=15)
        tk.Entry(frame_drive, textvariable=self.drive_var, font=("Arial", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(frame_drive, text="Procurar", command=lambda: self.select_folder(self.drive_var)).pack(side="right", padx=(10, 0))

        tk.Label(self.root, text="Pasta Destino (Para onde mover):", font=("Arial", 10, "bold")).pack(anchor="w", padx=15, pady=(6, 0))
        frame_mirror = tk.Frame(self.root)
        frame_mirror.pack(fill="x", padx=15)
        tk.Entry(frame_mirror, textvariable=self.mirror_var, font=("Arial", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(frame_mirror, text="Procurar", command=lambda: self.select_folder(self.mirror_var)).pack(side="right", padx=(10, 0))

        # --- Modo ---
        frame_options = tk.Frame(self.root)
        frame_options.pack(fill="x", padx=15, pady=(10, 0))
        tk.Label(frame_options, text="Modo de Comparação:", font=("Arial", 10, "bold")).pack(side="left")
        ttk.Combobox(frame_options, textvariable=self.mode_var, values=self.modos, state="readonly", width=50).pack(side="left", padx=10)

        # --- Botões ---
        frame_actions = tk.Frame(self.root)
        frame_actions.pack(fill="x", padx=15, pady=10)

        tk.Button(frame_actions, text="1. Analisar",
                  command=self.run_analysis,
                  width=15, font=("Arial", 10, "bold")).pack(side="left", padx=(0, 8))

        tk.Button(frame_actions, text="2. Mover Arquivos",
                  command=self.run_mirror,
                  width=15, bg="#28a745", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=(0, 8))

        self.btn_stop = tk.Button(frame_actions, text="⛔ Interromper",
                  command=self.stop_analysis,
                  width=14, bg="#dc3545", fg="white", font=("Arial", 10, "bold"),
                  state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 8))

        tk.Button(frame_actions, text="📋 Ver Log",
                  command=self.open_log_window,
                  width=10, font=("Arial", 10)).pack(side="left", padx=(0, 8))

        tk.Button(frame_actions, text="🧹 Limpar FX3",
                  command=self.open_cleanup_window,
                  width=13, bg="#6f42c1", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=(0, 8))

        tk.Button(frame_actions, text="📷 Fotos A7IV",
                  command=self.open_photo_compare_window,
                  width=13, bg="#0077b6", fg="white", font=("Arial", 10, "bold")).pack(side="left")

        # --- Status ---
        self.status_label = tk.Label(
            self.root, text="Aguardando ação...",
            fg="#333333", bg="#e9ecef", anchor="w", justify="left",
            relief="flat", font=("Arial", 10), padx=10
        )
        self.status_label.pack(fill="x", padx=15, pady=(0, 6), ipady=5)

        # --- Tabela de Resultados ---
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        tk.Label(tree_frame,
                 text="Arquivos com correspondência no Drive (serão movidos):",
                 font=("Arial", 9, "bold"), fg="#555555").pack(anchor="w")

        tbl_frame = tk.Frame(tree_frame)
        tbl_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tbl_frame, columns=("local", "dest"), show="headings", selectmode="none")
        self.tree.heading("local", text="Caminho Original (Local)")
        self.tree.heading("dest",  text="Novo Destino (Estrutura do Drive)")
        self.tree.column("local", width=440, anchor="w")
        self.tree.column("dest",  width=440, anchor="w")

        scroll_y = ttk.Scrollbar(tbl_frame, orient="vertical",   command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tbl_frame, orient="horizontal",  command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side="right",  fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

    # ================================================================
    # Janela de Log (Toplevel independente)
    # ================================================================

    def open_log_window(self):
        """Abre (ou traz para frente) a janela de log."""
        if self._log_win and tk.Toplevel.winfo_exists(self._log_win):
            self._log_win.lift()
            self._log_win.focus_force()
            return

        # Posiciona ao lado direito da janela principal
        self.root.update_idletasks()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        main_w = self.root.winfo_width()

        win = tk.Toplevel(self.root)
        win.title("Log de Análise")
        win.geometry(f"700x680+{main_x + main_w + 8}+{main_y}")
        win.resizable(True, True)
        win.protocol("WM_DELETE_WINDOW", win.destroy)

        # Toolbar da janela de log
        toolbar = tk.Frame(win, bg="#f0f0f0", bd=1, relief="solid")
        toolbar.pack(fill="x", padx=8, pady=(8, 0))

        tk.Label(toolbar, text="Log de Análise", font=("Arial", 10, "bold"),
                 bg="#f0f0f0").pack(side="left", padx=8, pady=4)

        tk.Button(toolbar, text="🗑 Limpar", font=("Arial", 9),
                  command=self._clear_log_ui, relief="flat", cursor="hand2").pack(side="right", padx=4, pady=2)

        tk.Button(toolbar, text="⛶ Maximizar", font=("Arial", 9),
                  command=lambda: win.state("zoomed"), relief="flat", cursor="hand2").pack(side="right", padx=4, pady=2)

        # Área de texto
        text_frame = tk.Frame(win)
        text_frame.pack(fill="both", expand=True, padx=8, pady=8)

        self._log_text = tk.Text(
            text_frame,
            font=("Courier New", 10),
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white",
            relief="flat", bd=0,
            wrap="none",
            state="disabled"
        )

        sy = ttk.Scrollbar(text_frame, orient="vertical",   command=self._log_text.yview)
        sx = ttk.Scrollbar(text_frame, orient="horizontal",  command=self._log_text.xview)
        self._log_text.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        sy.pack(side="right",  fill="y")
        sx.pack(side="bottom", fill="x")
        self._log_text.pack(side="left", fill="both", expand=True)

        # Tags de cor para o log
        self._log_text.tag_configure("header",   foreground="#569cd6", font=("Courier New", 10, "bold"))
        self._log_text.tag_configure("ok",       foreground="#4ec9b0")
        self._log_text.tag_configure("warn",     foreground="#ce9178")
        self._log_text.tag_configure("error",    foreground="#f44747")
        self._log_text.tag_configure("dim",      foreground="#888888")
        self._log_text.tag_configure("bullet",   foreground="#dcdcaa")

        self._log_win = win

    def _resolve_log_tag(self, text):
        """Decide qual tag de cor usar baseado no conteúdo da linha."""
        if text.startswith("="):
            return "header"
        if text.startswith("ETAPA"):
            return "header"
        if "✓" in text or "sucesso" in text.lower():
            return "ok"
        if "✗" in text or "ERRO" in text or "Falha" in text:
            return "error"
        if "SEM" in text or "sem sincronismo" in text.lower():
            return "warn"
        if text.strip().startswith("•"):
            return "bullet"
        if text.strip() == "":
            return "dim"
        return None

    def log(self, text):
        """Adiciona linha ao log (thread-safe). Abre a janela se ainda não existir."""
        self.root.after(0, self._log_ui, text)

    def _log_ui(self, text):
        # Garante que a janela exista
        if not self._log_win or not tk.Toplevel.winfo_exists(self._log_win):
            self.open_log_window()

        self._log_text.config(state="normal")
        tag = self._resolve_log_tag(text)
        if tag:
            self._log_text.insert(tk.END, text + "\n", tag)
        else:
            self._log_text.insert(tk.END, text + "\n")
        self._log_text.see(tk.END)
        self._log_text.config(state="disabled")

    def clear_log(self):
        self.root.after(0, self._clear_log_ui)

    def _clear_log_ui(self):
        if self._log_text and self._log_win and tk.Toplevel.winfo_exists(self._log_win):
            self._log_text.config(state="normal")
            self._log_text.delete("1.0", tk.END)
            self._log_text.config(state="disabled")

    # ================================================================
    # Helpers de UI
    # ================================================================

    def select_folder(self, var):
        current_path = var.get()
        initial_dir = ""
        if current_path:
            try:
                parent_dir = Path(current_path).parent
                if parent_dir.exists():
                    initial_dir = str(parent_dir)
            except Exception:
                pass
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            var.set(os.path.normpath(folder))

    def update_status(self, message, color="#333333"):
        self.root.after(0, self._update_status_ui, message, color)

    def _update_status_ui(self, message, color):
        self.status_label.config(text=message, fg=color)

    def add_tree_item(self, local_path, dest_path):
        self.root.after(0, lambda: self.tree.insert("", tk.END, values=(local_path, dest_path)))

    def clear_tree(self):
        self.root.after(0, self._clear_tree_ui)

    def _clear_tree_ui(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    # ================================================================
    # Lógica de comparação
    # ================================================================

    def _get_file_key(self, path_obj, mode_index):
        stat  = path_obj.stat()
        name  = path_obj.name.lower()
        size  = stat.st_size
        mtime = int(stat.st_mtime)

        if mode_index == 0:   return name
        elif mode_index == 1: return (name, size)
        elif mode_index == 2: return (name, size, mtime)
        elif mode_index == 3: return size
        return None

    def _mtime_label(self, path_obj):
        dt    = datetime.fromtimestamp(path_obj.stat().st_mtime)
        label = f"{MONTHS_PT[dt.month]} {dt.year}"
        return (dt.year, dt.month), label

    # ================================================================
    # Análise
    # ================================================================

    def stop_analysis(self):
        """Sinaliza a thread de análise para parar."""
        self._stop_event.set()
        self.update_status("Interrompendo análise...", "#dc3545")
        self.log("")
        self.log("⛔ Análise interrompida pelo usuário.")
        self.root.after(0, lambda: self.btn_stop.config(state="disabled"))

    def _set_analysis_running(self, running: bool):
        """Habilita/desabilita o botão Interromper (thread-safe)."""
        state = "normal" if running else "disabled"
        self.root.after(0, lambda: self.btn_stop.config(state=state))

    def run_analysis(self):
        local_dir  = self.local_var.get()
        drive_dir  = self.drive_var.get()
        mode_index = self.modos.index(self.mode_var.get())

        self._stop_event.clear()        # Reseta qualquer interrupção anterior
        self.pending_moves.clear()
        self.clear_tree()
        self.clear_log()
        self.open_log_window()          # Garante que a janela esteja aberta
        self._set_analysis_running(True)
        self.update_status("Indexando arquivos do Drive. Isso pode levar alguns segundos...", "#0078D7")

        threading.Thread(
            target=self._analyze_process,
            args=(local_dir, drive_dir, mode_index),
            daemon=True
        ).start()

    def _analyze_process(self, local_dir, drive_dir, mode_index):
        try:
            local_path  = Path(local_dir)
            drive_path  = Path(drive_dir)
            mirror_path = Path(self.mirror_var.get())

            # ---------------------------------------------------
            # ETAPA 1 — Inventário da Pasta Local (.mp4)
            # ---------------------------------------------------
            self.log("=" * 62)
            self.log("ETAPA 1 — Inventário da Pasta Local")
            self.log("=" * 62)
            self.update_status("Lendo pasta local...", "#0078D7")

            local_files = []
            for item in local_path.rglob('*'):
                if self._stop_event.is_set():
                    break
                if item.is_file() and item.suffix.lower() == '.mp4':
                    local_files.append(item)

            if self._stop_event.is_set():
                self._set_analysis_running(False)
                return

            by_month = defaultdict(list)
            for f in local_files:
                key, label = self._mtime_label(f)
                by_month[key].append((f, label))

            self.log(f"Total de .mp4 encontrados na pasta local: {len(local_files)}")
            self.log("")
            self.log("Distribuição por data de modificação:")
            for (year, month) in sorted(by_month.keys(), reverse=True):
                entries = by_month[(year, month)]
                label   = entries[0][1]
                self.log(f"  • {len(entries):>4} arquivo(s) em {label}")
            self.log("")

            # ---------------------------------------------------
            # ETAPA 2 — Indexar o Drive
            # ---------------------------------------------------
            self.log("=" * 62)
            self.log("ETAPA 2 — Indexando .mp4 no Drive")
            self.log("=" * 62)
            self.update_status("Indexando Drive...", "#0078D7")

            drive_index = {}
            drive_skipped = 0
            for item in drive_path.rglob('*'):
                if self._stop_event.is_set():
                    break
                if item.is_file() and item.suffix.lower() == '.mp4':
                    self.update_status(f"Indexando Drive: {item.name}", "#0078D7")
                    try:
                        key = self._get_file_key(item, mode_index)
                        drive_index.setdefault(key, []).append(item)
                    except Exception as e:
                        drive_skipped += 1
                        self.log(f"  ⚠ Pulado (Drive): {item.name} — {e}")

            if self._stop_event.is_set():
                self._set_analysis_running(False)
                return

            self.log(f"Total de .mp4 indexados no Drive: {len(drive_index)}" +
                     (f"  ({drive_skipped} pulado(s) por erro)" if drive_skipped else ""))
            self.log("")

            # ---------------------------------------------------
            # ETAPA 3 — Comparação
            # ---------------------------------------------------
            self.log("=" * 62)
            self.log("ETAPA 3 — Comparando .mp4 locais com o Drive")
            self.log("=" * 62)
            self.update_status("Comparando arquivos locais com o Drive...", "#FF8C00")

            matched_files   = []
            unmatched_files = []
            local_skipped   = 0

            for item in local_files:
                if self._stop_event.is_set():
                    break
                self.update_status(f"Verificando: {item.name}", "#FF8C00")
                try:
                    key = self._get_file_key(item, mode_index)
                except Exception as e:
                    local_skipped += 1
                    self.log(f"  ⚠ Pulado (Local): {item.name} — {e}")
                    continue

                if key in drive_index:
                    matched_drive_item = drive_index[key][0]
                    rel_drive_path     = matched_drive_item.relative_to(drive_path)
                    target_item        = mirror_path / rel_drive_path
                    self.pending_moves.append((item, target_item))
                    self.add_tree_item(str(item), str(target_item))
                    matched_files.append(item)
                else:
                    unmatched_files.append(item)

            if local_skipped:
                self.log(f"  ⚠ Arquivos pulados por erro de acesso: {local_skipped}")

            self.log(f"  ✓ Com correspondência no Drive:  {len(matched_files):>4} → prontos para mover")
            self.log(f"  ✗ SEM correspondência no Drive:  {len(unmatched_files):>4}")
            self.log("")

            # ---------------------------------------------------
            # ETAPA 4 — Relatório de sem sincronismo
            # ---------------------------------------------------
            self.log("=" * 62)
            self.log("ETAPA 4 — Arquivos SEM Sincronismo com o Drive")
            self.log("=" * 62)

            if unmatched_files:
                unmatched_by_month = defaultdict(list)
                for f in unmatched_files:
                    key, label = self._mtime_label(f)
                    unmatched_by_month[key].append((f, label))

                for (year, month) in sorted(unmatched_by_month.keys(), reverse=True):
                    entries = unmatched_by_month[(year, month)]
                    label   = entries[0][1]
                    self.log(f"  • {len(entries):>4} arquivo(s) sem sincronismo em {label}")
            else:
                self.log("  ✓ Todos os .mp4 locais foram encontrados no Drive!")

            self.log("")
            self.log("=" * 62)
            self._set_analysis_running(False)
            self.update_status(
                f"Análise concluída: {len(matched_files)} prontos para mover  |  "
                f"{len(unmatched_files)} sem correspondência no Drive.",
                "black"
            )

        except Exception as e:
            self.update_status(f"Erro na análise: {str(e)}", "red")
            self.log(f"ERRO: {str(e)}")
            self._set_analysis_running(False)

    # ================================================================
    # Comparação de Fotos — Sony A7IV
    # ================================================================

    def open_photo_compare_window(self):
        """Janela de comparação de fotos A7IV: ARW por tamanho, JPG por tamanho+EXIF."""
        win = tk.Toplevel(self.root)
        win.title("📷 Comparação de Fotos — Sony A7IV")
        win.resizable(True, True)

        self.root.update_idletasks()
        mx = self.root.winfo_x()
        my = self.root.winfo_y()
        mw = self.root.winfo_width()
        win.geometry(f"960x720+{mx + mw + 8}+{my}")

        # Estado interno da janela
        _pending   = []
        _stop_evt  = threading.Event()

        # ── Cabeçalho ──────────────────────────────────────────────
        hdr = tk.Frame(win, bg="#0077b6", pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📷  Comparação de Fotos — Sony A7IV",
                 font=("Arial", 11, "bold"), fg="white", bg="#0077b6").pack(padx=14, anchor="w")
        tk.Label(hdr, text="ARW → tamanho exato  |  JPG/JPEG → tamanho + Data EXIF  |  Nomes ignorados",
                 font=("Arial", 9), fg="#90e0ef", bg="#0077b6").pack(padx=14, anchor="w")

        # ── Pastas (herda valores do app principal) ─────────────────
        def _row(parent, label, var):
            tk.Label(parent, text=label, font=("Arial", 10, "bold")).pack(anchor="w", padx=12, pady=(6, 0))
            f = tk.Frame(parent)
            f.pack(fill="x", padx=12)
            tk.Entry(f, textvariable=var, font=("Arial", 10)).pack(side="left", fill="x", expand=True)
            tk.Button(f, text="Procurar",
                      command=lambda v=var: self.select_folder(v)).pack(side="right", padx=(8, 0))

        p_local_var  = tk.StringVar(value=self.local_var.get())
        p_drive_var  = tk.StringVar(value=self.drive_var.get())
        p_mirror_var = tk.StringVar(value=self.mirror_var.get())

        _row(win, "Pasta Local (fotos a verificar):",           p_local_var)
        _row(win, "Referência no Drive (fotos já enviadas):",   p_drive_var)
        _row(win, "Pasta Destino (para onde mover matches):",   p_mirror_var)

        # ── Extensões ──────────────────────────────────────────────
        ext_frame = tk.Frame(win)
        ext_frame.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(ext_frame, text="Extensões:", font=("Arial", 10, "bold")).pack(side="left")
        ext_vars = {}
        for ext in ['.ARW', '.JPG', '.JPEG', '.TIFF', '.HEIC', '.PNG']:
            v = tk.BooleanVar(value=ext in ('.ARW', '.JPG', '.JPEG'))
            ext_vars[ext.lower()] = v
            tk.Checkbutton(ext_frame, text=ext, variable=v,
                           font=("Arial", 10)).pack(side="left", padx=6)

        # ── Botões de ação ─────────────────────────────────────────
        ba = tk.Frame(win)
        ba.pack(fill="x", padx=12, pady=8)

        btn_analyze = tk.Button(ba, text="1. Analisar", width=14,
                                font=("Arial", 10, "bold"))
        btn_analyze.pack(side="left", padx=(0, 8))

        btn_move = tk.Button(ba, text="2. Mover Fotos", width=14,
                             bg="#28a745", fg="white", font=("Arial", 10, "bold"),
                             state="disabled")
        btn_move.pack(side="left", padx=(0, 8))

        btn_stop = tk.Button(ba, text="⛔ Interromper", width=14,
                             bg="#dc3545", fg="white", font=("Arial", 10, "bold"),
                             state="disabled")
        btn_stop.pack(side="left")

        # ── Status ─────────────────────────────────────────────────
        status_var = tk.StringVar(value="Aguardando análise...")
        status_lbl = tk.Label(win, textvariable=status_var,
                              font=("Arial", 10), bg="#e9ecef", fg="#333",
                              anchor="w", padx=10, relief="flat")
        status_lbl.pack(fill="x", padx=12, pady=(0, 4), ipady=4)

        # ── Log ────────────────────────────────────────────────────
        log_frame = tk.Frame(win, bd=1, relief="solid")
        log_frame.pack(fill="x", padx=12, pady=(0, 4))
        log_text = tk.Text(log_frame, height=7, font=("Courier New", 9),
                           bg="#1e1e1e", fg="#d4d4d4", wrap="none",
                           state="disabled", relief="flat")
        log_sy = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
        log_sx = ttk.Scrollbar(log_frame, orient="horizontal", command=log_text.xview)
        log_text.configure(yscrollcommand=log_sy.set, xscrollcommand=log_sx.set)
        log_text.tag_configure("hdr",  foreground="#569cd6", font=("Courier New", 9, "bold"))
        log_text.tag_configure("ok",   foreground="#4ec9b0")
        log_text.tag_configure("warn", foreground="#ce9178")
        log_text.tag_configure("err",  foreground="#f44747")
        log_sy.pack(side="right", fill="y")
        log_sx.pack(side="bottom", fill="x")
        log_text.pack(fill="both", expand=False)

        def _log(text, tag=None):
            def _do():
                log_text.config(state="normal")
                if tag:
                    log_text.insert(tk.END, text + "\n", tag)
                else:
                    log_text.insert(tk.END, text + "\n")
                log_text.see(tk.END)
                log_text.config(state="disabled")
            win.after(0, _do)

        def _clear_log():
            log_text.config(state="normal")
            log_text.delete("1.0", tk.END)
            log_text.config(state="disabled")

        # ── Tabela de resultados ───────────────────────────────────
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        tk.Label(tree_frame, text="Fotos com correspondência no Drive (serão movidas):",
                 font=("Arial", 9, "bold"), fg="#555").pack(anchor="w")

        tbl = tk.Frame(tree_frame)
        tbl.pack(fill="both", expand=True)
        tree = ttk.Treeview(tbl, columns=("local", "dest"), show="headings", selectmode="none")
        tree.heading("local", text="Caminho Original (Local)")
        tree.heading("dest",  text="Destino (Estrutura do Drive)")
        tree.column("local", width=450, anchor="w")
        tree.column("dest",  width=450, anchor="w")
        sy2 = ttk.Scrollbar(tbl, orient="vertical",   command=tree.yview)
        sx2 = ttk.Scrollbar(tbl, orient="horizontal",  command=tree.xview)
        tree.configure(yscrollcommand=sy2.set, xscrollcommand=sx2.set)
        sy2.pack(side="right", fill="y")
        sx2.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)

        # ── Funções de key por extensão ────────────────────────────
        def _photo_key(path_obj):
            size = path_obj.stat().st_size
            ext  = path_obj.suffix.lower()
            if ext in ('.jpg', '.jpeg'):
                exif_dt = _read_exif_datetime(path_obj)
                return (size, exif_dt) if exif_dt else (size,)
            # ARW e demais: tamanho apenas
            return (size,)

        # ── Análise ────────────────────────────────────────────────
        def _run_analysis():
            local_dir  = p_local_var.get()
            drive_dir  = p_drive_var.get()
            mirror_dir = p_mirror_var.get()

            selected_exts = {ext for ext, var in ext_vars.items() if var.get()}
            if not selected_exts:
                status_var.set("Selecione ao menos uma extensão.")
                return

            _pending.clear()
            _stop_evt.clear()
            for item in tree.get_children():
                tree.delete(item)
            win.after(0, _clear_log)

            btn_analyze.config(state="disabled")
            btn_move.config(state="disabled")
            btn_stop.config(state="normal")
            status_var.set("Iniciando análise...")

            def _thread():
                try:
                    local_path  = Path(local_dir)
                    drive_path  = Path(drive_dir)
                    mirror_path = Path(mirror_dir)

                    # ETAPA 1 — inventário local
                    _log("=" * 62, "hdr")
                    _log("ETAPA 1 — Inventário local", "hdr")
                    _log("=" * 62, "hdr")
                    win.after(0, lambda: status_var.set("Lendo pasta local..."))

                    local_files = []
                    for item in local_path.rglob('*'):
                        if _stop_evt.is_set(): break
                        if item.is_file() and item.suffix.lower() in selected_exts:
                            local_files.append(item)

                    if _stop_evt.is_set():
                        _log("⛔ Interrompido.", "warn")
                        win.after(0, _finish_buttons)
                        return

                    # Mini resumo por mês/ano
                    by_month = defaultdict(list)
                    for f in local_files:
                        key, label = self._mtime_label(f)
                        by_month[key].append((f, label))

                    _log(f"Total de fotos na pasta local: {len(local_files)}")
                    _log("")
                    _log("Distribuição por data de modificação:")
                    for (yr, mo) in sorted(by_month.keys(), reverse=True):
                        ents  = by_month[(yr, mo)]
                        label = ents[0][1]
                        _log(f"  • {len(ents):>4} foto(s) em {label}", "ok")
                    _log("")

                    # ETAPA 2 — indexar Drive
                    _log("=" * 62, "hdr")
                    _log("ETAPA 2 — Indexando fotos no Drive", "hdr")
                    _log("=" * 62, "hdr")
                    win.after(0, lambda: status_var.set("Indexando Drive..."))

                    drive_index = {}
                    for item in drive_path.rglob('*'):
                        if _stop_evt.is_set(): break
                        if item.is_file() and item.suffix.lower() in selected_exts:
                            win.after(0, lambda n=item.name: status_var.set(f"Drive: {n}"))
                            key = _photo_key(item)
                            drive_index.setdefault(key, []).append(item)

                    if _stop_evt.is_set():
                        _log("⛔ Interrompido.", "warn")
                        win.after(0, _finish_buttons)
                        return

                    _log(f"Total de fotos indexadas no Drive: {len(drive_index)}")
                    _log("")

                    # ETAPA 3 — comparação
                    _log("=" * 62, "hdr")
                    _log("ETAPA 3 — Comparando com o Drive", "hdr")
                    _log("=" * 62, "hdr")
                    win.after(0, lambda: status_var.set("Comparando..."))

                    matched   = []
                    unmatched = []

                    for item in local_files:
                        if _stop_evt.is_set(): break
                        win.after(0, lambda n=item.name: status_var.set(f"Verificando: {n}"))
                        key = _photo_key(item)
                        if key in drive_index:
                            matched_drv    = drive_index[key][0]
                            rel_drive_path = matched_drv.relative_to(drive_path)
                            target         = mirror_path / rel_drive_path
                            _pending.append((item, target))
                            tree.after(0, lambda l=str(item), d=str(target):
                                       tree.insert("", tk.END, values=(l, d)))
                            matched.append(item)
                        else:
                            unmatched.append(item)

                    _log(f"  ✓ Com correspondência:    {len(matched):>4} → prontos para mover", "ok")
                    _log(f"  ✗ SEM correspondência:    {len(unmatched):>4}", "warn")
                    _log("")

                    # ETAPA 4 — sem sincronismo
                    _log("=" * 62, "hdr")
                    _log("ETAPA 4 — Fotos SEM Sincronismo", "hdr")
                    _log("=" * 62, "hdr")
                    if unmatched:
                        um_by_month = defaultdict(list)
                        for f in unmatched:
                            key, label = self._mtime_label(f)
                            um_by_month[key].append((f, label))
                        for (yr, mo) in sorted(um_by_month.keys(), reverse=True):
                            ents  = um_by_month[(yr, mo)]
                            label = ents[0][1]
                            _log(f"  • {len(ents):>4} foto(s) sem sincronismo em {label}", "warn")
                    else:
                        _log("  ✓ Todas as fotos foram encontradas no Drive!", "ok")

                    _log("")
                    _log("=" * 62, "hdr")

                    win.after(0, lambda: status_var.set(
                        f"Análise concluída: {len(matched)} prontos para mover  |  "
                        f"{len(unmatched)} sem correspondência."
                    ))
                    win.after(0, lambda: btn_move.config(
                        state="normal" if _pending else "disabled"))

                except Exception as e:
                    _log(f"ERRO: {e}", "err")
                    win.after(0, lambda: status_var.set(f"Erro: {e}"))
                finally:
                    win.after(0, _finish_buttons)

            threading.Thread(target=_thread, daemon=True).start()

        def _finish_buttons():
            btn_analyze.config(state="normal")
            btn_stop.config(state="disabled")

        def _stop():
            _stop_evt.set()
            status_var.set("Interrompendo...")
            btn_stop.config(state="disabled")
            _log("⛔ Análise interrompida pelo usuário.", "warn")

        def _run_move():
            if not _pending:
                status_var.set("Nenhum arquivo na fila.")
                return
            btn_move.config(state="disabled")
            btn_analyze.config(state="disabled")
            status_var.set(f"Movendo {len(_pending)} foto(s)...")

            def _thread():
                moved  = 0
                failed = 0
                total  = len(_pending)
                for i, (src, dst) in enumerate(list(_pending), 1):
                    try:
                        win.after(0, lambda n=src.name, idx=i:
                                  status_var.set(f"[{idx}/{total}] Movendo: {n}"))
                        dst.parent.mkdir(parents=True, exist_ok=True)
                        s, d = str(src), str(dst)
                        if len(s) >= 256 and not s.startswith("\\\\?\\"):
                            s = "\\\\?\\" + s
                        if len(d) >= 256 and not d.startswith("\\\\?\\"):
                            d = "\\\\?\\" + d
                        ok, info = _try_move(s, d)
                        if ok:
                            moved += 1
                            if info == "cross-device":
                                _log(f"  ⚠ Copiado (drives diferentes): {src.name}", "warn")
                        else:
                            _log(f"  ✗ FALHA: {src.name}", "err")
                            _log(f"      Motivo: {info}", "err")
                            _log(f"      SRC: {s}", "err")
                            _log(f"      DST: {d}", "err")
                            failed += 1
                    except Exception as ex:
                        _log(f"  ✗ Exceção: {src.name} — {ex}", "err")
                        failed += 1
                _pending.clear()
                msg = f"✓ {moved} foto(s) movida(s)." + (f"  ✗ {failed} falha(s)." if failed else "")
                win.after(0, lambda: status_var.set(msg))
                win.after(0, lambda: btn_analyze.config(state="normal"))
                _log(f"\n{msg}", "ok" if not failed else "warn")

            threading.Thread(target=_thread, daemon=True).start()

        btn_analyze.config(command=_run_analysis)
        btn_stop.config(command=_stop)
        btn_move.config(command=_run_move)

    # ================================================================
    # Limpeza FX3 — arquivos auxiliares
    # ================================================================

    def open_cleanup_window(self):
        """Abre janela para listar e excluir arquivos auxiliares da FX3."""
        local_dir = self.local_var.get()
        if not local_dir or not Path(local_dir).exists():
            self.update_status("Configure a Pasta Local antes de usar a limpeza FX3.", "red")
            return

        win = tk.Toplevel(self.root)
        win.title("🧹 Limpar Arquivos Auxiliares FX3")
        win.resizable(True, True)

        # Posiciona ao lado da janela principal
        self.root.update_idletasks()
        mx = self.root.winfo_x()
        my = self.root.winfo_y()
        mw = self.root.winfo_width()
        win.geometry(f"620x580+{mx + mw + 8}+{my}")

        # --- Cabeçalho ---
        hdr = tk.Frame(win, bg="#6f42c1", pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🧹  Limpeza de Arquivos Auxiliares — Sony FX3",
                 font=("Arial", 11, "bold"), fg="white", bg="#6f42c1").pack(padx=14, anchor="w")
        tk.Label(hdr, text=f"Pasta: {local_dir}",
                 font=("Arial", 9), fg="#ddccff", bg="#6f42c1", wraplength=580, justify="left").pack(padx=14, anchor="w")

        # --- Info extensões ---
        info_frame = tk.Frame(win, bg="#f3f0ff", bd=1, relief="solid")
        info_frame.pack(fill="x", padx=10, pady=(8, 0))
        exts_str = "  ".join(sorted(e.upper() for e in FX3_CLEANUP_EXTENSIONS))
        tk.Label(info_frame, text=f"Extensões monitoradas:  {exts_str}",
                 font=("Courier New", 8), fg="#555", bg="#f3f0ff", anchor="w", padx=8, pady=4).pack(fill="x")

        # --- Toolbar de seleção ---
        sel_bar = tk.Frame(win)
        sel_bar.pack(fill="x", padx=10, pady=(8, 2))
        tk.Label(sel_bar, text="Selecione as pastas/tipos para excluir:",
                 font=("Arial", 9, "bold")).pack(side="left")
        tk.Button(sel_bar, text="Marcar Todos",   font=("Arial", 8),
                  command=lambda: _toggle_all(True),  relief="flat", cursor="hand2").pack(side="right", padx=(4, 0))
        tk.Button(sel_bar, text="Desmarcar Todos", font=("Arial", 8),
                  command=lambda: _toggle_all(False), relief="flat", cursor="hand2").pack(side="right")

        # --- Lista com scroll ---
        list_outer = tk.Frame(win, bd=1, relief="solid")
        list_outer.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        canvas   = tk.Canvas(list_outer, bg="white", highlightthickness=0)
        scroll_y = ttk.Scrollbar(list_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        list_frame = tk.Frame(canvas, bg="white")
        canvas_win = canvas.create_window((0, 0), window=list_frame, anchor="nw")

        def _on_frame_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_configure(e):
            canvas.itemconfig(canvas_win, width=e.width)
        list_frame.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # Scroll com roda do mouse
        def _on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        win.protocol("WM_DELETE_WINDOW", lambda: [canvas.unbind_all("<MouseWheel>"), win.destroy()])

        # --- Status inferior ---
        status_frm = tk.Frame(win, bg="#e9ecef")
        status_frm.pack(fill="x", padx=10, pady=(0, 4))
        status_var = tk.StringVar(value="Aguardando varredura...")
        tk.Label(status_frm, textvariable=status_var, font=("Arial", 9),
                 bg="#e9ecef", anchor="w", padx=6).pack(fill="x", ipady=3)

        # --- Botão excluir ---
        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        btn_scan = tk.Button(btn_frame, text="🔍 Escanear Pasta",
                             font=("Arial", 10, "bold"), width=18)
        btn_scan.pack(side="left", padx=(0, 8))
        btn_delete = tk.Button(btn_frame, text="🗑 Excluir Selecionados",
                               font=("Arial", 10, "bold"), bg="#dc3545", fg="white",
                               width=20, state="disabled")
        btn_delete.pack(side="left")

        # --- Estado interno ---
        check_vars = {}   # key -> BooleanVar
        groups     = {}   # key -> list[Path]

        def _toggle_all(state):
            for var in check_vars.values():
                var.set(state)

        def _do_scan():
            # Limpa lista
            for w in list_frame.winfo_children():
                w.destroy()
            check_vars.clear()
            groups.clear()
            btn_delete.config(state="disabled")
            status_var.set("Escaneando...")
            win.update_idletasks()

            local_path = Path(local_dir)
            raw_groups = defaultdict(list)

            for item in local_path.rglob('*'):
                if item.is_file() and item.suffix.lower() in FX3_CLEANUP_EXTENSIONS:
                    folder_name = item.parent.name.upper() or item.parent.as_posix()
                    key = (folder_name, item.suffix.upper())
                    raw_groups[key].append(item)

            if not raw_groups:
                tk.Label(list_frame, text="Nenhum arquivo auxiliar FX3 encontrado.",
                         font=("Arial", 10), fg="#888", bg="white").pack(pady=20)
                status_var.set("Nada encontrado.")
                return

            groups.update(raw_groups)

            # Cabeçalho da lista
            hdr_row = tk.Frame(list_frame, bg="#6f42c1")
            hdr_row.pack(fill="x")
            tk.Label(hdr_row, text="  ", bg="#6f42c1", width=3).pack(side="left")
            tk.Label(hdr_row, text="Qtd",  width=5,  font=("Arial", 9, "bold"), fg="white", bg="#6f42c1", anchor="center").pack(side="left")
            tk.Label(hdr_row, text="Extensão", width=10, font=("Arial", 9, "bold"), fg="white", bg="#6f42c1", anchor="w").pack(side="left")
            tk.Label(hdr_row, text="em Pasta", font=("Arial", 9, "bold"), fg="white", bg="#6f42c1", anchor="w").pack(side="left", padx=4)

            # Uma linha por grupo, ordenado por pasta depois extensão
            for i, key in enumerate(sorted(groups.keys())):
                folder_name, ext = key
                count  = len(groups[key])
                bg_row = "#ffffff" if i % 2 == 0 else "#f8f4ff"

                row = tk.Frame(list_frame, bg=bg_row)
                row.pack(fill="x")

                var = tk.BooleanVar(value=False)
                check_vars[key] = var

                cb = tk.Checkbutton(row, variable=var, bg=bg_row,
                                    activebackground=bg_row, cursor="hand2")
                cb.pack(side="left", padx=(4, 0))

                tk.Label(row, text=f"{count:>4}", width=5, font=("Courier New", 10, "bold"),
                         fg="#dc3545", bg=bg_row, anchor="center").pack(side="left")
                tk.Label(row, text=ext, width=10, font=("Courier New", 10),
                         fg="#6f42c1", bg=bg_row, anchor="w").pack(side="left")
                tk.Label(row, text=f"em  {folder_name}", font=("Arial", 10),
                         fg="#333", bg=bg_row, anchor="w").pack(side="left", padx=4)

            total_files = sum(len(v) for v in groups.values())
            status_var.set(f"{len(groups)} grupo(s) encontrado(s)  —  {total_files} arquivo(s) no total")
            btn_delete.config(state="normal")

        def _do_delete():
            selected_keys = [k for k, v in check_vars.items() if v.get()]
            if not selected_keys:
                status_var.set("Nenhum grupo selecionado.")
                return

            total_sel = sum(len(groups[k]) for k in selected_keys)

            # Confirmação
            import tkinter.messagebox as mb
            ok = mb.askyesno(
                "Confirmar Envio para Lixeira",
                f"Enviar {total_sel} arquivo(s) de {len(selected_keys)} grupo(s) para a Lixeira?\n\n"
                + "\n".join(f"  {len(groups[k]):>4}  {k[0]} / {k[1]}" for k in selected_keys)
                + "\n\nVocê poderá recuperar os arquivos pela Lixeira do Windows.",
                parent=win
            )
            if not ok:
                return

            # Coleta todos os arquivos dos grupos selecionados
            all_files = []
            for key in selected_keys:
                all_files.extend(groups[key])

            # Desabilita botões durante a operação
            btn_delete.config(state="disabled")
            btn_scan.config(state="disabled")
            status_var.set(f"Enviando {len(all_files)} arquivo(s) para a Lixeira...")

            def _run_in_thread():
                try:
                    success = _send_to_trash(all_files)  # UMA chamada batch — Windows mostra progress
                    if success:
                        msg = f"✓ {len(all_files)} arquivo(s) enviado(s) para a Lixeira."
                    else:
                        msg = f"✗ Falha ao enviar para a Lixeira (verifique permissões)."
                except Exception as ex:
                    msg = f"Erro: {ex}"
                finally:
                    # Atualiza UI e re-escaneia (sempre na thread principal)
                    win.after(0, lambda: status_var.set(msg))
                    win.after(0, lambda: btn_delete.config(state="normal"))
                    win.after(0, lambda: btn_scan.config(state="normal"))
                    win.after(100, _do_scan)

            threading.Thread(target=_run_in_thread, daemon=True).start()

        btn_scan.config(command=_do_scan)
        btn_delete.config(command=_do_delete)

        # Escaneia automaticamente ao abrir
        win.after(100, _do_scan)

    # ================================================================
    # Mover arquivos
    # ================================================================

    def run_mirror(self):
        if not self.pending_moves:
            self.update_status("Nenhum arquivo na fila. Execute a análise primeiro.", "red")
            return
        self.update_status("Movendo arquivos MP4...", "#28a745")
        threading.Thread(target=self._mirror_process, daemon=True).start()

    def _mirror_process(self):
        moved  = 0
        failed = 0
        try:
            total = len(self.pending_moves)
            for i, (src_path, dest_path) in enumerate(self.pending_moves, 1):
                self.update_status(f"[{i}/{total}] Movendo: {src_path.name}", "#28a745")
                dest_path.parent.mkdir(parents=True, exist_ok=True)

                src_str = str(src_path)
                dst_str = str(dest_path)
                if len(src_str) >= 256 and not src_str.startswith("\\\\?\\"):
                    src_str = "\\\\?\\" + src_str
                if len(dst_str) >= 256 and not dst_str.startswith("\\\\?\\"):
                    dst_str = "\\\\?\\" + dst_str

                ok, info = _try_move(src_str, dst_str)
                if ok:
                    moved += 1
                    if info == "cross-device":
                        self.log(f"  ⚠ Copiado (drives diferentes): {src_path.name}")
                else:
                    self.log(f"  ✗ FALHA: {src_path.name}")
                    self.log(f"      Motivo: {info}")
                    self.log(f"      SRC: {src_str}")
                    self.log(f"      DST: {dst_str}")
                    failed += 1

            self.pending_moves.clear()
            cor = "#28a745" if not failed else "#FF8C00"
            self.update_status(
                f"Concluído! {moved} movido(s)" + (f" | {failed} falha(s)" if failed else "."),
                cor)
            self.log(f"\n✓ {moved} arquivo(s) movido(s)." +
                     (f"  ✗ {failed} falha(s)." if failed else ""))

        except Exception as e:
            self.update_status(f"Erro ao mover: {str(e)}", "red")
            self.log(f"ERRO ao mover: {str(e)}")


if __name__ == '__main__':
    root = tk.Tk()
    app  = DriveOrganizerApp(root)
    root.mainloop()