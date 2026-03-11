import random
import threading
import time
import sys
import json
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Callable, List, Optional, Tuple

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    TK_AVAILABLE = True
    TK_IMPORT_ERROR = None
except ModuleNotFoundError as exc:
    tk = None
    ttk = None
    messagebox = None
    filedialog = None
    TK_AVAILABLE = False
    TK_IMPORT_ERROR = exc

try:
    import pyautogui
except ImportError:
    pyautogui = None

try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    keyboard = None
    KEYBOARD_AVAILABLE = False


Region = Tuple[int, int, int, int]
APP_TITLE = "Macro de Cliques por Região"
DEFAULT_PROFILE_PATH = Path("macro_profile.json")


@dataclass
class Step:
    region: Region
    delay: float

    def to_dict(self) -> dict:
        return {"region": list(self.region), "delay": self.delay}

    @staticmethod
    def from_dict(data: dict) -> "Step":
        region = tuple(data["region"])
        delay = float(data["delay"])
        if len(region) != 4:
            raise ValueError("Region must have 4 coordinates")
        region = normalize_region(*region)
        if not is_valid_region(region):
            raise ValueError("Invalid region size")
        if delay < 0:
            raise ValueError("Delay must be >= 0")
        return Step(region=region, delay=delay)


def parse_delay(value: str) -> float:
    parsed = float(value.strip().replace(",", "."))
    if parsed < 0:
        raise ValueError("Delay must be >= 0")
    return parsed


def normalize_region(x1: int, y1: int, x2: int, y2: int) -> Region:
    left = min(x1, x2)
    top = min(y1, y2)
    right = max(x1, x2)
    bottom = max(y1, y2)
    return left, top, right, bottom


def is_valid_region(region: Region, min_size: int = 3) -> bool:
    x1, y1, x2, y2 = region
    return (x2 - x1) >= min_size and (y2 - y1) >= min_size


def random_point_in_region(region: Region, rng: Optional[random.Random] = None) -> Tuple[int, int]:
    rng = rng or random
    x1, y1, x2, y2 = region
    return rng.randint(x1, x2), rng.randint(y1, y2)


def save_profile(path: Path, steps: List[Step], loops: int, hotkeys_enabled: bool):
    payload = {
        "version": 1,
        "loops": loops,
        "hotkeys_enabled": hotkeys_enabled,
        "steps": [step.to_dict() for step in steps],
    }
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def load_profile(path: Path) -> Tuple[List[Step], int, bool]:
    data = json.loads(path.read_text(encoding="utf-8"))
    loops = int(data.get("loops", 0))
    hotkeys_enabled = bool(data.get("hotkeys_enabled", True))
    raw_steps = data.get("steps", [])
    steps = [Step.from_dict(item) for item in raw_steps]
    if loops < 0:
        raise ValueError("Loops must be >= 0")
    return steps, loops, hotkeys_enabled


class MacroEngine:
    def __init__(
        self,
        click_func: Callable[[int, int], None],
        sleep_func: Callable[[float], None] = time.sleep,
        rng: Optional[random.Random] = None,
    ):
        self.click_func = click_func
        self.sleep_func = sleep_func
        self.rng = rng or random.Random()
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.is_running = False
        self.worker_thread: Optional[threading.Thread] = None
        self.completed_cycles = 0

    def start(
        self,
        steps: List[Step],
        loops: int = 0,
        on_status: Optional[Callable[[str], None]] = None,
        on_cycle: Optional[Callable[[int], None]] = None,
        on_finished: Optional[Callable[[], None]] = None,
    ):
        if self.is_running:
            raise RuntimeError("Macro is already running")
        if not steps:
            raise ValueError("At least one region must be configured")
        if loops < 0:
            raise ValueError("Loops must be >= 0")

        self.stop_event.clear()
        self.pause_event.clear()
        self.is_running = True
        self.completed_cycles = 0
        self.worker_thread = threading.Thread(
            target=self._run_loop,
            args=(list(steps), loops, on_status, on_cycle, on_finished),
            daemon=True,
        )
        self.worker_thread.start()

    def stop(self):
        self.stop_event.set()
        self.pause_event.clear()

    def pause(self):
        if self.is_running:
            self.pause_event.set()

    def resume(self):
        if self.is_running:
            self.pause_event.clear()

    def _wait_if_paused(self, on_status: Optional[Callable[[str], None]] = None):
        while self.pause_event.is_set() and not self.stop_event.is_set():
            if on_status:
                on_status("pausado")
            self.sleep_func(0.05)
        if self.is_running and not self.stop_event.is_set() and on_status:
            on_status("executando")

    def _run_loop(
        self,
        steps: List[Step],
        loops: int,
        on_status: Optional[Callable[[str], None]] = None,
        on_cycle: Optional[Callable[[int], None]] = None,
        on_finished: Optional[Callable[[], None]] = None,
    ):
        try:
            if on_status:
                on_status("executando")

            cycle_target = loops if loops > 0 else None
            while not self.stop_event.is_set():
                if cycle_target is not None and self.completed_cycles >= cycle_target:
                    break

                for step in steps:
                    if self.stop_event.is_set():
                        break

                    self._wait_if_paused(on_status)
                    if self.stop_event.is_set():
                        break

                    x, y = random_point_in_region(step.region, self.rng)
                    self.click_func(x, y)

                    end_time = time.time() + step.delay
                    while time.time() < end_time:
                        if self.stop_event.is_set():
                            break
                        self._wait_if_paused(on_status)
                        self.sleep_func(0.05)

                if self.stop_event.is_set():
                    break

                self.completed_cycles += 1
                if on_cycle:
                    on_cycle(self.completed_cycles)
        finally:
            self.is_running = False
            self.stop_event.clear()
            self.pause_event.clear()
            if on_status:
                on_status("parado")
            if on_finished:
                on_finished()


if TK_AVAILABLE:
    class RegionSelector(tk.Toplevel):
        def __init__(self, parent, on_region_selected):
            super().__init__(parent)
            self.parent = parent
            self.on_region_selected = on_region_selected
            self.start_x = None
            self.start_y = None
            self.rect = None
            self.label = None

            self.withdraw()
            self.overrideredirect(True)
            self.attributes("-topmost", True)
            self.attributes("-alpha", 0.28)
            self.configure(bg="black")

            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            self.geometry(f"{screen_w}x{screen_h}+0+0")

            self.canvas = tk.Canvas(self, cursor="cross", bg="black", highlightthickness=0)
            self.canvas.pack(fill="both", expand=True)

            self.info_id = self.canvas.create_text(
                20,
                20,
                anchor="nw",
                text="Arraste para selecionar uma região | ESC cancela",
                fill="white",
                font=("Segoe UI", 11, "bold"),
            )
            self.size_id = self.canvas.create_text(
                20,
                48,
                anchor="nw",
                text="",
                fill="#aee1ff",
                font=("Segoe UI", 10),
            )

            self.canvas.bind("<ButtonPress-1>", self.on_press)
            self.canvas.bind("<B1-Motion>", self.on_drag)
            self.canvas.bind("<ButtonRelease-1>", self.on_release)
            self.bind("<Escape>", lambda e: self.cancel())

            self.after(80, self.show_overlay)

        def show_overlay(self):
            self.deiconify()
            self.focus_force()

        def on_press(self, event):
            self.start_x = event.x
            self.start_y = event.y
            if self.rect:
                self.canvas.delete(self.rect)
            self.rect = self.canvas.create_rectangle(
                self.start_x,
                self.start_y,
                self.start_x,
                self.start_y,
                outline="#ff4d4d",
                width=2,
                fill="#ffffff",
                stipple="gray25",
            )

        def on_drag(self, event):
            if self.rect and self.start_x is not None and self.start_y is not None:
                self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)
                region = normalize_region(self.start_x, self.start_y, event.x, event.y)
                width = region[2] - region[0]
                height = region[3] - region[1]
                self.canvas.itemconfig(
                    self.size_id,
                    text=f"x1={region[0]} y1={region[1]} x2={region[2]} y2={region[3]} | {width}x{height}",
                )

        def on_release(self, event):
            if self.start_x is None or self.start_y is None:
                self.destroy()
                return

            region = normalize_region(self.start_x, self.start_y, event.x, event.y)
            if not is_valid_region(region):
                messagebox.showwarning("Região inválida", "Selecione uma área maior.")
                self.destroy()
                return

            self.on_region_selected(region)
            self.destroy()

        def cancel(self):
            self.destroy()


    class MacroClickApp:
        def __init__(self, root):
            self.root = root
            self.root.title(APP_TITLE)
            self.root.geometry("980x640")
            self.root.minsize(900, 560)

            self.steps: List[Step] = []
            self.profile_path: Optional[Path] = None
            self.engine = MacroEngine(click_func=self._click_screen)
            self.hotkeys_registered = False

            self.build_ui()
            self.register_hotkeys_if_needed(show_message=False)
            self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        def _click_screen(self, x: int, y: int):
            if pyautogui is None:
                raise RuntimeError("pyautogui não está instalado")
            pyautogui.click(x, y)

        def build_ui(self):
            main = ttk.Frame(self.root, padding=12)
            main.pack(fill="both", expand=True)

            title = ttk.Label(main, text=APP_TITLE, font=("Segoe UI", 16, "bold"))
            title.pack(anchor="w", pady=(0, 10))

            topbar = ttk.Frame(main)
            topbar.pack(fill="x", pady=(0, 10))

            ttk.Label(topbar, text="Delay padrão (s):").grid(row=0, column=0, sticky="w")
            self.default_delay_var = tk.StringVar(value="1.0")
            ttk.Entry(topbar, textvariable=self.default_delay_var, width=10).grid(row=0, column=1, padx=(6, 14))

            ttk.Label(topbar, text="Loops (0 = infinito):").grid(row=0, column=2, sticky="w")
            self.loops_var = tk.StringVar(value="0")
            ttk.Entry(topbar, textvariable=self.loops_var, width=10).grid(row=0, column=3, padx=(6, 14))

            self.hotkeys_enabled_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(
                topbar,
                text="Atalhos globais (F6 iniciar | F7 pausar | F8 continuar | ESC parar)",
                variable=self.hotkeys_enabled_var,
                command=self.on_toggle_hotkeys,
            ).grid(row=0, column=4, sticky="w")

            toolbar = ttk.Frame(main)
            toolbar.pack(fill="x", pady=(0, 10))

            ttk.Button(toolbar, text="Adicionar região", command=self.add_region).pack(side="left", padx=4)
            ttk.Button(toolbar, text="Remover selecionada", command=self.remove_selected).pack(side="left", padx=4)
            ttk.Button(toolbar, text="Mover para cima", command=self.move_up).pack(side="left", padx=4)
            ttk.Button(toolbar, text="Mover para baixo", command=self.move_down).pack(side="left", padx=4)
            ttk.Button(toolbar, text="Limpar tudo", command=self.clear_all).pack(side="left", padx=4)
            ttk.Button(toolbar, text="Salvar perfil", command=self.save_profile_dialog).pack(side="left", padx=10)
            ttk.Button(toolbar, text="Carregar perfil", command=self.load_profile_dialog).pack(side="left", padx=4)

            content = ttk.Frame(main)
            content.pack(fill="both", expand=True)

            left = ttk.LabelFrame(content, text="Regiões adicionadas", padding=8)
            left.pack(side="left", fill="both", expand=True)

            columns = ("ordem", "regiao", "delay")
            self.tree = ttk.Treeview(left, columns=columns, show="headings", height=18)
            self.tree.heading("ordem", text="#")
            self.tree.heading("regiao", text="Região (x1, y1, x2, y2)")
            self.tree.heading("delay", text="Delay (s)")
            self.tree.column("ordem", width=50, anchor="center")
            self.tree.column("regiao", width=490, anchor="w")
            self.tree.column("delay", width=110, anchor="center")
            self.tree.pack(side="left", fill="both", expand=True)
            self.tree.bind("<Double-1>", self.edit_delay_selected)

            scrollbar = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
            scrollbar.pack(side="right", fill="y")
            self.tree.configure(yscrollcommand=scrollbar.set)

            right = ttk.LabelFrame(content, text="Painel", padding=12)
            right.pack(side="right", fill="y", padx=(12, 0))

            self.status_var = tk.StringVar(value="Status: parado")
            self.cycle_var = tk.StringVar(value="Ciclos concluídos: 0")
            self.profile_var = tk.StringVar(value="Perfil: não salvo")
            self.hotkey_status_var = tk.StringVar(value=self._build_hotkey_status_text())

            ttk.Label(right, textvariable=self.status_var, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))
            ttk.Label(right, textvariable=self.cycle_var).pack(anchor="w", pady=4)
            ttk.Label(right, textvariable=self.profile_var, wraplength=220).pack(anchor="w", pady=4)
            ttk.Label(right, textvariable=self.hotkey_status_var, wraplength=220).pack(anchor="w", pady=4)

            ttk.Separator(right, orient="horizontal").pack(fill="x", pady=10)

            ttk.Button(right, text="Iniciar", command=self.start_macro, width=22).pack(fill="x", pady=4)
            ttk.Button(right, text="Pausar", command=self.pause_macro, width=22).pack(fill="x", pady=4)
            ttk.Button(right, text="Continuar", command=self.resume_macro, width=22).pack(fill="x", pady=4)
            ttk.Button(right, text="Parar", command=self.stop_macro, width=22).pack(fill="x", pady=4)

            help_text = (
                "Dica: clique duplo em uma linha para editar o delay daquela região. "
                "Os cliques acontecem em posição aleatória dentro da área escolhida."
            )
            ttk.Label(main, text=help_text, foreground="#444").pack(anchor="w", pady=(10, 0))

        def _build_hotkey_status_text(self) -> str:
            if not self.hotkeys_enabled_var.get():
                return "Atalhos globais: desativados"
            if KEYBOARD_AVAILABLE:
                return "Atalhos globais: ativos (F6 iniciar | F7 pausar | F8 continuar | ESC parar)"
            return "Atalhos globais: biblioteca 'keyboard' não instalada"

        def _update_hotkey_status_label(self):
            self.hotkey_status_var.set(self._build_hotkey_status_text())

        def _update_profile_label(self):
            if self.profile_path:
                self.profile_var.set(f"Perfil: {self.profile_path}")
            else:
                self.profile_var.set("Perfil: não salvo")

        def add_region(self):
            if self.engine.is_running:
                messagebox.showinfo("Em execução", "Pare a macro antes de adicionar novas regiões.")
                return
            RegionSelector(self.root, self.on_region_selected)

        def on_region_selected(self, region: Region):
            try:
                delay = parse_delay(self.default_delay_var.get())
            except ValueError:
                messagebox.showerror("Delay inválido", "Informe um delay padrão válido, como 1 ou 1.5")
                return

            self.steps.append(Step(region=region, delay=delay))
            self.refresh_tree()

        def refresh_tree(self):
            for item in self.tree.get_children():
                self.tree.delete(item)

            for index, step in enumerate(self.steps, start=1):
                region_text = f"({step.region[0]}, {step.region[1]}, {step.region[2]}, {step.region[3]})"
                self.tree.insert("", "end", values=(index, region_text, step.delay))

        def get_selected_index(self) -> Optional[int]:
            selected = self.tree.selection()
            if not selected:
                return None
            return self.tree.index(selected[0])

        def remove_selected(self):
            if self.engine.is_running:
                messagebox.showinfo("Em execução", "Pare a macro antes de remover regiões.")
                return

            index = self.get_selected_index()
            if index is None:
                messagebox.showwarning("Nenhuma seleção", "Selecione uma região para remover.")
                return

            del self.steps[index]
            self.refresh_tree()

        def move_up(self):
            if self.engine.is_running:
                return
            index = self.get_selected_index()
            if index is None or index == 0:
                return
            self.steps[index - 1], self.steps[index] = self.steps[index], self.steps[index - 1]
            self.refresh_tree()
            self.tree.selection_set(self.tree.get_children()[index - 1])

        def move_down(self):
            if self.engine.is_running:
                return
            index = self.get_selected_index()
            if index is None or index >= len(self.steps) - 1:
                return
            self.steps[index + 1], self.steps[index] = self.steps[index], self.steps[index + 1]
            self.refresh_tree()
            self.tree.selection_set(self.tree.get_children()[index + 1])

        def clear_all(self):
            if self.engine.is_running:
                messagebox.showinfo("Em execução", "Pare a macro antes de limpar as regiões.")
                return
            if not self.steps:
                return
            if messagebox.askyesno("Limpar tudo", "Deseja remover todas as regiões?"):
                self.steps.clear()
                self.refresh_tree()
                self.cycle_var.set("Ciclos concluídos: 0")

        def edit_delay_selected(self, _event=None):
            if self.engine.is_running:
                return

            index = self.get_selected_index()
            if index is None:
                return

            step = self.steps[index]
            dialog = tk.Toplevel(self.root)
            dialog.title("Editar delay")
            dialog.geometry("300x120")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()

            ttk.Label(dialog, text="Novo delay desta região (segundos):").pack(pady=(14, 8))
            delay_var = tk.StringVar(value=str(step.delay))
            entry = ttk.Entry(dialog, textvariable=delay_var, width=18)
            entry.pack()
            entry.focus()

            def save():
                try:
                    new_delay = parse_delay(delay_var.get())
                except ValueError:
                    messagebox.showerror(
                        "Valor inválido",
                        "Digite um número válido maior ou igual a zero.",
                        parent=dialog,
                    )
                    return
                self.steps[index].delay = new_delay
                self.refresh_tree()
                dialog.destroy()

            ttk.Button(dialog, text="Salvar", command=save).pack(pady=12)
            dialog.bind("<Return>", lambda e: save())

        def parse_loops(self) -> int:
            try:
                loops = int(self.loops_var.get().strip())
            except ValueError:
                raise ValueError("Loops deve ser um número inteiro")
            if loops < 0:
                raise ValueError("Loops deve ser maior ou igual a zero")
            return loops

        def _set_status(self, value: str):
            self.root.after(0, lambda: self.status_var.set(f"Status: {value}"))

        def _set_cycle(self, value: int):
            self.root.after(0, lambda: self.cycle_var.set(f"Ciclos concluídos: {value}"))

        def _on_finished(self):
            self.root.after(0, lambda: None)

        def start_macro(self):
            if pyautogui is None:
                messagebox.showerror(
                    "Dependência ausente",
                    "A biblioteca pyautogui não está instalada.\n\nInstale com:\npip install pyautogui",
                )
                return

            try:
                loops = self.parse_loops()
                self.engine.start(
                    self.steps,
                    loops=loops,
                    on_status=self._set_status,
                    on_cycle=self._set_cycle,
                    on_finished=self._on_finished,
                )
            except ValueError as exc:
                if "region" in str(exc).lower() or "configured" in str(exc).lower():
                    messagebox.showwarning("Sem regiões", "Adicione pelo menos uma região antes de iniciar.")
                else:
                    messagebox.showwarning("Valor inválido", str(exc))
            except RuntimeError:
                messagebox.showinfo("Em execução", "A macro já está em execução.")
            except Exception as exc:
                messagebox.showerror("Erro", f"Não foi possível iniciar a macro:\n{exc}")

        def pause_macro(self):
            if self.engine.is_running:
                self.engine.pause()
                self.status_var.set("Status: pausado")

        def resume_macro(self):
            if self.engine.is_running:
                self.engine.resume()
                self.status_var.set("Status: executando")

        def stop_macro(self):
            if not self.engine.is_running:
                return
            self.status_var.set("Status: parando...")
            self.engine.stop()

        def save_profile_dialog(self):
            try:
                loops = self.parse_loops()
            except ValueError as exc:
                messagebox.showwarning("Valor inválido", str(exc))
                return

            path_str = filedialog.asksaveasfilename(
                title="Salvar perfil",
                defaultextension=".json",
                filetypes=[("JSON", "*.json")],
                initialfile=self.profile_path.name if self.profile_path else DEFAULT_PROFILE_PATH.name,
            )
            if not path_str:
                return

            path = Path(path_str)
            try:
                save_profile(path, self.steps, loops, self.hotkeys_enabled_var.get())
            except Exception as exc:
                messagebox.showerror("Erro ao salvar", f"Não foi possível salvar o perfil:\n{exc}")
                return

            self.profile_path = path
            self._update_profile_label()
            messagebox.showinfo("Perfil salvo", "Perfil salvo com sucesso.")

        def load_profile_dialog(self):
            if self.engine.is_running:
                messagebox.showinfo("Em execução", "Pare a macro antes de carregar um perfil.")
                return

            path_str = filedialog.askopenfilename(
                title="Carregar perfil",
                filetypes=[("JSON", "*.json")],
            )
            if not path_str:
                return

            path = Path(path_str)
            try:
                steps, loops, hotkeys_enabled = load_profile(path)
            except Exception as exc:
                messagebox.showerror("Erro ao carregar", f"Não foi possível carregar o perfil:\n{exc}")
                return

            self.steps = steps
            self.loops_var.set(str(loops))
            self.hotkeys_enabled_var.set(hotkeys_enabled)
            self.profile_path = path
            self.refresh_tree()
            self._update_profile_label()
            self.on_toggle_hotkeys(show_message=False)
            messagebox.showinfo("Perfil carregado", "Perfil carregado com sucesso.")

        def _safe_hotkey_call(self, callback: Callable[[], None]):
            self.root.after(0, callback)

        def register_hotkeys_if_needed(self, show_message: bool = True):
            if not self.hotkeys_enabled_var.get():
                self.unregister_hotkeys()
                self._update_hotkey_status_label()
                return

            if not KEYBOARD_AVAILABLE:
                self._update_hotkey_status_label()
                if show_message:
                    messagebox.showwarning(
                        "Atalhos indisponíveis",
                        "A biblioteca 'keyboard' não está instalada.\n\nInstale com:\npip install keyboard",
                    )
                return

            if self.hotkeys_registered:
                self._update_hotkey_status_label()
                return

            try:
                keyboard.add_hotkey("f6", lambda: self._safe_hotkey_call(self.start_macro))
                keyboard.add_hotkey("f7", lambda: self._safe_hotkey_call(self.pause_macro))
                keyboard.add_hotkey("f8", lambda: self._safe_hotkey_call(self.resume_macro))
                keyboard.add_hotkey("esc", lambda: self._safe_hotkey_call(self.stop_macro))
                self.hotkeys_registered = True
            except Exception as exc:
                self.hotkeys_registered = False
                if show_message:
                    messagebox.showwarning(
                        "Falha ao registrar atalhos",
                        "Não foi possível ativar os atalhos globais.\n"
                        "Em alguns sistemas pode ser necessário executar como administrador.\n\n"
                        f"Detalhe: {exc}",
                    )
            finally:
                self._update_hotkey_status_label()

        def unregister_hotkeys(self):
            if KEYBOARD_AVAILABLE and self.hotkeys_registered:
                try:
                    keyboard.unhook_all_hotkeys()
                except Exception:
                    pass
            self.hotkeys_registered = False
            self._update_hotkey_status_label()

        def on_toggle_hotkeys(self, show_message: bool = True):
            if self.hotkeys_enabled_var.get():
                self.register_hotkeys_if_needed(show_message=show_message)
            else:
                self.unregister_hotkeys()

        def on_close(self):
            self.engine.stop()
            self.unregister_hotkeys()
            self.root.destroy()


class MissingTkAppError(RuntimeError):
    pass


def raise_missing_tk_error():
    raise MissingTkAppError(
        "Este script precisa do módulo tkinter para abrir a interface gráfica.\n\n"
        "Como resolver:\n"
        "- Windows/macOS: reinstale Python usando a versão oficial, que normalmente já inclui tkinter.\n"
        "- Ubuntu/Debian: sudo apt install python3-tk\n"
        "- Fedora: sudo dnf install python3-tkinter\n"
        "- Arch: sudo pacman -S tk\n\n"
        f"Erro original: {TK_IMPORT_ERROR}"
    )


def run_tests():
    import tempfile
    import unittest

    class MacroTests(unittest.TestCase):
        def test_parse_delay_accepts_comma(self):
            self.assertEqual(parse_delay("1,5"), 1.5)

        def test_parse_delay_rejects_negative(self):
            with self.assertRaises(ValueError):
                parse_delay("-1")

        def test_normalize_region(self):
            self.assertEqual(normalize_region(10, 20, 5, 2), (5, 2, 10, 20))

        def test_normalize_region_keeps_ordered_region(self):
            self.assertEqual(normalize_region(1, 2, 3, 4), (1, 2, 3, 4))

        def test_invalid_region(self):
            self.assertFalse(is_valid_region((0, 0, 2, 2)))
            self.assertTrue(is_valid_region((0, 0, 3, 3)))

        def test_valid_region_accepts_larger_area(self):
            self.assertTrue(is_valid_region((10, 10, 20, 30)))

        def test_random_point_in_region(self):
            rng = random.Random(123)
            x, y = random_point_in_region((10, 20, 30, 40), rng)
            self.assertTrue(10 <= x <= 30)
            self.assertTrue(20 <= y <= 40)

        def test_step_serialization(self):
            step = Step(region=(1, 2, 10, 20), delay=1.25)
            recovered = Step.from_dict(step.to_dict())
            self.assertEqual(recovered, step)

        def test_save_and_load_profile(self):
            with tempfile.TemporaryDirectory() as temp_dir:
                path = Path(temp_dir) / "profile.json"
                steps = [Step(region=(10, 10, 20, 20), delay=1.0)]
                save_profile(path, steps, loops=3, hotkeys_enabled=True)
                loaded_steps, loaded_loops, loaded_hotkeys = load_profile(path)
                self.assertEqual(loaded_loops, 3)
                self.assertTrue(loaded_hotkeys)
                self.assertEqual(loaded_steps, steps)

        def test_engine_clicks_in_order(self):
            clicks = []
            engine = MacroEngine(
                click_func=lambda x, y: (clicks.append((x, y)), engine.stop()),
                sleep_func=lambda _: None,
                rng=random.Random(1),
            )
            steps = [Step(region=(5, 5, 5, 5), delay=0)]
            engine.start(steps)
            engine.worker_thread.join(timeout=1)
            self.assertEqual(clicks, [(5, 5)])
            self.assertFalse(engine.is_running)

        def test_engine_pause_and_resume(self):
            clicks = []
            proceed = {"allow_second": False}

            def click(x, y):
                clicks.append((x, y))
                if len(clicks) == 1:
                    engine.pause()
                    def release_pause():
                        proceed["allow_second"] = True
                        engine.resume()
                    threading.Timer(0.05, release_pause).start()
                elif len(clicks) == 2:
                    engine.stop()

            engine = MacroEngine(
                click_func=click,
                sleep_func=lambda s: time.sleep(min(s, 0.01)),
                rng=random.Random(2),
            )
            steps = [
                Step(region=(1, 1, 1, 1), delay=0),
                Step(region=(2, 2, 2, 2), delay=0),
            ]
            engine.start(steps)
            engine.worker_thread.join(timeout=1)
            self.assertEqual(clicks[:2], [(1, 1), (2, 2)])
            self.assertFalse(engine.is_running)

        def test_engine_respects_loop_limit(self):
            clicks = []
            engine = MacroEngine(
                click_func=lambda x, y: clicks.append((x, y)),
                sleep_func=lambda _: None,
                rng=random.Random(4),
            )
            steps = [Step(region=(7, 7, 7, 7), delay=0)]
            engine.start(steps, loops=2)
            engine.worker_thread.join(timeout=1)
            self.assertEqual(clicks, [(7, 7), (7, 7)])
            self.assertEqual(engine.completed_cycles, 2)
            self.assertFalse(engine.is_running)

    suite = unittest.defaultTestLoader.loadTestsFromTestCase(MacroTests)
    runner = unittest.TextTestRunner(verbosity=2)
    return runner.run(suite).wasSuccessful()


def main():
    if "--test" in sys.argv:
        success = run_tests()
        raise SystemExit(0 if success else 1)

    if not TK_AVAILABLE:
        raise_missing_tk_error()

    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    MacroClickApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
