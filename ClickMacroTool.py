import random
import threading
import time
import sys
from dataclasses import dataclass
from typing import Callable, List, Optional, Tuple

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
    TK_AVAILABLE = True
    TK_IMPORT_ERROR = None
except ModuleNotFoundError as exc:
    tk = None
    ttk = None
    messagebox = None
    TK_AVAILABLE = False
    TK_IMPORT_ERROR = exc

try:
    import pyautogui
except ImportError:
    pyautogui = None


Region = Tuple[int, int, int, int]


@dataclass
class Step:
    region: Region
    delay: float


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
        self.is_running = False
        self.worker_thread: Optional[threading.Thread] = None

    def start(self, steps: List[Step], on_status: Optional[Callable[[str], None]] = None):
        if self.is_running:
            raise RuntimeError("Macro is already running")
        if not steps:
            raise ValueError("At least one region must be configured")

        self.stop_event.clear()
        self.is_running = True
        self.worker_thread = threading.Thread(
            target=self._run_loop,
            args=(list(steps), on_status),
            daemon=True,
        )
        self.worker_thread.start()

    def stop(self):
        self.stop_event.set()

    def _run_loop(self, steps: List[Step], on_status: Optional[Callable[[str], None]] = None):
        try:
            if on_status:
                on_status("executando")
            while not self.stop_event.is_set():
                for step in steps:
                    if self.stop_event.is_set():
                        break
                    x, y = random_point_in_region(step.region, self.rng)
                    self.click_func(x, y)

                    end_time = time.time() + step.delay
                    while time.time() < end_time:
                        if self.stop_event.is_set():
                            break
                        self.sleep_func(0.05)
        finally:
            self.is_running = False
            self.stop_event.clear()
            if on_status:
                on_status("parado")


if TK_AVAILABLE:
    class RegionSelector(tk.Toplevel):
        def __init__(self, parent, on_region_selected):
            super().__init__(parent)
            self.parent = parent
            self.on_region_selected = on_region_selected
            self.start_x = None
            self.start_y = None
            self.rect = None

            self.withdraw()
            self.overrideredirect(True)
            self.attributes("-topmost", True)
            self.attributes("-alpha", 0.25)
            self.configure(bg="black")

            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            self.geometry(f"{screen_w}x{screen_h}+0+0")

            self.canvas = tk.Canvas(self, cursor="cross", bg="black", highlightthickness=0)
            self.canvas.pack(fill="both", expand=True)

            self.canvas.bind("<ButtonPress-1>", self.on_press)
            self.canvas.bind("<B1-Motion>", self.on_drag)
            self.canvas.bind("<ButtonRelease-1>", self.on_release)
            self.bind("<Escape>", lambda e: self.cancel())

            self.after(100, self.show_overlay)

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
                outline="red",
                width=2,
            )

        def on_drag(self, event):
            if self.rect:
                self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

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
            self.root.title("Macro de Cliques por Região")
            self.root.geometry("760x500")
            self.root.minsize(720, 450)

            self.steps: List[Step] = []
            self.engine = MacroEngine(click_func=self._click_screen)

            self.build_ui()

        def _click_screen(self, x: int, y: int):
            if pyautogui is None:
                raise RuntimeError("pyautogui não está instalado")
            pyautogui.click(x, y)

        def build_ui(self):
            main = ttk.Frame(self.root, padding=12)
            main.pack(fill="both", expand=True)

            title = ttk.Label(main, text="Automação de Cliques por Região", font=("Segoe UI", 15, "bold"))
            title.pack(anchor="w", pady=(0, 10))

            controls = ttk.Frame(main)
            controls.pack(fill="x", pady=(0, 10))

            ttk.Label(controls, text="Delay padrão entre cliques (s):").grid(row=0, column=0, sticky="w")
            self.default_delay_var = tk.StringVar(value="1.0")
            ttk.Entry(controls, textvariable=self.default_delay_var, width=10).grid(row=0, column=1, padx=(8, 14))

            ttk.Button(controls, text="Adicionar região", command=self.add_region).grid(row=0, column=2, padx=4)
            ttk.Button(controls, text="Remover selecionada", command=self.remove_selected).grid(row=0, column=3, padx=4)
            ttk.Button(controls, text="Limpar tudo", command=self.clear_all).grid(row=0, column=4, padx=4)

            list_frame = ttk.LabelFrame(main, text="Regiões adicionadas", padding=8)
            list_frame.pack(fill="both", expand=True)

            columns = ("ordem", "regiao", "delay")
            self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=14)
            self.tree.heading("ordem", text="#")
            self.tree.heading("regiao", text="Região (x1, y1, x2, y2)")
            self.tree.heading("delay", text="Delay (s)")

            self.tree.column("ordem", width=50, anchor="center")
            self.tree.column("regiao", width=470, anchor="w")
            self.tree.column("delay", width=100, anchor="center")
            self.tree.pack(side="left", fill="both", expand=True)

            scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
            scrollbar.pack(side="right", fill="y")
            self.tree.configure(yscrollcommand=scrollbar.set)

            self.tree.bind("<Double-1>", self.edit_delay_selected)

            bottom = ttk.Frame(main)
            bottom.pack(fill="x", pady=(12, 0))

            self.status_var = tk.StringVar(value="Status: parado")
            ttk.Label(bottom, textvariable=self.status_var).pack(side="left")

            actions = ttk.Frame(bottom)
            actions.pack(side="right")
            ttk.Button(actions, text="Iniciar", command=self.start_macro).pack(side="left", padx=4)
            ttk.Button(actions, text="Parar", command=self.stop_macro).pack(side="left", padx=4)

            help_text = (
                "Dica: clique duplo em uma linha para editar o delay daquela região. "
                "O clique será feito em posição aleatória dentro da área selecionada."
            )
            ttk.Label(main, text=help_text, foreground="#444").pack(anchor="w", pady=(8, 0))

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

        def remove_selected(self):
            if self.engine.is_running:
                messagebox.showinfo("Em execução", "Pare a macro antes de remover regiões.")
                return

            selected = self.tree.selection()
            if not selected:
                messagebox.showwarning("Nenhuma seleção", "Selecione uma região para remover.")
                return

            index = self.tree.index(selected[0])
            del self.steps[index]
            self.refresh_tree()

        def clear_all(self):
            if self.engine.is_running:
                messagebox.showinfo("Em execução", "Pare a macro antes de limpar as regiões.")
                return

            if not self.steps:
                return

            if messagebox.askyesno("Limpar tudo", "Deseja remover todas as regiões?"):
                self.steps.clear()
                self.refresh_tree()

        def edit_delay_selected(self, _event=None):
            if self.engine.is_running:
                return

            selected = self.tree.selection()
            if not selected:
                return

            index = self.tree.index(selected[0])
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

        def _set_status(self, value: str):
            self.root.after(0, lambda: self.status_var.set(f"Status: {value}"))

        def start_macro(self):
            if pyautogui is None:
                messagebox.showerror(
                    "Dependência ausente",
                    "A biblioteca pyautogui não está instalada.\n\nInstale com:\npip install pyautogui",
                )
                return

            try:
                self.engine.start(self.steps, on_status=self._set_status)
            except ValueError:
                messagebox.showwarning("Sem regiões", "Adicione pelo menos uma região antes de iniciar.")
            except RuntimeError:
                messagebox.showinfo("Em execução", "A macro já está em execução.")
            except Exception as exc:
                messagebox.showerror("Erro", f"Não foi possível iniciar a macro:\n{exc}")

        def stop_macro(self):
            if not self.engine.is_running:
                return
            self.status_var.set("Status: parando...")
            self.engine.stop()


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
    import unittest

    class MacroTests(unittest.TestCase):
        def test_parse_delay_accepts_comma(self):
            self.assertEqual(parse_delay("1,5"), 1.5)

        def test_parse_delay_rejects_negative(self):
            with self.assertRaises(ValueError):
                parse_delay("-1")

        def test_normalize_region(self):
            self.assertEqual(normalize_region(10, 20, 5, 2), (5, 2, 10, 20))

        def test_invalid_region(self):
            self.assertFalse(is_valid_region((0, 0, 2, 2)))
            self.assertTrue(is_valid_region((0, 0, 3, 3)))

        def test_random_point_in_region(self):
            rng = random.Random(123)
            x, y = random_point_in_region((10, 20, 30, 40), rng)
            self.assertTrue(10 <= x <= 30)
            self.assertTrue(20 <= y <= 40)

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

        def test_normalize_region_keeps_ordered_region(self):
            self.assertEqual(normalize_region(1, 2, 3, 4), (1, 2, 3, 4))

        def test_valid_region_accepts_larger_area(self):
            self.assertTrue(is_valid_region((10, 10, 20, 30)))

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
