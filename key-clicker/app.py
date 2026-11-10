"""Key Clicker - GUI. Run:  python app.py

Pick what to press/click, set the interval, hit Start (or F6 anywhere)."""
import json
import queue
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from engine import Action, Runner, PynputBackend, MOUSE_BUTTONS, parse_key_spec

HOTKEY = "<f6>"
COMMON_KEYS = ["space", "enter", "tab", "esc", "up", "down", "left", "right",
               "backspace", "delete", "home", "end", "page_up", "page_down",
               "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10", "f11", "f12",
               "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o",
               "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z",
               "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
               "ctrl+c", "ctrl+v", "ctrl+a", "ctrl+s", "alt+tab"]


class App(tk.Tk):
    def __init__(self, backend=None):
        super().__init__()
        self.title("Key Clicker")
        self.geometry("560x520")
        self.backend = backend or PynputBackend()
        self.runner = Runner(self.backend)
        # worker/hotkey threads must not touch Tk: they post to a queue the UI polls
        self.events = queue.Queue()
        self.runner.on_step = lambda i, a: self.events.put((self._on_step, (i, a)))
        self.runner.on_finish = lambda err: self.events.put((self._on_finish, (err,)))
        self.after(50, self._poll)
        self.hotkeys = None
        self._build()
        self._start_hotkey()
        self.protocol("WM_DELETE_WINDOW", self._quit)

    # ---- UI ----
    def _build(self):
        pad = dict(padx=8, pady=4)
        add = ttk.LabelFrame(self, text="Add an action")
        add.pack(fill="x", **pad)

        self.key_var = tk.StringVar(value="space")
        r1 = ttk.Frame(add); r1.pack(fill="x", **pad)
        ttk.Label(r1, text="Key:").pack(side="left")
        self.key_box = ttk.Combobox(r1, textvariable=self.key_var, values=COMMON_KEYS, width=14)
        self.key_box.pack(side="left", padx=4)
        ttk.Button(r1, text="Capture key", command=self._capture_key).pack(side="left")
        ttk.Button(r1, text="Add key press", command=self._add_key).pack(side="right")

        r2 = ttk.Frame(add); r2.pack(fill="x", **pad)
        self.btn_var = tk.StringVar(value="left")
        self.dbl_var = tk.BooleanVar(value=False)
        ttk.Label(r2, text="Mouse:").pack(side="left")
        ttk.Combobox(r2, textvariable=self.btn_var, values=MOUSE_BUTTONS, width=7,
                     state="readonly").pack(side="left", padx=4)
        ttk.Checkbutton(r2, text="double", variable=self.dbl_var).pack(side="left")
        ttk.Button(r2, text="Add click at cursor", command=self._add_click_here).pack(side="right")
        ttk.Button(r2, text="Pick a spot on screen…", command=self._pick_spot).pack(side="right", padx=4)

        r3 = ttk.Frame(add); r3.pack(fill="x", **pad)
        self.x_var, self.y_var = tk.StringVar(), tk.StringVar()
        ttk.Label(r3, text="or type x,y:").pack(side="left")
        ttk.Entry(r3, textvariable=self.x_var, width=6).pack(side="left", padx=2)
        ttk.Entry(r3, textvariable=self.y_var, width=6).pack(side="left", padx=2)
        ttk.Button(r3, text="Add click at x,y", command=self._add_click_xy).pack(side="left", padx=4)

        lst = ttk.LabelFrame(self, text="Sequence (runs top to bottom, then repeats)")
        lst.pack(fill="both", expand=True, **pad)
        self.listbox = tk.Listbox(lst, height=8, activestyle="none")
        self.listbox.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        side = ttk.Frame(lst); side.pack(side="right", padx=4)
        for text, cmd in (("Up", lambda: self._move(-1)), ("Down", lambda: self._move(1)),
                          ("Remove", self._remove), ("Clear", self._clear),
                          ("Save…", self._save), ("Load…", self._load)):
            ttk.Button(side, text=text, command=cmd, width=8).pack(pady=1)

        opt = ttk.LabelFrame(self, text="Timing")
        opt.pack(fill="x", **pad)
        self.interval_var = tk.StringVar(value="1.0")
        self.repeat_var = tk.StringVar(value="0")
        ttk.Label(opt, text="Interval between actions (seconds):").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(opt, textvariable=self.interval_var, width=8).grid(row=0, column=1)
        ttk.Label(opt, text="Repeat sequence (0 = forever):").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(opt, textvariable=self.repeat_var, width=8).grid(row=1, column=1)

        run = ttk.Frame(self); run.pack(fill="x", **pad)
        self.start_btn = ttk.Button(run, text="Start (F6)", command=self.toggle)
        self.start_btn.pack(side="left")
        self.status = tk.StringVar(value="Idle")
        ttk.Label(run, textvariable=self.status).pack(side="left", padx=10)

    # ---- adding actions ----
    def _refresh(self):
        self.listbox.delete(0, "end")
        for i, a in enumerate(self.runner.actions, 1):
            self.listbox.insert("end", f"{i}. {a.describe()}")

    def _append(self, action):
        try:
            action.validate()
        except ValueError as e:
            messagebox.showerror("Invalid action", str(e), parent=self)
            return
        self.runner.actions.append(action)
        self._refresh()

    def _add_key(self):
        self._append(Action(kind="key", key=self.key_var.get().strip()))

    def _add_click_here(self):
        self._append(Action(kind="click", button=self.btn_var.get(), double=self.dbl_var.get()))

    def _add_click_xy(self):
        try:
            x, y = int(self.x_var.get()), int(self.y_var.get())
        except ValueError:
            messagebox.showerror("Invalid position", "x and y must be whole numbers.", parent=self)
            return
        self._append(Action(kind="click", x=x, y=y, button=self.btn_var.get(), double=self.dbl_var.get()))

    def _capture_key(self):
        """Next key (with modifiers) pressed in a small popup becomes the key."""
        top = tk.Toplevel(self); top.title("Press a key"); top.geometry("260x80")
        ttk.Label(top, text="Press the key or combo you want…").pack(expand=True)
        top.focus_force()
        mods = {"Control_L": "ctrl", "Control_R": "ctrl", "Shift_L": "shift", "Shift_R": "shift",
                "Alt_L": "alt", "Alt_R": "alt"}

        def on_key(ev):
            if ev.keysym in mods:
                return
            name = {"Return": "enter", "Escape": "esc", "BackSpace": "backspace", "Prior": "page_up",
                    "Next": "page_down", "space": "space"}.get(ev.keysym, ev.keysym.lower())
            parts = [m for bit, m in ((0x4, "ctrl"), (0x1, "shift"), (0x8, "alt")) if ev.state & bit]
            self.key_var.set("+".join(parts + [name]))
            top.destroy()
        top.bind("<KeyPress>", on_key)

    def _pick_spot(self):
        """Hide the window; the next real click anywhere on screen is recorded as a target."""
        from pynput import mouse
        self.withdraw()
        button, double = self.btn_var.get(), self.dbl_var.get()

        def on_click(x, y, b, pressed):
            if pressed:
                self.events.put((self._finish_pick, (int(x), int(y), button, double)))
                return False
        self._picker = mouse.Listener(on_click=on_click)
        self._picker.start()

    def _finish_pick(self, x, y, button, double):
        self.deiconify()
        self._append(Action(kind="click", x=x, y=y, button=button, double=double,
                            label="picked"))

    # ---- list editing / persistence ----
    def _sel(self):
        s = self.listbox.curselection()
        return s[0] if s else None

    def _move(self, d):
        i = self._sel()
        acts = self.runner.actions
        if i is None or not 0 <= i + d < len(acts):
            return
        acts[i], acts[i + d] = acts[i + d], acts[i]
        self._refresh(); self.listbox.selection_set(i + d)

    def _remove(self):
        i = self._sel()
        if i is not None:
            del self.runner.actions[i]; self._refresh()

    def _clear(self):
        self.runner.actions.clear(); self._refresh()

    def _save(self):
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if p:
            with open(p, "w") as f:
                json.dump({"interval": self.interval_var.get(), "repeat": self.repeat_var.get(),
                           "actions": [a.to_dict() for a in self.runner.actions]}, f, indent=2)

    def _load(self):
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if p:
            self.load_file(p)

    def load_file(self, path):
        with open(path) as f:
            d = json.load(f)
        self.interval_var.set(str(d.get("interval", "1.0")))
        self.repeat_var.set(str(d.get("repeat", "0")))
        self.runner.actions[:] = [Action.from_dict(a) for a in d.get("actions", [])]
        self._refresh()

    # ---- running ----
    def toggle(self):
        if self.runner.running:
            self.runner.stop()
            return
        try:
            self.runner.interval = float(self.interval_var.get())
            self.runner.repeat = int(self.repeat_var.get())
            self.runner.start()
        except ValueError as e:
            messagebox.showerror("Cannot start", str(e), parent=self)
            return
        self.start_btn.config(text="Stop (F6)")
        self.status.set("Running…")

    def _on_step(self, i, a):
        self.status.set(f"Running… {self.runner.count} actions done")
        self.listbox.selection_clear(0, "end"); self.listbox.selection_set(i)

    def _on_finish(self, err):
        self.start_btn.config(text="Start (F6)")
        self.status.set(f"Error: {err}" if err else f"Stopped after {self.runner.count} actions")

    def _start_hotkey(self):
        try:
            from pynput import keyboard
            self.hotkeys = keyboard.GlobalHotKeys({HOTKEY: lambda: self.events.put((self.toggle, ()))})
            self.hotkeys.start()
        except Exception:
            self.status.set("Idle (global F6 hotkey unavailable)")

    def _poll(self):
        try:
            while True:
                fn, args = self.events.get_nowait()
                fn(*args)
        except queue.Empty:
            pass
        self.after(50, self._poll)

    def _quit(self):
        self.runner.stop(join=False)
        if self.hotkeys:
            self.hotkeys.stop()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
