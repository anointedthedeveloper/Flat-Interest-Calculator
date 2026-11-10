"""Core engine for Key Clicker: actions, key parsing, and the timed runner.

The engine is GUI-free and talks to the OS through a small backend object, so it
can be unit-tested with a fake backend and run for real with PynputBackend.
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass, asdict
from typing import Callable, List, Optional

MODIFIERS = {"ctrl", "shift", "alt", "cmd"}
MOUSE_BUTTONS = ("left", "right", "middle")
ALIASES = {
    "control": "ctrl", "return": "enter", "esc": "esc", "escape": "esc",
    "del": "delete", "ins": "insert", "pgup": "page_up", "pgdn": "page_down",
    "win": "cmd", "super": "cmd", "meta": "cmd", " ": "space",
}


def parse_key_spec(spec: str):
    """'ctrl+shift+a' -> (['ctrl', 'shift'], 'a'). Raises ValueError if invalid."""
    if spec is None or not spec.strip():
        raise ValueError("empty key")
    spec = spec.strip()
    parts = ["+" if spec == "+" else p for p in (spec.split("+") if spec != "+" else ["+"])]
    parts = [p.strip().lower() if len(p.strip()) > 1 else p for p in parts]
    parts = [ALIASES.get(p, p) for p in parts]
    if any(p == "" for p in parts):
        raise ValueError(f"bad key combo: {spec!r}")
    *mods, key = parts
    for m in mods:
        if m not in MODIFIERS:
            raise ValueError(f"{m!r} is not a modifier (ctrl/shift/alt/cmd) in {spec!r}")
    if key in MODIFIERS and mods:
        raise ValueError(f"combo {spec!r} needs a non-modifier key last")
    return mods, key


@dataclass
class Action:
    """One step. kind='key' presses `key`; kind='click' clicks at (x, y) (or at the
    current cursor when x/y are None). `delay` is seconds to wait AFTER this step
    (None -> use the global interval)."""
    kind: str = "key"
    key: str = ""
    x: Optional[int] = None
    y: Optional[int] = None
    button: str = "left"
    double: bool = False
    delay: Optional[float] = None
    label: str = ""

    def validate(self):
        if self.kind == "key":
            parse_key_spec(self.key)
        elif self.kind == "click":
            if self.button not in MOUSE_BUTTONS:
                raise ValueError(f"unknown mouse button {self.button!r}")
            if (self.x is None) != (self.y is None):
                raise ValueError("click needs both x and y, or neither")
        else:
            raise ValueError(f"unknown action kind {self.kind!r}")
        if self.delay is not None and self.delay < 0:
            raise ValueError("delay cannot be negative")

    def describe(self) -> str:
        if self.kind == "key":
            return f"Press {self.key}"
        where = "at cursor" if self.x is None else f"at ({self.x}, {self.y})"
        what = ("Double-click " if self.double else "Click ") + self.button
        name = f" [{self.label}]" if self.label else ""
        return f"{what} {where}{name}"

    def to_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, d):
        return cls(**{k: v for k, v in d.items() if k in cls.__dataclass_fields__})


class PynputBackend:
    """Real keyboard/mouse control. Imported lazily so tests don't need a display."""

    def __init__(self):
        from pynput import keyboard, mouse
        self._kb_mod, self._ms_mod = keyboard, mouse
        self._kb, self._ms = keyboard.Controller(), mouse.Controller()

    def _resolve(self, name: str):
        K = self._kb_mod.Key
        if len(name) == 1:
            return name
        try:
            return getattr(K, name)
        except AttributeError:
            raise ValueError(f"unknown key {name!r}")

    def press_key(self, spec: str):
        mods, key = parse_key_spec(spec)
        mod_keys = [self._resolve(m) for m in mods]
        k = self._resolve(key)
        for m in mod_keys:
            self._kb.press(m)
        try:
            self._kb.press(k)
            self._kb.release(k)
        finally:
            for m in reversed(mod_keys):
                self._kb.release(m)

    def click(self, x, y, button="left", double=False):
        if x is not None:
            self._ms.position = (x, y)
            time.sleep(0.02)  # let the pointer settle before the press
        self._ms.click(getattr(self._ms_mod.Button, button), 2 if double else 1)

    def position(self):
        return tuple(self._ms.position)


class Runner:
    """Runs a list of actions in a loop on a background thread."""

    def __init__(self, backend, sleep: Callable[[float], None] = None):
        self.backend = backend
        self.actions: List[Action] = []
        self.interval = 1.0      # default seconds between steps
        self.repeat = 0          # 0 = until stopped
        self.on_step: Optional[Callable[[int, Action], None]] = None
        self.on_finish: Optional[Callable[[Optional[BaseException]], None]] = None
        self.count = 0           # completed steps
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._sleep = sleep

    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.running:
            return
        if not self.actions:
            raise ValueError("add at least one action first")
        if self.interval < 0:
            raise ValueError("interval cannot be negative")
        for a in self.actions:
            a.validate()
        self._stop.clear()
        self.count = 0
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self, join: bool = True):
        self._stop.set()
        if join and self._thread and self._thread is not threading.current_thread():
            self._thread.join(timeout=2)

    def toggle(self):
        self.stop() if self.running else self.start()

    def _wait(self, seconds: float):
        if self._sleep:
            self._sleep(seconds)
        else:
            self._stop.wait(seconds)  # wakes immediately on stop

    def _do(self, a: Action):
        if a.kind == "key":
            self.backend.press_key(a.key)
        else:
            self.backend.click(a.x, a.y, a.button, a.double)

    def _run(self):
        err = None
        try:
            loops = 0
            while not self._stop.is_set():
                for i, a in enumerate(self.actions):
                    if self._stop.is_set():
                        break
                    self._do(a)
                    self.count += 1
                    if self.on_step:
                        self.on_step(i, a)
                    self._wait(self.interval if a.delay is None else a.delay)
                loops += 1
                if self.repeat and loops >= self.repeat:
                    break
        except BaseException as e:  # surface backend errors to the UI
            err = e
        finally:
            if self.on_finish:
                self.on_finish(err)
