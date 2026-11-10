"""Real end-to-end test: a Tk window with a Button and an Entry; the actual
PynputBackend must click the button and type into the entry via the OS (X11).
Run under a display, e.g.  xvfb-run -a pytest test_integration.py"""
import os
import time
import pytest

pytestmark = pytest.mark.skipif(not os.environ.get("DISPLAY"), reason="needs a display")


def pump(root, secs):
    end = time.time() + secs
    while time.time() < end:
        root.update(); time.sleep(0.01)


@pytest.fixture
def win():
    import tkinter as tk
    root = tk.Tk(); root.geometry("300x200+50+50")
    clicks = {"n": 0, "right": 0}
    btn = tk.Button(root, text="Target", command=lambda: clicks.__setitem__("n", clicks["n"] + 1))
    btn.place(x=20, y=20, width=120, height=50)
    btn.bind("<Button-3>", lambda e: clicks.__setitem__("right", clicks["right"] + 1))
    ent = tk.Entry(root); ent.place(x=20, y=100, width=200)
    root.update(); root.focus_force(); ent.focus_force(); pump(root, 0.3)
    yield root, btn, ent, clicks
    root.destroy()


def center(root, w):
    root.update()
    return w.winfo_rootx() + w.winfo_width() // 2, w.winfo_rooty() + w.winfo_height() // 2


def run_sequence(root, actions, repeat=1):
    from engine import Runner, PynputBackend
    r = Runner(PynputBackend()); r.actions = actions; r.interval = 0.05; r.repeat = repeat
    r.start()
    while r.running:
        root.update(); time.sleep(0.01)
    pump(root, 0.2)


def test_clicks_a_chosen_on_screen_button(win):
    from engine import Action
    root, btn, ent, clicks = win
    x, y = center(root, btn)
    run_sequence(root, [Action(kind="click", x=x, y=y)], repeat=5)
    assert clicks["n"] == 5


def test_right_click_and_double_click(win):
    from engine import Action
    root, btn, ent, clicks = win
    x, y = center(root, btn)
    run_sequence(root, [Action(kind="click", x=x, y=y, button="right")], repeat=2)
    assert clicks["right"] == 2
    run_sequence(root, [Action(kind="click", x=x, y=y, double=True)], repeat=1)
    assert clicks["n"] == 2


def test_presses_selected_keys(win):
    from engine import Action
    root, btn, ent, clicks = win
    x, y = center(root, ent)
    run_sequence(root, [Action(kind="click", x=x, y=y), Action(key="h"), Action(key="i"),
                        Action(key="space"), Action(key="x")])
    assert ent.get() == "hi x"


def test_mixed_keys_and_button_clicks_repeating(win):
    from engine import Action
    root, btn, ent, clicks = win
    bx, by = center(root, btn); ex, ey = center(root, ent)
    run_sequence(root, [Action(kind="click", x=ex, y=ey), Action(key="z"),
                        Action(kind="click", x=bx, y=by)], repeat=3)
    assert ent.get() == "zzz" and clicks["n"] == 3
