"""Drives the real Key Clicker window: adds actions through its widgets, presses Start."""
import os, time, json
import pytest
pytestmark = pytest.mark.skipif(not os.environ.get("DISPLAY"), reason="needs a display")


def pump(root, secs):
    end = time.time() + secs
    while time.time() < end:
        root.update(); time.sleep(0.01)


def test_gui_start_stop_and_save_load(tmp_path):
    import tkinter as tk
    from app import App
    app = App(); app.geometry("560x520+400+300"); pump(app, 0.3)
    target = tk.Toplevel(app); target.geometry("200x120+20+20")
    n = {"c": 0}
    b = tk.Button(target, text="hit me", command=lambda: n.__setitem__("c", n["c"] + 1)); b.pack(fill="both", expand=True)
    pump(app, 0.3)
    x, y = b.winfo_rootx() + 10, b.winfo_rooty() + 10
    app.x_var.set(str(x)); app.y_var.set(str(y)); app._add_click_xy()
    app.key_var.set("bogus+++"); app._append = lambda a: None  # no modal in tests
    assert len(app.runner.actions) == 1
    app.interval_var.set("0.05"); app.repeat_var.set("4")
    app.toggle()
    t = time.time()
    while app.runner.running and time.time() - t < 5: pump(app, 0.02)
    pump(app, 0.2)
    assert n["c"] == 4 and "Stopped after 4" in app.status.get()
    p = tmp_path / "s.json"
    p.write_text(json.dumps({"interval": "0.2", "repeat": "7", "actions": [{"kind": "key", "key": "f5"}]}))
    app.load_file(str(p))
    assert app.interval_var.get() == "0.2" and app.runner.actions[0].describe() == "Press f5"
    app._quit()
