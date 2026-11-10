# Key Clicker

Desktop auto key-presser / auto-clicker. Choose which keys to press and which
screen buttons to click, set the interval, and run.

```bash
pip install -r requirements.txt
python app.py
```
(Needs Python 3 with Tk — included on Windows/macOS; on Linux `sudo apt install python3-tk`. X11 on Linux; macOS needs Accessibility permission.)

- **Key press**: type/choose a key (`space`, `f5`, `a`, `ctrl+c`…) or **Capture key**.
- **Click a button on screen**: **Pick a spot on screen…** hides the window — click the button you want, and its position is stored. You can also type x,y or click at the cursor. Left/right/middle and double-click supported.
- Actions run top to bottom, then repeat. **Interval** = seconds between actions; **Repeat** 0 = forever.
- **F6** starts/stops from anywhere. Save/Load sequences as JSON.

Tests: `pip install pytest`, then `pytest test_engine.py` and `xvfb-run -a pytest` on Linux (the integration tests click a real Tk button and type into a real entry through the OS).
