import time
import pytest
from engine import Action, Runner, parse_key_spec


class Fake:
    def __init__(self): self.log = []
    def press_key(self, s): self.log.append(("key", s))
    def click(self, x, y, b="left", d=False): self.log.append(("click", x, y, b, d))


def wait(cond, t=3):
    end = time.time() + t
    while time.time() < end and not cond():
        time.sleep(0.005)
    return cond()


def test_parse():
    assert parse_key_spec("Ctrl+Shift+a") == (["ctrl", "shift"], "a")
    assert parse_key_spec("Enter") == ([], "enter")
    assert parse_key_spec("+") == ([], "+")
    for bad in ("", "  ", "foo+a", "ctrl+", "ctrl+shift"):
        with pytest.raises(ValueError):
            parse_key_spec(bad)


def test_validate():
    with pytest.raises(ValueError): Action(kind="click", button="x").validate()
    with pytest.raises(ValueError): Action(kind="click", x=5).validate()
    with pytest.raises(ValueError): Action(kind="key", key="").validate()
    with pytest.raises(ValueError): Action(kind="key", key="a", delay=-1).validate()
    Action(kind="click", x=1, y=2).validate()


def test_sequence_repeats_in_order():
    f = Fake(); r = Runner(f); r.interval = 0
    r.actions = [Action(key="a"), Action(kind="click", x=10, y=20, button="right"), Action(key="b")]
    r.repeat = 2
    r.start(); assert wait(lambda: not r.running)
    assert f.log == [("key", "a"), ("click", 10, 20, "right", False), ("key", "b")] * 2
    assert r.count == 6


def test_interval_respected():
    f = Fake(); r = Runner(f); r.interval = 0.1; r.repeat = 1
    r.actions = [Action(key="a"), Action(key="b"), Action(key="c")]
    t = time.time(); r.start(); wait(lambda: not r.running)
    assert time.time() - t >= 0.29  # 3 waits of 0.1s


def test_per_action_delay_overrides_interval():
    f = Fake(); r = Runner(f); r.interval = 5; r.repeat = 1
    r.actions = [Action(key="a", delay=0)]
    t = time.time(); r.start(); wait(lambda: not r.running)
    assert time.time() - t < 1


def test_stop_is_immediate_even_with_long_interval():
    f = Fake(); r = Runner(f); r.interval = 60
    r.actions = [Action(key="a")]
    r.start(); assert wait(lambda: len(f.log) == 1)
    t = time.time(); r.stop()
    assert not r.running and time.time() - t < 1


def test_toggle_and_empty_start():
    r = Runner(Fake())
    with pytest.raises(ValueError): r.start()
    r.actions = [Action(key="a")]; r.interval = 0.01
    r.toggle(); assert r.running
    r.toggle(); assert not r.running


def test_backend_error_reported_and_stops():
    class Boom(Fake):
        def press_key(self, s): raise RuntimeError("nope")
    got = []
    r = Runner(Boom()); r.actions = [Action(key="a")]; r.on_finish = got.append
    r.start(); assert wait(lambda: got)
    assert isinstance(got[0], RuntimeError) and not r.running


def test_roundtrip_dict():
    a = Action(kind="click", x=1, y=2, button="middle", double=True, delay=0.5)
    assert Action.from_dict(a.to_dict()) == a
