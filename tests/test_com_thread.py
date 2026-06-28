import threading
import pytest
from com_thread import COMThread


def test_run_executes_callable_on_background_thread():
    com = COMThread()
    caller_thread = threading.current_thread()
    captured = {}

    def fn():
        captured["thread"] = threading.current_thread()
        return 42

    result = com.run(fn)
    assert result == 42
    assert captured["thread"] is not caller_thread


def test_run_propagates_exceptions():
    com = COMThread()

    def fn():
        raise ValueError("boom")

    with pytest.raises(ValueError, match="boom"):
        com.run(fn)


def test_run_timeout_raises():
    import time
    com = COMThread()

    def fn():
        time.sleep(5)
        return 1

    with pytest.raises(Exception):
        com.run(fn, timeout=0.1)
