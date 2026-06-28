import queue
import threading
from concurrent.futures import Future

import pythoncom


class COMThread:
    """Dedicated STA thread for all COM interactions."""

    def __init__(self):
        self._queue = queue.Queue()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _loop(self):
        pythoncom.CoInitialize()
        try:
            while True:
                item = self._queue.get()
                if item is None:
                    break
                fn, future = item
                try:
                    future.set_result(fn())
                except Exception as exc:
                    future.set_exception(exc)
        finally:
            pythoncom.CoUninitialize()

    def run(self, fn, timeout=30):
        """Submit fn to the COM thread and block until it returns.

        NOTE: On timeout, the caller gets TimeoutError but the COM thread
        continues running fn. If Revu hangs (e.g. on large Flatten/Export),
        all subsequent calls will queue behind the hung call until it returns
        or the server process is restarted.
        """
        future = Future()
        self._queue.put((fn, future))
        return future.result(timeout=timeout)

    def stop(self):
        self._queue.put(None)
        self._thread.join()
