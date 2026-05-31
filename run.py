import ctypes
import socket as _socket
import sys
from socketserver import ThreadingMixIn
from wsgiref.simple_server import WSGIServer, WSGIRequestHandler

sys.path.insert(0, ".")
from api.index import app

PORT = 5000


def _raw_listen(sock_obj: _socket.socket, backlog: int = 5) -> bool:
    """
    Try multiple ctypes strategies to call ws2_32.listen().
    Returns True on success, False if all fail (caller may ignore and
    let Python's own socket.listen() run as a last-ditch attempt).
    """
    handle = sock_obj.fileno()
    print(f"[run.py] socket.fileno() = {handle}  (type: {type(handle).__name__})")

    # ws2_32 might alias a socket fd rather than the raw SOCKET on Python 3.14.
    # Try to recover the real Windows HANDLE via the CRT.
    try:
        msvcrt = ctypes.CDLL("msvcrt")
        msvcrt._get_osfhandle.restype = ctypes.c_size_t
        msvcrt._get_osfhandle.argtypes = [ctypes.c_int]
        os_handle = msvcrt._get_osfhandle(handle)
        print(f"[run.py] _get_osfhandle({handle}) = {os_handle}")
    except Exception as e:
        print(f"[run.py] _get_osfhandle failed: {e}")
        os_handle = None

    candidates = [handle]
    if os_handle and os_handle != handle and os_handle not in (0, 0xFFFFFFFFFFFFFFFF):
        candidates.insert(0, int(os_handle))

    ws2 = ctypes.WinDLL("ws2_32")
    ws2.WSAGetLastError.restype = ctypes.c_int

    for raw in candidates:
        # Strategy A: c_size_t (pointer-sized, safest for SOCKET on x64)
        try:
            f = ctypes.WinDLL("ws2_32")
            f.listen.argtypes = [ctypes.c_size_t, ctypes.c_int]
            f.listen.restype = ctypes.c_int
            rc = f.listen(ctypes.c_size_t(raw), ctypes.c_int(backlog))
            if rc == 0:
                print(f"[run.py] listen OK via c_size_t handle={raw}")
                return True
            print(f"[run.py] c_size_t({raw}) -> rc={rc}, WSA={f.WSAGetLastError()}")
        except Exception as e:
            print(f"[run.py] c_size_t attempt error: {e}")

        # Strategy B: no argtypes (Python int -> default c_int, works if handle < 2^31)
        try:
            f2 = ctypes.WinDLL("ws2_32")
            f2.listen.restype = ctypes.c_int
            rc2 = f2.listen(raw, backlog)
            if rc2 == 0:
                print(f"[run.py] listen OK via no-argtypes handle={raw}")
                return True
            print(f"[run.py] no-argtypes({raw}) -> rc={rc2}, WSA={f2.WSAGetLastError()}")
        except Exception as e:
            print(f"[run.py] no-argtypes attempt error: {e}")

    return False


class _Handler(WSGIRequestHandler):
    protocol_version = "HTTP/1.0"

    def log_message(self, fmt, *args):
        sys.stdout.write(f"{self.address_string()} - {fmt % args}\n")
        sys.stdout.flush()


class _Server(ThreadingMixIn, WSGIServer):
    daemon_threads = True

    def server_activate(self):
        ok = _raw_listen(self.socket)
        if not ok:
            # Last resort: let Python try its own path (may raise WinError 10014)
            print("[run.py] All ctypes strategies failed — falling back to socket.listen()")
            self.socket.listen(self.request_queue_size)


server = _Server(("", PORT), _Handler)
server.set_app(app)

print(f"서버 실행 중: http://localhost:{PORT}")
print("종료하려면 Ctrl+C 를 누르세요.")
try:
    server.serve_forever()
except KeyboardInterrupt:
    pass
