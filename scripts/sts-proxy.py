"""
HTTP-proxy for STS API (pos.autooplata.ru/tms)
1C connects to 127.0.0.1:8099 (HTTP), proxy forwards via curl (Schannel).
"""

import http.server
import socketserver
import subprocess
import tempfile
import os

PORT = 8099
TARGET = "https://pos.autooplata.ru/tms"
LOG = open("D:/Users/magsp/ELSYPLUS/NefteUchet/scripts/proxy.log", "w", encoding="utf-8")

def log(msg):
    print(msg, flush=True)
    LOG.write(msg + "\n")
    LOG.flush()

class Handler(http.server.BaseHTTPRequestHandler):

    def do(self):
        url = TARGET + self.path
        body = None
        length = self.headers.get("Content-Length")
        if length:
            body = self.rfile.read(int(length))

        log(f">>> {self.command} {url}")

        # Write response to temp file to avoid stdout parsing issues
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".tmp")
        tmp_headers = tempfile.NamedTemporaryFile(delete=False, suffix=".hdr")
        tmp.close()
        tmp_headers.close()

        cmd = ["curl", "-s", "--max-time", "30",
               "-X", self.command,
               "-H", "Content-Type: application/json",
               "-o", tmp.name,
               "-D", tmp_headers.name,
               url]

        auth = self.headers.get("Authorization")
        if auth:
            cmd.extend(["-H", f"Authorization: {auth}"])

        if body:
            cmd.extend(["--data-binary", "@-"])

        try:
            result = subprocess.run(
                cmd, input=body, capture_output=True, timeout=35)

            # Read response body
            with open(tmp.name, "rb") as f:
                data = f.read()

            # Parse status code from headers
            status_code = 200
            try:
                with open(tmp_headers.name, "r") as f:
                    first_line = f.readline()  # HTTP/1.1 200 OK
                    parts = first_line.split(" ", 2)
                    if len(parts) >= 2:
                        status_code = int(parts[1])
            except Exception:
                pass

            log(f"<<< {status_code} [{len(data)} bytes]")

            self.send_response(status_code)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        except Exception as e:
            log(f"<<< ERROR: {e}")
            msg = f"Proxy error: {e}".encode()
            self.send_response(502)
            self.send_header("Content-Length", str(len(msg)))
            self.end_headers()
            self.wfile.write(msg)
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass
            try:
                os.unlink(tmp_headers.name)
            except Exception:
                pass

    do_GET = do_POST = do_PUT = do_PATCH = do_DELETE = do

    def log_message(self, fmt, *args):
        pass

class ThreadedServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True

if __name__ == "__main__":
    log(f"STS Proxy (curl): http://127.0.0.1:{PORT} -> {TARGET}")
    ThreadedServer(("", PORT), Handler).serve_forever()
