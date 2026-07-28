"""Мок-сервер GitLab API для проверки приложения без реального GitLab.

Отдаёт YAML-файл на любой запрос вида ``.../repository/files/<путь>/raw`` и
требует заголовок ``PRIVATE-TOKEN: test-token`` — так проверяется и передача
токена приложением.

Запуск:  python3 tests/mock_gitlab.py [порт] [путь_к_yaml]
По умолчанию порт 9999 и файл skdpu_http_upstreams.yml из корня репозитория.

Пример URL для формы приложения (изнутри контейнера хост доступен как
host.docker.internal):
    http://host.docker.internal:9999/api/v4/projects/1/repository/files/skdpu_http_upstreams.yml/raw?ref=main
Токен: test-token
"""

import http.server
import json
import pathlib
import sys

TOKEN = "test-token"
DEFAULT_FILE = pathlib.Path(__file__).resolve().parent.parent / "skdpu_http_upstreams.yml"


class Handler(http.server.BaseHTTPRequestHandler):
    yaml_bytes = b""

    def do_GET(self):
        if self.headers.get("PRIVATE-TOKEN") != TOKEN:
            self._send(401, {"message": "401 Unauthorized"})
        elif "/repository/files/" in self.path and "/raw" in self.path:
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(self.yaml_bytes)))
            self.end_headers()
            self.wfile.write(self.yaml_bytes)
        else:
            self._send(404, {"message": "404 Not Found"})

    def _send(self, status, payload):
        body = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        sys.stderr.write("mock-gitlab: %s\n" % (fmt % args))


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 9999
    file_path = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_FILE
    if not file_path.is_file():
        sys.exit(
            f"mock-gitlab: файл {file_path} не найден. "
            "Укажите путь к YAML-файлу вторым аргументом: "
            "python3 tests/mock_gitlab.py 9999 /путь/к/файлу.yml"
        )
    Handler.yaml_bytes = file_path.read_bytes()
    server = http.server.HTTPServer(("0.0.0.0", port), Handler)
    print(f"mock-gitlab: слушаю на :{port}, отдаю {file_path} (токен: {TOKEN})")
    server.serve_forever()


if __name__ == "__main__":
    main()
