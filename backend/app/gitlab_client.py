"""Получение файла из GitLab через REST API v4.

Пользователь может вставить ссылку в любом из привычных форматов:

* ссылка из веб-интерфейса:  https://gitlab.example.com/group/project/-/blob/main/path/file.yml
* raw-ссылка:                https://gitlab.example.com/group/project/-/raw/main/path/file.yml
* готовый вызов API:         https://gitlab.example.com/api/v4/projects/123/repository/files/path%2Ffile.yml/raw?ref=main

Первые два формата автоматически преобразуются в вызов API
``GET /api/v4/projects/:id/repository/files/:file_path/raw?ref=:ref``.
Если ссылка не распознана, она запрашивается как есть (с тем же заголовком
авторизации) — это позволяет работать и с нестандартными раздачами файлов.
"""

import base64
import binascii
import json
import posixpath
import re
from dataclasses import dataclass
from urllib.parse import quote, unquote, urlsplit

import requests

_TIMEOUT = (10, 60)  # секунды: на подключение, на чтение

# Современный формат: /group/subgroup/project/-/blob/<ref>/<path> (или /-/raw/).
# Разделитель "/-/" зарезервирован GitLab и не встречается внутри пути проекта.
_MODERN_BLOB_URL_RE = re.compile(r"^/(?P<project>.+?)/-/(?:blob|raw)/(?P<ref>[^/]+)/(?P<path>.+)$")
# Устаревший формат без "/-/": /group/project/blob/<ref>/<path>
_LEGACY_BLOB_URL_RE = re.compile(r"^/(?P<project>.+?)/(?:blob|raw)/(?P<ref>[^/]+)/(?P<path>.+)$")
# путь к файлу внутри готового вызова API
_API_FILES_RE = re.compile(r"/repository/files/(?P<path>[^/]+)")


class GitLabFetchError(Exception):
    """Ошибка получения файла из GitLab с сообщением, пригодным для показа пользователю."""


@dataclass
class FetchedFile:
    file_name: str
    content: str


def _build_request(url: str) -> tuple[str, str]:
    """Возвращает (адрес для запроса, имя файла) по ссылке, введённой пользователем."""
    parts = urlsplit(url)
    if parts.scheme not in ("http", "https") or not parts.netloc:
        raise GitLabFetchError(
            "Некорректный URL: укажите полный адрес файла, начиная с http:// или https://"
        )

    # Уже готовый вызов API — используем как есть.
    if "/api/v4/" in parts.path:
        match = _API_FILES_RE.search(parts.path)
        raw_path = unquote(match.group("path")) if match else unquote(parts.path)
        return url, posixpath.basename(raw_path) or "file"

    match = _MODERN_BLOB_URL_RE.match(parts.path) or _LEGACY_BLOB_URL_RE.match(parts.path)
    if match:
        project = quote(unquote(match.group("project")).strip("/"), safe="")
        file_path = quote(unquote(match.group("path")), safe="")
        ref = quote(unquote(match.group("ref")), safe="")
        api_url = (
            f"{parts.scheme}://{parts.netloc}/api/v4/projects/{project}"
            f"/repository/files/{file_path}/raw?ref={ref}"
        )
        return api_url, posixpath.basename(unquote(match.group("path")))

    # Не похоже на ссылку GitLab — запрашиваем напрямую.
    return url, posixpath.basename(unquote(parts.path)) or "file"


def _error_details(response: requests.Response) -> str:
    """Достаёт поле message из JSON-ответа GitLab, если оно есть."""
    try:
        payload = response.json()
    except (json.JSONDecodeError, ValueError):
        return ""
    if isinstance(payload, dict):
        return str(payload.get("message") or payload.get("error") or "")
    return ""


def fetch_file(url: str, token: str, verify_ssl: bool = True) -> FetchedFile:
    """Скачивает файл из GitLab и возвращает его текстовое содержимое."""
    request_url, file_name = _build_request(url)

    headers = {}
    if token:
        headers["PRIVATE-TOKEN"] = token

    try:
        response = requests.get(
            request_url, headers=headers, timeout=_TIMEOUT, verify=verify_ssl
        )
    except requests.exceptions.SSLError as exc:
        raise GitLabFetchError(
            "Не удалось проверить SSL-сертификат сервера. Если GitLab использует "
            "самоподписанный сертификат, включите опцию «Не проверять SSL-сертификат»."
        ) from exc
    except requests.exceptions.Timeout as exc:
        raise GitLabFetchError("Сервер GitLab не ответил за отведённое время.") from exc
    except requests.exceptions.RequestException as exc:
        raise GitLabFetchError(f"Не удалось подключиться к {urlsplit(url).netloc}: {exc}") from exc

    if response.status_code in (401,):
        raise GitLabFetchError("GitLab вернул 401: токен не принят. Проверьте токен доступа.")
    if response.status_code == 403:
        raise GitLabFetchError(
            "GitLab вернул 403: у токена недостаточно прав. "
            "Нужен персональный токен со scope read_api (или api)."
        )
    if response.status_code == 404:
        raise GitLabFetchError(
            "GitLab вернул 404: проект, ветка или файл не найдены. "
            "Проверьте URL и права токена (для приватных проектов 404 часто означает «нет доступа»)."
        )
    if not response.ok:
        details = _error_details(response)
        suffix = f" ({details})" if details else ""
        raise GitLabFetchError(f"GitLab вернул ошибку {response.status_code}{suffix}.")

    # Эндпоинт /repository/files/:path без /raw возвращает JSON с base64-содержимым.
    if "json" in response.headers.get("content-type", "").lower():
        try:
            payload = response.json()
        except (json.JSONDecodeError, ValueError):
            payload = None
        if isinstance(payload, dict) and payload.get("encoding") == "base64":
            try:
                content = base64.b64decode(payload.get("content", "")).decode(
                    "utf-8", errors="replace"
                )
            except (binascii.Error, ValueError) as exc:
                raise GitLabFetchError("Не удалось декодировать содержимое файла из ответа API.") from exc
            return FetchedFile(str(payload.get("file_name") or file_name), content)

    response.encoding = response.encoding or "utf-8"
    return FetchedFile(file_name, response.text)
