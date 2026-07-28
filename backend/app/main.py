"""HTTP API мини-приложения «YAML из GitLab → XLSX»."""

import logging
import re
from urllib.parse import quote

from fastapi import FastAPI
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel

from .converter import YamlConversionError, convert_yaml_to_xlsx
from .gitlab_client import GitLabFetchError, fetch_file

logger = logging.getLogger("app")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

app = FastAPI(title="GitLab YAML → XLSX", docs_url=None, redoc_url=None)

_XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class ConvertRequest(BaseModel):
    url: str
    token: str = ""
    verify_ssl: bool = True


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/api/convert")
def convert(request: ConvertRequest) -> Response:
    url = request.url.strip()
    if not url:
        return JSONResponse(status_code=400, content={"detail": "Укажите URL файла в GitLab."})

    logger.info("Запрос конвертации: %s (verify_ssl=%s)", url, request.verify_ssl)
    try:
        fetched = fetch_file(url, request.token.strip(), verify_ssl=request.verify_ssl)
        xlsx_bytes, row_count, root_key = convert_yaml_to_xlsx(
            fetched.content, source_url=url, file_name=fetched.file_name
        )
    except GitLabFetchError as exc:
        logger.warning("Ошибка загрузки из GitLab: %s", exc)
        return JSONResponse(status_code=502, content={"detail": str(exc)})
    except YamlConversionError as exc:
        logger.warning("Ошибка конвертации: %s", exc)
        return JSONResponse(status_code=422, content={"detail": str(exc)})

    base_name = re.sub(r"\.(ya?ml)$", "", fetched.file_name, flags=re.IGNORECASE) or "converted"
    download_name = f"{base_name}.xlsx"
    logger.info("Готово: %s, строк данных: %d", download_name, row_count)

    headers = {
        "Content-Disposition": (
            f"attachment; filename=\"converted.xlsx\"; filename*=UTF-8''{quote(download_name)}"
        ),
        "X-Row-Count": str(row_count),
    }
    if root_key:
        headers["X-Root-Key"] = quote(root_key)  # URL-кодирование: в HTTP-заголовках только ASCII

    return Response(content=xlsx_bytes, media_type=_XLSX_MEDIA_TYPE, headers=headers)
