# GitLab YAML → XLSX

Мини-приложение: скачивает YAML-файл из GitLab через REST API и преобразует его
в таблицу Excel (`.xlsx`) с сохранением всех уровней вложенности.

## Как это работает

1. В форме указываются ссылка на файл в GitLab и токен доступа.
2. Бэкенд превращает ссылку в вызов `GET /api/v4/projects/:id/repository/files/:path/raw?ref=:ref`
   и скачивает файл с заголовком `PRIVATE-TOKEN`.
3. YAML преобразуется в таблицу без потери уровней вложенности. Для файлов вида
   «каталог записей» (как `nginx_http_upstreams`: словарь записей с полями и
   вложенными группами `main`/`backup`) строится блочная таблица: имя записи в
   первой колонке, поля по колонкам, группы — с двухуровневой шапкой, списки
   растягиваются вниз, блоки разделены пустой строкой. Для YAML любой другой
   структуры строится универсальная таблица: одна строка на каждое конечное
   значение, колонки «Уровень 1..N» — путь до него.
4. Готовый `.xlsx` (лист с данными + лист «Инфо») скачивается браузером.

## Запуск

```bash
docker compose up --build
```

Интерфейс: <http://localhost:9000> (порт меняется в `docker-compose.yml`).

## Использование

* **URL файла** — подойдёт любой из форматов:
  * `https://gitlab.example.com/group/project/-/blob/main/path/file.yml` (ссылка из браузера)
  * `https://gitlab.example.com/group/project/-/raw/main/path/file.yml`
  * `https://gitlab.example.com/api/v4/projects/123/repository/files/path%2Ffile.yml/raw?ref=main`
* **Токен** — персональный токен GitLab (Settings → Access Tokens) со scope
  `read_api` или `api`. Для публичных проектов поле можно оставить пустым.
* **Не проверять SSL-сертификат** — включите, если GitLab использует
  самоподписанный сертификат.

Ограничение: если имя ветки содержит `/` (например, `feature/x`), ссылка из
веб-интерфейса неоднозначна — используйте формат прямого вызова API.

## Структура проекта

```
backend/            FastAPI: загрузка из GitLab + конвертация в XLSX
  app/gitlab_client.py   разбор ссылки и запрос в GitLab API
  app/converter.py       YAML → XLSX (openpyxl)
  app/main.py            HTTP API (POST /api/convert)
frontend/           статическая страница + nginx (проксирует /api на бэкенд)
docker-compose.yml  два сервиса: backend и frontend (порт 9000)
tests/mock_gitlab.py     мок GitLab API для локальной проверки
```

## Проверка без реального GitLab

```bash
python3 tests/mock_gitlab.py          # мок GitLab API на порту 9999
docker compose up --build
```

В форме укажите URL
`http://host.docker.internal:9999/api/v4/projects/1/repository/files/skdpu_http_upstreams.yml/raw?ref=main`
и токен `test-token`.
