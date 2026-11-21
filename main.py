
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote_plus, urlparse, urlunparse, parse_qsl, urlencode, urlsplit, urlunsplit

import feedparser
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt
from googlenewsdecoder import gnewsdecoder


# ============================================================
# CONSTANTS
# ============================================================

CONFIG_FILE = Path("config.json")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 Chrome/120.0 Safari/537.36"
)

TRACKING_PARAMS = {
    "utm_source",
    "utm_medium",
    "utm_campaign",
    "utm_term",
    "utm_content",
    "fbclid",
    "gclid",
}


# ============================================================
# CONFIG
# ============================================================

def load_config():
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"Не найден файл конфигурации: {CONFIG_FILE}"
        )

    with CONFIG_FILE.open("r", encoding="utf-8") as file:
        config = json.load(file)

    # Тема может быть пустой.
    topic = config.get("topic", "")

    if not isinstance(topic, str):
        raise ValueError("topic должен быть строкой.")

    config["topic"] = topic.strip()

    # Ключевые слова могут отсутствовать или быть пустыми.
    keywords = config.get("keywords", [])

    if keywords is None:
        keywords = []

    if not isinstance(keywords, list):
        raise ValueError("keywords должен быть списком.")

    config["keywords"] = [
        keyword.strip()
        for keyword in keywords
        if isinstance(keyword, str) and keyword.strip()
    ]

    # Остальные настройки.
    if not isinstance(config.get("days_back"), int):
        raise ValueError("days_back должен быть целым числом.")

    if config["days_back"] < 1:
        raise ValueError("days_back должен быть не меньше 1.")

    if not isinstance(config.get("max_articles"), int):
        raise ValueError("max_articles должен быть целым числом.")

    if config["max_articles"] < 1:
        raise ValueError("max_articles должен быть не меньше 1.")

    if not config.get("feeds"):
        raise ValueError("В config.json должен быть хотя бы один RSS-источник.")

    if not isinstance(config.get("llm"), dict):
        raise ValueError("Не найдена конфигурация llm.")

    return config


# ============================================================
# TEXT AND URL HELPERS
# ============================================================

def normalize_text(value):
    value = value or ""
    value = value.lower()
    value = re.sub(r"\s+", " ", value)
    value = re.sub(r"[^\w\s]", "", value)
    return value.strip()

def get_article_source_url(article: dict) -> str:
    return (
        article.get("resolved_url")
        or article.get("url")
        or ""
    ).strip()

def parse_date(value):
    """
    Преобразует дату RSS в datetime с UTC.
    """

    if not value:
        return None

    if hasattr(value, "tm_year"):
        try:
            return datetime(
                value.tm_year,
                value.tm_mon,
                value.tm_mday,
                value.tm_hour,
                value.tm_min,
                value.tm_sec,
                tzinfo=timezone.utc
            )
        except (ValueError, TypeError):
            return None

    try:
        parsed = datetime.fromisoformat(
            str(value).replace("Z", "+00:00")
        )

        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)

        return parsed.astimezone(timezone.utc)

    except (ValueError, TypeError):
        return None


def extract_date_from_html(soup):
    """
    Пытается найти дату публикации в HTML статьи.
    """

    selectors = [
        ("meta", {"property": "article:published_time"}),
        ("meta", {"name": "pubdate"}),
        ("meta", {"name": "date"}),
        ("meta", {"property": "og:updated_time"}),
    ]

    for tag_name, attrs in selectors:
        tag = soup.find(tag_name, attrs=attrs)

        if tag and tag.get("content"):
            result = parse_date(tag["content"])

            if result:
                return result

    return None


# ============================================================
# RSS SOURCES
# ============================================================

def build_feed_list(config):
    """
    Формирует список RSS-источников.

    Если тема и keywords пустые, Google News ищет новости
    по общему запросу Texas.
    """

    feeds = list(config.get("feeds", []))

    if not config.get("google_news", False):
        return feeds

    topic = config["topic"]
    keywords = config["keywords"]

    queries = []

    if topic:
        queries.append(topic)

    queries.extend(keywords)

    # Общий режим: нет тематических запросов.
    if not queries:
        queries = ["Texas"]

    # Убираем дубликаты запросов.
    queries = list(dict.fromkeys(queries))

    days_back = config["days_back"]

    for query in queries:
        search_query = f"{query} when:{days_back}d"

        url = (
            "https://news.google.com/rss/search?q="
            f"{quote_plus(search_query)}"
            "&hl=en-US&gl=US&ceid=US:en"
        )

        feeds.append({
            "name": f"Google News: {query}",
            "url": url
        })

    return feeds


def search_rss(feed_url, source_name, cutoff, now):
    """
    Загружает RSS и возвращает новости за заданный период.
    """

    try:
        response = requests.get(
            feed_url,
            headers={"User-Agent": USER_AGENT},
            timeout=20
        )

        response.raise_for_status()

        feed = feedparser.parse(response.content)

    except requests.RequestException as error:
        print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось загрузить {source_name}: {error}")
        return []

    articles = []

    for entry in feed.entries:
        title = entry.get("title", "").strip()
        link = entry.get("link", "").strip()

        if not title or not link:
            continue

        published = (
            entry.get("published_parsed")
            or entry.get("updated_parsed")
            or entry.get("published")
            or entry.get("updated")
        )

        published_date = parse_date(published)

        # Без даты невозможно надёжно проверить период.
        if not published_date:
            continue

        if published_date < cutoff or published_date > now:
            continue

        summary = entry.get("summary", "")

        if summary:
            summary_soup = BeautifulSoup(summary, "html.parser")
            summary = summary_soup.get_text(" ", strip=True)

        author = (
            entry.get("author")
            or entry.get("dc_creator")
            or "Не указан"
        )

        articles.append({
            "title": title,
            "url": normalize_url(link),
            "source": source_name,
            "author": author,
            "published": published_date,
            "summary": summary,
            "text": ""
        })

    return articles


# ============================================================
# ARTICLE EXTRACTION
# ============================================================


import trafilatura


# def get_article_text(url, timeout, max_chars):
#     """
#     Загружает страницу и извлекает основной текст публикации.

#     Возвращает:
#         {
#             "text": str,
#             "final_url": str,
#             "status": str
#         }
#     """

#     result = {
#         "text": "",
#         "final_url": url,
#         "status": "failed"
#     }

#     try:
#         response = requests.get(
#             url,
#             headers={
#                 "User-Agent": USER_AGENT,
#                 "Accept": (
#                     "text/html,application/xhtml+xml,"
#                     "application/xml;q=0.9,*/*;q=0.8"
#                 ),
#                 "Accept-Language": "en-US,en;q=0.9"
#             },
#             timeout=timeout,
#             allow_redirects=True
#         )

#         response.raise_for_status()

#         result["final_url"] = response.url

#         content_type = response.headers.get(
#             "Content-Type", ""
#         ).lower()

#         if "html" not in content_type:
#             result["status"] = "not_html"
#             return result

#         html = response.text

#         # 1. Основной способ — Trafilatura.
#         extracted = trafilatura.extract(
#             html,
#             include_comments=False,
#             include_tables=True,
#             include_links=False,
#             favor_precision=True
#         )

#         if extracted:
#             extracted = re.sub(
#                 r"\s+",
#                 " ",
#                 extracted
#             ).strip()

#         # 2. Резервный способ — BeautifulSoup.
#         if not extracted or len(extracted) < 300:
#             soup = BeautifulSoup(html, "html.parser")

#             for element in soup([
#                 "script",
#                 "style",
#                 "nav",
#                 "footer",
#                 "header",
#                 "aside",
#                 "form",
#                 "noscript",
#                 "iframe"
#             ]):
#                 element.decompose()

#             article = (
#                 soup.find("article")
#                 or soup.find("main")
#                 or soup.body
#                 or soup
#             )

#             extracted = article.get_text(
#                 " ",
#                 strip=True
#             )

#             extracted = re.sub(
#                 r"\s+",
#                 " ",
#                 extracted
#             ).strip()

#         # Не принимаем заголовок или короткое описание
#         # за полноценный текст публикации.
#         if len(extracted) < 300:
#             result["text"] = extracted
#             result["status"] = "insufficient_text"
#             return result

#         result["text"] = extracted[:max_chars]
#         result["status"] = "success"

#         return result

#     except requests.Timeout:
#         result["status"] = "timeout"
#         return result

#     except requests.HTTPError as error:
#         result["status"] = f"http_error: {error}"
#         return result

#     except requests.RequestException as error:
#         result["status"] = f"request_error: {error}"
#         return result

#     except Exception as error:
#         result["status"] = f"extraction_error: {error}"
#         return result

import requests
import trafilatura

from bs4 import BeautifulSoup
from trafilatura import extract, html2txt

def resolve_google_news_url(url: str) -> str | None:
    if "news.google.com" not in url:
        return url

    try:
        result = gnewsdecoder(
            url,
            interval=1,
            timeout=20,
        )

        if result.get("success"):
            return result["decoded_url"]

        print(f"Не удалось разрешить Google News URL: {result.get('message')}")
        return None

    except Exception as exc:
        print(f"Ошибка декодирования Google News URL: {exc}")
        return None


def get_article_text(url: str, timeout: int = 20, max_chars: int = 8000):

    original_url = url

    if "news.google.com" in url:
        resolved_url = resolve_google_news_url(url)

        if not resolved_url:
            return {
                "text": "",
                "final_url": url,
                "status": "url_resolution_failed",
            }

        url = resolved_url

    print(f"Исходный URL: {original_url}")
    print(f"URL издателя: {url}")
    
    result = {
        "text": "",
        "final_url": url,
        "status": "unknown",
    }

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/131.0.0.0 Safari/537.36"
        ),
        "Accept": (
            "text/html,application/xhtml+xml,"
            "application/xml;q=0.9,*/*;q=0.8"
        ),
        "Accept-Language": "en-US,en;q=0.9",
    }


    try:
        response = requests.get(
            url,
            headers=headers,
            timeout=timeout,
            allow_redirects=True,
        )

        result["final_url"] = response.url

        print(f"\nURL: {url}")
        print(f"HTTP: {response.status_code}")
        print(f"Final URL: {response.url}")
        print(f"Content-Type: {response.headers.get('Content-Type')}")
        print(f"HTML size: {len(response.text)}")

        response.raise_for_status()

        html = response.text

        # Способ 1: стандартное извлечение
        text = extract(html, url=response.url)

        print(f"Trafilatura standard: {len(text or '')}")

        # Способ 2: более мягкое извлечение
        if not text or len(text.strip()) < 300:
            text = extract(
                html,
                url=response.url,
                favor_recall=True,
            )

        print(f"Trafilatura recall: {len(text or '')}")

        # Способ 3: извлечение всего текста HTML
        if not text or len(text.strip()) < 300:
            text = html2txt(html)

        print(f"HTML2TXT: {len(text or '')}")

        # Способ 4: BeautifulSoup
        if not text or len(text.strip()) < 300:
            soup = BeautifulSoup(html, "html.parser")

            for element in soup([
                "script", "style", "nav",
                "footer", "header", "noscript"
            ]):
                element.decompose()

            text = soup.get_text(
                separator="\n",
                strip=True,
            )

        text = (text or "").strip()

        result["text"] = text[:max_chars]

        if len(text) >= 300:
            result["status"] = "success"
        else:
            result["status"] = "insufficient_text"

        print(f"Final extracted text: {len(text)}")
        print(f"Status: {result['status']}")

        return result

    except requests.exceptions.Timeout:
        result["status"] = "timeout"
        print("Request timeout")
        return result

    except requests.exceptions.HTTPError as exc:
        result["status"] = f"http_error_{response.status_code}"
        print(f"HTTP error: {exc}")
        return result

    except requests.exceptions.RequestException as exc:
        result["status"] = "request_error"
        print(f"Request error: {exc}")
        return result

    except Exception as exc:
        result["status"] = "extraction_error"
        print(f"Extraction error: {type(exc).__name__}: {exc}")
        return result
# ============================================================
# DEDUPLICATION AND COLLECTION
# ============================================================

def is_duplicate(article, existing_articles):
    """
    Проверяет дубликаты по URL и заголовку.
    """

    current_url = normalize_url(article["url"])
    current_title = normalize_text(article["title"])

    for existing in existing_articles:
        if current_url and current_url == normalize_url(existing["url"]):
            return True

        existing_title = normalize_text(existing["title"])

        if current_title and current_title == existing_title:
            return True

    return False


def collect_articles(config):
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(days=config["days_back"])

    feeds = build_feed_list(config)

    print(f"\\nПериод: последние {config['days_back']} дн.")
    print(f"RSS-источников: {len(feeds)}")
    print(
        "Режим поиска:",
        "общий мониторинг"
        if not config["topic"] and not config["keywords"]
        else "тематический поиск"
    )

    all_articles = []

    for feed_config in feeds:
        name = feed_config["name"]
        url = feed_config["url"]

        print(f"\\n[RSS] {name}")

        articles = search_rss(
            feed_url=url,
            source_name=name,
            cutoff=cutoff,
            now=now
        )

        print(f"  Найдено: {len(articles)}")
        all_articles.extend(articles)

    all_articles.sort(
        key=lambda article: article["published"],
        reverse=True
    )

    unique_articles = []
    for article in all_articles:
        if not is_duplicate(article, unique_articles):
            unique_articles.append(article)

    print(f"\\nУникальных новостей: {len(unique_articles)}")

    unique_articles = unique_articles[:config["max_articles"]]

    print("\\nЗагрузка текстов статей...")

    for index, article in enumerate(unique_articles, start=1):
        print(
            f"[{index}/{len(unique_articles)}] "
            f"{article['title'][:80]}"
        )

        original_url = article["url"]

        result = get_article_text(
            url=original_url,
            timeout=config["request_timeout"],
            max_chars=config["max_chars_per_article"]
        )

        article["text"] = result["text"]
        article["resolved_url"] = result["final_url"] or original_url
        article["extraction_status"] = result["status"]

        # Для успешного извлечения используем прямую ссылку издателя.
        # Исходную ссылку сохраняем отдельно для трассировки.
        if (
            result["status"] == "success"
            and result["final_url"]
            and "news.google.com" not in urlparse(
                result["final_url"]
            ).netloc.lower()
        ):
            article["original_url"] = original_url
            article["url"] = result["final_url"]

        print(f"  Статус извлечения: {article['extraction_status']}")
        print(f"  Длина текста: {len(article['text'])} символов")
        print(
            f"  URL источника: {get_article_source_url(article)}"
        )

        time.sleep(0.3)

    return unique_articles


# ============================================================
# LLM ANALYSIS
# ============================================================

def build_instructions(config):
    """
    Формирует инструкции для LLM.
    """

    topic = config["topic"]
    keywords = config["keywords"]

    if topic or keywords:
        search_mode = (
            f"Тема: {topic or 'не задана'}\n"
            f"Ключевые слова: {', '.join(keywords) or 'не заданы'}"
        )
    else:
        search_mode = (
            "Режим общего мониторинга: тематические ограничения "
            "отсутствуют. Проанализируй все предоставленные новости."
        )

    return f"""
Ты аналитик открытых источников (OSINT).

Твоя задача — подготовить аналитическую справку по новостям.

{search_mode}

Правила анализа:

1. Используй только предоставленные новости.
2. Не выдумывай факты, даты, цитаты и источники.
3. Отделяй подтверждаемые факты от мнений и оценок.
4. Если информации недостаточно, прямо укажи это.
5. Не представляй предположения как установленные факты.
6. Не делай выводов, которые не подтверждаются источниками.
7. Для каждого существенного вывода указывай source_urls.
8. Если новость содержит мнение автора или цитату эксперта,
   обозначай это как мнение или позицию.
9. Не смешивай разные события в одно.
10. Если источники противоречат друг другу, укажи противоречие.

КРИТИЧЕСКИ ВАЖНО:

1. Анализируй содержание поля text каждой статьи.
2. Не делай содержательных выводов только на основании title.
3. Если текст статьи отсутствует или недостаточен,
   не пытайся реконструировать содержание по заголовку.
4. Не выдавай предположения за факты.
5. Не анализируй статью, если её text содержит только заголовок
   или короткое описание из RSS.
6. Учитывай дату публикации из переданных метаданных.
7. Не придумывай текущую дату и не изменяй даты источников.

Верни только JSON-объект.
Не используй Markdown.
Не оборачивай JSON в ```json или другие блоки кода.
Первый символ ответа должен быть открывающей фигурной скобкой.
Последний символ ответа должен быть закрывающей фигурной скобкой.
Не добавляй пояснений до или после JSON.

Верни результат строго в формате JSON:

{{
  "executive_summary": "Краткая сводка",
  "key_findings": [
    {{
      "statement": "Вывод",
      "source_urls": ["https://example.com"]
    }}
  ],
  "facts": [
    {{
      "statement": "Установленный факт",
      "source_urls": ["https://example.com"]
    }}
  ],
  "opinions": [
    {{
      "statement": "Мнение или оценка",
      "source_urls": ["https://example.com"]
    }}
  ],
  "potential_tasks": [
    {{
      "task": "Возможная задача или направление для проверки",
      "rationale": "Почему это может быть важно",
      "source_urls": ["https://example.com"]
    }}
  ],
  "limitations": [
    "Ограничение анализа"
  ],
  "source_urls": [
    "https://example.com"
  ]
}}

Если категория не содержит данных, верни пустой массив.
"""


def extract_response_text(response_data):
    """
    Извлекает текст из ответа Responses API.
    """

    if isinstance(response_data.get("output_text"), str):
        return response_data["output_text"]

    parts = []

    for output_item in response_data.get("output", []):
        for content_item in output_item.get("content", []):
            if content_item.get("type") in ("output_text", "text"):
                text = content_item.get("text")

                if text:
                    parts.append(text)

    if not parts:
        raise ValueError("LLM вернула ответ без текстового содержимого.")

    return "\n".join(parts)


def parse_json_response(raw_text: str, debug_dir: Path | None = None):
    """
    Строго разбирает JSON-ответ модели.

    Не пытается автоматически исправлять повреждённый JSON:
    такое исправление может незаметно изменить смысл анализа.
    При ошибке сохраняет исходный ответ для диагностики.
    """
    if not isinstance(raw_text, str) or not raw_text.strip():
        raise ValueError("LLM вернула пустой ответ вместо JSON.")

    cleaned_text = raw_text.strip()

    # Убираем только внешнюю markdown-обёртку.
    if cleaned_text.startswith("```"):
        cleaned_text = re.sub(
            r"^```(?:json)?\\s*",
            "",
            cleaned_text,
            count=1,
            flags=re.IGNORECASE
        )
        cleaned_text = re.sub(r"\\s*```$", "", cleaned_text, count=1)

    try:
        parsed = json.loads(cleaned_text)
    except json.JSONDecodeError as exc:
        debug_path = None

        if debug_dir is not None:
            try:
                debug_dir.mkdir(parents=True, exist_ok=True)
                debug_path = debug_dir / (
                    f"llm_invalid_json_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                )
                debug_path.write_text(raw_text, encoding="utf-8")
            except OSError:
                debug_path = None

        start = max(0, exc.pos - 250)
        end = min(len(cleaned_text), exc.pos + 250)
        fragment = cleaned_text[start:end]

        message = (
            "Не удалось разобрать JSON-ответ LLM. "
            f"Строка: {exc.lineno}, колонка: {exc.colno}, "
            f"позиция: {exc.pos}. "
            f"Фрагмент около ошибки: {fragment!r}"
        )

        if debug_path:
            message += f"\\nПолный ответ сохранён: {debug_path.resolve()}"

        raise ValueError(message) from exc

    if not isinstance(parsed, dict):
        raise ValueError(
            "Корень JSON-ответа должен быть объектом, "
            f"получен тип: {type(parsed).__name__}."
        )

    return parsed


def normalize_url(url: str) -> str:
    """
    Нормализует URL для сравнения ссылок.

    Удаляет www, fragment и query-параметры.
    Важно: query-параметры могут быть частью маршрута некоторых сайтов,
    поэтому здесь они удаляются только для сравнения источников.
    """
    if not isinstance(url, str) or not url.strip():
        return ""

    try:
        parsed = urlsplit(url.strip())
        scheme = parsed.scheme.lower()
        hostname = (parsed.hostname or "").lower()

        if not scheme or not hostname:
            return ""

        if hostname.startswith("www."):
            hostname = hostname[4:]

        port = parsed.port
        netloc = hostname

        if port and not (
            (scheme == "https" and port == 443)
            or (scheme == "http" and port == 80)
        ):
            netloc = f"{hostname}:{port}"

        path = parsed.path.rstrip("/") or "/"

        return urlunsplit((scheme, netloc, path, "", ""))

    except (TypeError, ValueError):
        return url.strip()


def validate_analysis(analysis, articles):
    """
    Проверяет структуру ответа LLM и все ссылки на источники.
    В качестве допустимых ссылок учитывает исходный и разрешённый URL статьи.
    """
    if not isinstance(analysis, dict):
        raise ValueError("Результат анализа должен быть JSON-объектом.")

    required_fields = (
        "executive_summary",
        "key_findings",
        "facts",
        "opinions",
        "potential_tasks",
        "limitations",
        "source_urls",
    )

    for field in required_fields:
        if field not in analysis:
            raise ValueError(f"В анализе отсутствует поле: {field}")

    if not isinstance(analysis["executive_summary"], str):
        raise ValueError("executive_summary должен быть строкой.")

    known_urls = set()
    for article in articles:
        for candidate in (
            article.get("url", ""),
            article.get("resolved_url", ""),
            article.get("original_url", ""),
        ):
            normalized = normalize_url(candidate)
            if normalized:
                known_urls.add(normalized)

    def validate_source_list(items, location):
        if not isinstance(items, list):
            raise ValueError(
                f"Поле source_urls должно быть списком: {location}"
            )

        for url in items:
            if not isinstance(url, str) or not url.strip():
                raise ValueError(
                    f"Некорректная ссылка в {location}: {url!r}"
                )

            normalized = normalize_url(url)
            if not normalized or normalized not in known_urls:
                raise ValueError(
                    f"LLM указала неизвестный источник: {url}"
                )

    validate_source_list(analysis["source_urls"], "source_urls")

    for field in ("key_findings", "facts", "opinions"):
        if not isinstance(analysis[field], list):
            raise ValueError(f"{field} должен быть списком.")

        for index, item in enumerate(analysis[field], start=1):
            if not isinstance(item, dict):
                raise ValueError(
                    f"Элемент {index} в {field} должен быть объектом."
                )

            if not isinstance(item.get("statement"), str):
                raise ValueError(
                    f"В элементе {index} поля {field} отсутствует "
                    "строковое поле statement."
                )

            validate_source_list(
                item.get("source_urls", []),
                f"{field}[{index}].source_urls"
            )

    if not isinstance(analysis["potential_tasks"], list):
        raise ValueError("potential_tasks должен быть списком.")

    for index, item in enumerate(analysis["potential_tasks"], start=1):
        if not isinstance(item, dict):
            raise ValueError(
                f"Элемент {index} в potential_tasks должен быть объектом."
            )

        if not isinstance(item.get("task"), str):
            raise ValueError(
                f"В potential_tasks[{index}] отсутствует строковое поле task."
            )

        if not isinstance(item.get("rationale", ""), str):
            raise ValueError(
                f"Поле rationale в potential_tasks[{index}] должно быть строкой."
            )

        validate_source_list(
            item.get("source_urls", []),
            f"potential_tasks[{index}].source_urls"
        )

    if not isinstance(analysis["limitations"], list) or not all(
        isinstance(item, str) for item in analysis["limitations"]
    ):
        raise ValueError("limitations должен быть списком строк.")

    return analysis


# ---------------------------------------------------------------------------


def call_llm(config, articles):
    """
    Отправляет новости в OpenRouter Responses API,
    проверяет статус ответа, JSON и ссылки на источники.
    """
    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        raise RuntimeError(
            "Не найдена переменная окружения OPENROUTER_API_KEY."
        )

    if not articles:
        raise RuntimeError(
            "Нет новостей для анализа. Word-отчёт не создан."
        )

    llm_config = config["llm"]
    base_url = llm_config["base_url"].rstrip("/")
    endpoint = f"{base_url}/responses"
    model = llm_config["model"]

    news_data = []
    for index, article in enumerate(articles, start=1):
        news_data.append({
            "id": index,
            "title": article.get("title", ""),
            "source": article.get("source", ""),
            "author": article.get("author", "Не указан"),
            "published": (
                article["published"].isoformat()
                if hasattr(article.get("published"), "isoformat")
                else str(article.get("published", ""))
            ),
            "url": get_article_source_url(article),
            "text": article.get("text", ""),
        })

    payload = {
        "model": model,
        "instructions": build_instructions(config),
        "input": json.dumps(
            news_data,
            ensure_ascii=False,
            indent=2
        ),
        "max_output_tokens": llm_config.get("max_output_tokens", 5000),
    }

    print("\\nОтправка новостей в OpenRouter...")
    print(f"Модель: {model}")
    print(f"Количество статей: {len(articles)}")
    print(
        "Общий объём текста статей: "
        f"{sum(len(item.get('text', '')) for item in articles)} символов"
    )

    try:
        response = requests.post(
            endpoint,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
                "HTTP-Referer": "http://localhost",
                "X-OpenRouter-Title": "Texas OSINT Study Project",
            },
            json=payload,
            timeout=180,
        )
        response.raise_for_status()

    except requests.HTTPError as error:
        status_code = (
            error.response.status_code
            if error.response is not None
            else "unknown"
        )
        body = (
            error.response.text[:5000]
            if error.response is not None
            else str(error)
        )
        raise RuntimeError(
            f"OpenRouter вернул HTTP {status_code}. Ответ: {body}"
        ) from error

    except requests.RequestException as error:
        raise RuntimeError(
            f"Ошибка соединения с OpenRouter: {error}"
        ) from error

    try:
        response_data = response.json()
    except ValueError as error:
        raise RuntimeError(
            "OpenRouter вернул невалидный JSON на уровне HTTP-ответа. "
            f"Фрагмент: {response.text[:2000]!r}"
        ) from error

    if not isinstance(response_data, dict):
        raise RuntimeError(
            "Неожиданный формат ответа OpenRouter: ожидался JSON-объект."
        )

    response_status = response_data.get("status")
    incomplete_details = response_data.get("incomplete_details")

    print(f"Статус Responses API: {response_status or 'не указан'}")
    if incomplete_details:
        print(
            "Ответ неполный (incomplete_details): "
            f"{json.dumps(incomplete_details, ensure_ascii=False)}"
        )

    if response_status == "incomplete" or incomplete_details:
        raise RuntimeError(
            "Модель вернула неполный ответ. "
            f"status={response_status}, "
            f"incomplete_details={incomplete_details}. "
            "Увеличьте llm.max_output_tokens или уменьшите объём входных данных."
        )

    if response_data.get("error"):
        raise RuntimeError(
            "OpenRouter сообщил об ошибке: "
            + json.dumps(response_data["error"], ensure_ascii=False)
        )

    raw_text = extract_response_text(response_data)

    debug_dir = Path(config.get("output_dir", "reports")) / "debug"
    analysis = parse_json_response(raw_text, debug_dir=debug_dir)

    return validate_analysis(analysis, articles)


# ============================================================
# WORD DOCUMENT
# ============================================================

def set_run_font(run, name="Times New Roman", size=14, bold=None):
    """Принудительно задаёт шрифт для обычного и восточноазиатского текста."""
    run.font.name = name
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold

    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)

    for font_attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        r_fonts.set(qn(f"w:{font_attr}"), name)


def add_hyperlink(paragraph, text, url):
    """Добавляет кликабельную ссылку с оформлением Times New Roman 14."""
    if not url:
        return

    relationship_id = paragraph.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)

    new_run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")

    fonts = OxmlElement("w:rFonts")
    for font_attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        fonts.set(qn(f"w:{font_attr}"), "Times New Roman")
    run_properties.append(fonts)

    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "28")  # 14 pt в half-points
    run_properties.append(size)

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    run_properties.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    run_properties.append(underline)

    new_run.append(run_properties)

    text_element = OxmlElement("w:t")
    text_element.text = str(text)
    new_run.append(text_element)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_source_links(document, urls, articles):
    """
    Добавляет ссылки на источники по URL.
    """

    if not urls:
        document.add_paragraph("Источники не указаны.")
        return

    article_map = {}
    for article in articles:
        for candidate in (
            article.get("url", ""),
            article.get("resolved_url", ""),
            article.get("original_url", ""),
        ):
            normalized = normalize_url(candidate)
            if normalized:
                article_map[normalized] = article

    for url in urls:
        article = article_map.get(normalize_url(url))

        if not article:
            continue

        paragraph = document.add_paragraph(style="List Bullet")

        add_hyperlink(
            paragraph,
            article["title"],
            article["url"]
        )


def add_analysis_section(document, title, items, articles):
    document.add_heading(title, level=2)

    if not items:
        document.add_paragraph("Информация отсутствует.")
        return

    for item in items:
        if isinstance(item, dict):
            statement = item.get("statement", "")
            urls = item.get("source_urls", [])

            if "task" in item:
                statement = item.get("task", "")

                rationale = item.get("rationale", "")

                if rationale:
                    statement += f"\nОбоснование: {rationale}"

        else:
            statement = str(item)
            urls = []

        document.add_paragraph(
            statement,
            style="List Bullet"
        )

        if urls:
            add_source_links(document, urls, articles)


def configure_document_styles(document):
    """Настраивает единое оформление документа в стиле деловых документов РФ."""
    styles = document.styles

    for style in styles:
        if not hasattr(style, "font"):
            continue
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        r_pr = style._element.get_or_add_rPr()
        r_fonts = r_pr.rFonts
        if r_fonts is None:
            r_fonts = OxmlElement("w:rFonts")
            r_pr.insert(0, r_fonts)
        for font_attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            r_fonts.set(qn(f"w:{font_attr}"), "Times New Roman")

    # Основной текст: выравнивание по ширине, 1,5 интервала,
    # абзацный отступ 1,25 см.
    normal = styles["Normal"]
    normal.paragraph_format.alignment = 3  # WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.first_line_indent = Cm(1.25)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    # Заголовки сохраняют размер 14 пт, выделяются полужирным.
    for style_name in ("Title", "Heading 1", "Heading 2", "Heading 3"):
        style = styles[style_name]
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.bold = True
        style.paragraph_format.first_line_indent = Cm(0)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)

    for style_name in ("List Bullet", "List Number"):
        if style_name in styles:
            style = styles[style_name]
            style.font.name = "Times New Roman"
            style.font.size = Pt(14)
            style.paragraph_format.line_spacing = 1.5
            style.paragraph_format.space_after = Pt(0)


def create_word(config, articles, analysis):
    """Создаёт DOCX-отчёт с оформлением Times New Roman 14 и кратким реестром ссылок."""
    output_dir = Path(config.get("output_dir", "reports"))
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_path = output_dir / f"texas_osint_report_{timestamp}.docx"

    document = Document()
    section = document.sections[0]

    # Формат страницы и поля: A4; левое 30 мм, правое 10 мм,
    # верхнее и нижнее по 20 мм.
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(1)

    configure_document_styles(document)

    title = document.add_paragraph(style="Title")
    title.paragraph_format.alignment = 1  # CENTER
    title.paragraph_format.first_line_indent = Cm(0)
    run = title.add_run("OSINT — обзор новостей Техаса")
    set_run_font(run, bold=True)

    metadata = [
        f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        f"Период: последние {config['days_back']} дн.",
        f"Количество собранных новостей: {len(articles)}",
    ]

    if config["topic"] or config["keywords"]:
        metadata.append(f"Тема: {config['topic'] or 'не задана'}")
        if config["keywords"]:
            metadata.append(f"Ключевые слова: {', '.join(config['keywords'])}")
    else:
        metadata.append("Режим: общий мониторинг без тематической фильтрации")

    for line in metadata:
        paragraph = document.add_paragraph(line)
        paragraph.paragraph_format.first_line_indent = Cm(0)
        paragraph.paragraph_format.alignment = 0  # LEFT

    document.add_heading("1. Краткая сводка", level=1)
    document.add_paragraph(
        analysis.get("executive_summary", "Сводка отсутствует.")
    )

    add_analysis_section(
        document, "2. Основные выводы", analysis.get("key_findings", []), articles
    )
    add_analysis_section(
        document, "3. Установленные факты", analysis.get("facts", []), articles
    )
    add_analysis_section(
        document, "4. Мнения и оценки", analysis.get("opinions", []), articles
    )
    add_analysis_section(
        document,
        "5. Возможные задачи и направления проверки",
        analysis.get("potential_tasks", []),
        articles,
    )
    add_analysis_section(
        document, "6. Ограничения анализа", analysis.get("limitations", []), articles
    )

    # Реестр: только нумерованный список заголовков-ссылок.
    # Полный текст и фрагменты статей в конец отчёта не добавляются.
    document.add_heading("7. Реестр собранных новостей", level=1)

    for article in articles:
        paragraph = document.add_paragraph(style="List Number")
        paragraph.paragraph_format.first_line_indent = Cm(0)
        paragraph.paragraph_format.alignment = 0  # LEFT
        source_url = get_article_source_url(article) or article.get("url", "")
        title_text = article.get("title", "Без заголовка")

        if source_url:
            add_hyperlink(paragraph, title_text, source_url)
        else:
            run = paragraph.add_run(title_text)
            set_run_font(run)

    document.save(output_path)
    return output_path


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("TEXAS OSINT NEWS COLLECTOR")
    print("=" * 60)

    try:
        config = load_config()
        articles = collect_articles(config)

        if not articles:
            raise RuntimeError(
                "За указанный период новости не найдены. "
                "Проверь RSS-источники и настройки периода."
            )

        print(f"\\nСобрано новостей: {len(articles)}")

        analyzable_articles = [
            article
            for article in articles
            if len(article.get("text", "")) >= 300
        ]

        skipped_articles = [
            article
            for article in articles
            if len(article.get("text", "")) < 300
        ]

        print(
            f"\\nСтатей с доступным текстом: "
            f"{len(analyzable_articles)}"
        )
        print(
            f"Статей без полноценного текста: "
            f"{len(skipped_articles)}"
        )

        if skipped_articles:
            print("\\nПропущенные статьи:")
            for article in skipped_articles:
                print(
                    f"  - {article.get('title', 'Без заголовка')} "
                    f"[{article.get('extraction_status', 'unknown')}]"
                )

        if not analyzable_articles:
            raise RuntimeError(
                "Не удалось извлечь полноценный текст ни одной статьи. "
                "Анализ отменён."
            )

        # В LLM передаём только статьи с полноценным текстом.
        analysis = call_llm(config, analyzable_articles)

        # В отчёт включаем все собранные новости, в том числе пропущенные
        # для анализа, чтобы сохранить полный реестр сбора.
        output_path = create_word(
            config,
            articles,
            analysis
        )

        print("\\n" + "=" * 60)
        print("[УСПЕХ] Отчёт создан.")
        print(f"Файл: {output_path.resolve()}")
        print("=" * 60)

    except KeyboardInterrupt:
        print("\\n[ОТМЕНА] Выполнение прервано пользователем.")
        sys.exit(1)

    except Exception as error:
        print(f"\\n[ОШИБКА] {error}")
        print("Word-отчёт не создан.")
        sys.exit(1)


if __name__ == "__main__":
    main()
