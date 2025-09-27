#!/usr/bin/env python3
"""
Streamlit App: Postman Collection → CSV/XLSX Doc Exporter

• Загрузи Postman collection JSON (v2.x) + (опц.) env/globals.
• На выходе: CSV или XLSX с таблицей документации.
• Ничего не сохраняется на сервере — только в памяти; даем файл на скачивание.

Запуск локально:
  pip install -r requirements.txt
  streamlit run streamlit_postman_doc_app.py
"""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import streamlit as st

# ----------------- Core helpers -----------------
VAR_RE = re.compile(r"{{\s*([\w\-\.\[\]]+)\s*}}")
ALLURE_RE = re.compile(r"@allure.id=([A-Za-z0-9\-_.]+)")

@dataclass
class Row:
    allure_id: str
    folder: str
    name: str
    method: str
    url_raw: str
    url_resolved: str
    path_params: str
    query_params: str
    headers: str
    body_mode: str
    body_snippet: str
    body_keys: str
    prerequest_script: str
    test_script: str
    examples: str

CSV_FIELDS = [
    "Allure ID",
    "Folder",
    "Request name",
    "Method",
    "URL (raw)",
    "URL (resolved)",
    "Path params",
    "Query params",
    "Headers",
    "Body mode",
    "Body snippet",
    "Body keys",
    "Pre-request script",
    "Test script",
    "Examples",
]

def _json_load_from_upload(upload) -> Optional[dict]:
    if not upload:
        return None
    try:
        return json.loads(upload.getvalue().decode("utf-8"))
    except Exception:
        return None

def _pm_vars_to_dict(blob: Optional[dict]) -> Dict[str, str]:
    res: Dict[str, str] = {}
    if not blob:
        return res
    values = blob.get("values") or []
    for v in values:
        if v.get("enabled", True) and v.get("key") is not None and v.get("value") is not None:
            res[str(v["key"])] = str(v["value"])
    return res

def _collection_vars_to_dict(coll: dict) -> Dict[str, str]:
    res: Dict[str, str] = {}
    for v in coll.get("variable", []) or []:
        key = v.get("key"); val = v.get("value")
        if key is not None and val is not None:
            res[str(key)] = str(val)
    return res

def _resolve_vars(text: Optional[str], sources: List[Dict[str, str]]) -> str:
    if not text:
        return ""
    def repl(m: re.Match[str]) -> str:
        key = m.group(1)
        for src in sources:
            if key in src:
                return str(src[key])
        return m.group(0)
    return VAR_RE.sub(repl, text)

def _json_try_load(s: str) -> Optional[Any]:
    s = (s or "").strip()
    if not s:
        return None
    try:
        return json.loads(s)
    except Exception:
        return None

def _truncate(s: str, n: int) -> str:
    s = s or ""
    if len(s) <= n:
        return s
    return s[: max(0, n - 1)] + "…"

def _join_kv(items: Iterable[Tuple[str, str]]) -> str:
    return "\n".join(f"{k}: {v}" for k, v in items if str(v).strip() != "")

# ---- Robust iterator: yields every request item no matter how nested ----
def iter_items(coll: dict) -> Iterable[Tuple[List[str], dict]]:
    def walk(nodes: List[dict], stack: List[str]):
        for node in (nodes or []):
            if not isinstance(node, dict):
                continue
            # Request node (explicit 'request')
            if "request" in node and node.get("request") is not None:
                yield stack, node
            # Folder
            elif "item" in node:
                name = node.get("name") or ""
                yield from walk(node.get("item") or [], stack + ([name] if name else []))
            # Ad-hoc request shape on item itself
            elif any(k in node for k in ("url", "method", "header", "body")):
                yield stack, node
            # else: ignore unknown nodes
    yield from walk(coll.get("item") or [], [])

def _extract_scripts(item: dict) -> Tuple[str, str]:
    tests: List[str] = []; pres: List[str] = []
    for ev in item.get("event", []) or []:
        listen = (ev.get("listen") or "").lower()
        script = ev.get("script") or {}
        exec_lines = script.get("exec") or []
        code = "\n".join(str(x) for x in exec_lines) if isinstance(exec_lines, list) else str(exec_lines or "")
        if listen == "test": tests.append(code)
        elif listen == "prerequest": pres.append(code)
    return "\n\n".join(pres).strip(), "\n\n".join(tests).strip()

def _extract_url(u: Any) -> Tuple[str, Dict[str, str], Dict[str, str]]:
    raw = ""; path_params: Dict[str, str] = {}; query_params: Dict[str, str] = {}
    if not u: return raw, path_params, query_params
    if isinstance(u, str):
        raw = u
        if "?" in raw:
            _, q = raw.split("?", 1)
            for part in q.split("&"):
                if not part: continue
                if "=" in part:
                    k, v = part.split("=", 1); query_params[k] = v
                else:
                    query_params[part] = ""
        for m in VAR_RE.findall(raw):
            path_params.setdefault(m, f"{{{{{m}}}}}")
        return raw, path_params, query_params
    if isinstance(u, dict):
        raw = u.get("raw") or ""
        for p in u.get("variable", []) or []:
            k = p.get("key"); v = p.get("value")
            if k: path_params[str(k)] = "" if v is None else str(v)
        for q in u.get("query", []) or []:
            k = q.get("key"); v = q.get("value")
            if k: query_params[str(k)] = "" if v is None else str(v)
    return raw, path_params, query_params

def _extract_headers(h: Any) -> Dict[str, str]:
    res: Dict[str, str] = {}
    if isinstance(h, list):
        for x in h:
            k = x.get("key"); v = x.get("value")
            if k: res[str(k)] = "" if v is None else str(v)
    elif isinstance(h, dict):
        for k, v in h.items(): res[str(k)] = "" if v is None else str(v)
    return res

def _extract_body(b: Any) -> Tuple[str, str, str]:
    if not b: return "", "", ""
    mode = (b.get("mode") or "").lower()
    snippet = ""; keys: List[str] = []
    if mode == "raw":
        snippet = str(b.get("raw") or "")
        obj = _json_try_load(snippet)
        if isinstance(obj, dict): keys = list(obj.keys())
    elif mode == "urlencoded":
        parts = []
        for x in b.get("urlencoded") or []:
            k = x.get("key"); v = x.get("value")
            if k: parts.append((str(k), "" if v is None else str(v))); keys.append(str(k))
        snippet = _join_kv(parts)
    elif mode == "formdata":
        parts = []
        for x in b.get("formdata") or []:
            k = x.get("key"); v = x.get("value")
            if k: parts.append((str(k), "<file>" if (x.get("type") == "file") else ("" if v is None else str(v)))); keys.append(str(k))
        snippet = _join_kv(parts)
    elif mode == "file":
        snippet = str(b.get("src") or "<file>")
    else:
        snippet = json.dumps(b, ensure_ascii=False)
    return mode, snippet, ", ".join(keys)

def _extract_examples(item: dict, max_chars: int) -> str:
    chunks: List[str] = []
    for r in item.get("response", []) or []:
        name = r.get("name") or ""
        code = r.get("code")
        body = r.get("body") or ""
        head = f"[{code}] {name}" if code is not None else name
        chunks.append(f"{head}\n{_truncate(str(body), max_chars)}")
    return "\n\n".join(chunks).strip()

def _extract_allure_id(test_script: str) -> str:
    m = ALLURE_RE.search(test_script or "")
    return m.group(1) if m else ""

def collect_rows(coll: dict, var_sources: List[Dict[str, str]], max_body_chars: int, include_examples: bool) -> List[Row]:
    rows: List[Row] = []
    for path, item in iter_items(coll):
        name = item.get("name") or ""
        req = item.get("request") or {}
        # fallback to item-level fields if needed
        method = (req.get("method") or item.get("method") or "").upper()
        raw_url, path_params, query_params = _extract_url(req.get("url") or item.get("url"))
        headers = _extract_headers(req.get("header") or item.get("header"))
        body_mode, body_snippet, body_keys = _extract_body(req.get("body") or item.get("body"))
        pre, tests = _extract_scripts(item)

        url_resolved = _resolve_vars(raw_url, var_sources)
        body_snippet_resolved = _resolve_vars(body_snippet, var_sources)
        examples = _extract_examples(item, max_body_chars) if include_examples else ""
        allure_id = _extract_allure_id(tests)

        rows.append(Row(
            allure_id=allure_id,
            folder=" / ".join(path),
            name=name,
            method=method,
            url_raw=raw_url,
            url_resolved=url_resolved,
            path_params=_join_kv(path_params.items()),
            query_params=_join_kv(query_params.items()),
            headers=_join_kv(headers.items()),
            body_mode=body_mode,
            body_snippet=_truncate(body_snippet_resolved, max_body_chars),
            body_keys=body_keys,
            prerequest_script=_truncate(pre, max_body_chars),
            test_script=_truncate(tests, max_body_chars),
            examples=_truncate(examples, max_body_chars),
        ))
    return rows

def build_dataframe(rows: List[Row]) -> pd.DataFrame:
    data = [{
        "Allure ID": r.allure_id,
        "Folder": r.folder,
        "Request name": r.name,
        "Method": r.method,
        "URL (raw)": r.url_raw,
        "URL (resolved)": r.url_resolved,
        "Path params": r.path_params,
        "Query params": r.query_params,
        "Headers": r.headers,
        "Body mode": r.body_mode,
        "Body snippet": r.body_snippet,
        "Body keys": r.body_keys,
        "Pre-request script": r.prerequest_script,
        "Test script": r.test_script,
        "Examples": r.examples,
    } for r in rows]
    return pd.DataFrame(data, columns=CSV_FIELDS)

# ----------------- Streamlit UI -----------------
st.set_page_config(page_title="Postman → Docs (CSV/XLSX)", page_icon="📦", layout="wide")

st.title("📦 Postman → 📄 Documentation Export")
st.caption("Загрузи Postman коллекцию и забери CSV/XLSX. Ничего не сохраняем — всё в памяти.")

col_left, col_right = st.columns([2, 1])
with col_left:
    collection_file = st.file_uploader("Postman collection (.json)", type=["json"], accept_multiple_files=False)
    #env_file = st.file_uploader("Environment (опционально)", type=["json"], accept_multiple_files=False)
    #globals_file = st.file_uploader("Globals (опционально)", type=["json"], accept_multiple_files=False)

with col_right:
    out_format = st.radio("Формат вывода", options=["CSV", "XLSX"], horizontal=True)
    max_body_chars = st.slider("Макс. длина сниппетов", min_value=100, max_value=4000, value=400, step=100)
    include_examples = st.checkbox("Добавлять Example responses из коллекции", value=False)

st.markdown("---")

go = st.button("Сформировать таблицу")

if go:
    if not collection_file:
        st.error("Загрузи collection.json")
    else:
        coll = _json_load_from_upload(collection_file)
        if not isinstance(coll, dict):
            st.error("Некорректный JSON коллекции")
        else:
            #env = _pm_vars_to_dict(_json_load_from_upload(env_file))
            #glb = _pm_vars_to_dict(_json_load_from_upload(globals_file))
            coll_vars = _collection_vars_to_dict(coll)
            #var_sources = [env, glb, coll_vars]
            var_sources = [coll_vars]

            with st.spinner("Обрабатываю коллекцию…"):
                rows = collect_rows(coll, var_sources, max_body_chars, include_examples)
                df = build_dataframe(rows)

            st.success(f"Готово: найдено запросов — {len(df)}")
            st.dataframe(df, use_container_width=True)

            file_stem = (Path(collection_file.name).stem or "postman").replace(" ", "_")
            if out_format == "CSV":
                buf = io.StringIO()
                df.to_csv(buf, index=False)
                data_bytes = buf.getvalue().encode("utf-8-sig")
                st.download_button(
                    label="⬇️ Скачать CSV",
                    data=data_bytes,
                    file_name=f"{file_stem}.csv",
                    mime="text/csv",
                )
            else:  # XLSX
                bio = io.BytesIO()
                with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Requests")
                    ws = writer.book["Requests"]
                    for col_idx, col_name in enumerate(df.columns, 1):
                        max_len = max([len(str(col_name))] + [len(str(x)) for x in df[col_name].astype(str).tolist()[:200]])
                        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max(12, max_len * 0.9), 80)
                st.download_button(
                    label="⬇️ Скачать XLSX",
                    data=bio.getvalue(),
                    file_name=f"{file_stem}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

with st.expander("ℹ️ Что вытягиваем из коллекции"):
    st.write(
        """
        • Allure ID • Папка/Путь • Имя запроса • Метод • URL (raw/резолв) • Path/Query • Headers • Body mode/snippet/keys • Pre-request/Test scripts • (опц.) Examples.
        Allure ID ищется в Test script строкой `// @allure.id=...`
        Переменные `{{var}}` подставляются из: Environment → Globals → variables коллекции.
        Ничего не пишем на диск: формируем файл в памяти и отдаем на скачивание.
        """
    )
