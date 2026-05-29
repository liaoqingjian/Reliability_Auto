#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Azure Document Intelligence OCR 后端封装，供 convert_office_to_markdown.py 在
扫描版 PDF 处理时按需调用：

- 无表格页 → prebuilt-read 模型（纯文字 OCR + 坐标）
- 有表格页 → prebuilt-layout 模型（文字 + 表格结构）

对外主入口 `recognize_page_items()` 返回 list[(y, x, text)]，坐标空间与 RapidOCR
分支一致（渲染图的像素坐标），可直接喂给主流程的按坐标排序逻辑。

认证（与 azure_document_intelligence_ocr_test.py 保持一致的解析顺序）:
  - 显式传入 endpoint / api_key
  - 环境变量 AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT / *_KEY 等
  - 模块内 DEFAULT_ENDPOINT / DEFAULT_API_KEY（仅本地调试，勿提交真实密钥）

依赖:
  pip install azure-ai-documentintelligence
"""

from __future__ import annotations

import os
from typing import Any, List, Optional, Tuple

# 自动加载同目录 .env（密钥放 .env，不写入仓库）；未安装 python-dotenv 则忽略
try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

# 与测试脚本相同的中国区多服务终结点 / 临时调试密钥（生产请改用环境变量）
DEFAULT_ENDPOINT = "https://chinaeast2.api.cognitive.azure.cn/"
DEFAULT_API_KEY = "0093731c776f4aa0a0e69a542db51d0b"

# 模型 ID
MODEL_READ = "prebuilt-read"
MODEL_LAYOUT = "prebuilt-layout"

# (y, x, text)
PageItem = Tuple[float, float, str]


def _normalize_endpoint(url: str) -> str:
    return url.strip().rstrip("/") + "/"


def _resolve_endpoint(explicit: Optional[str]) -> str:
    if explicit and explicit.strip():
        return _normalize_endpoint(explicit)
    for name in (
        "AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT",
        "DOCUMENTINTELLIGENCE_ENDPOINT",
        "AZURE_COMPUTER_VISION_ENDPOINT",
    ):
        v = os.environ.get(name, "").strip()
        if v:
            return _normalize_endpoint(v)
    return _normalize_endpoint(DEFAULT_ENDPOINT)


def _resolve_api_key(explicit: Optional[str]) -> str:
    if explicit and explicit.strip():
        return explicit.strip()
    for name in (
        "AZURE_DOCUMENT_INTELLIGENCE_KEY",
        "DOCUMENTINTELLIGENCE_API_KEY",
        "AZURE_COMPUTER_VISION_KEY",
        "COMPUTER_VISION_SUBSCRIPTION_KEY",
    ):
        v = os.environ.get(name, "").strip()
        if v:
            return v
    return (DEFAULT_API_KEY or "").strip()


_CLIENT_CACHE: dict[tuple[str, str], Any] = {}


def _get_client(endpoint: str, api_key: str) -> Any:
    """缓存 DocumentIntelligenceClient（同一份 PDF 转换内复用连接）。"""
    cache_key = (endpoint, api_key)
    client = _CLIENT_CACHE.get(cache_key)
    if client is not None:
        return client
    from azure.ai.documentintelligence import DocumentIntelligenceClient
    from azure.core.credentials import AzureKeyCredential

    client = DocumentIntelligenceClient(
        endpoint=endpoint, credential=AzureKeyCredential(api_key)
    )
    _CLIENT_CACHE[cache_key] = client
    return client


def _analyze(client: Any, model_id: str, image_bytes: bytes) -> Any:
    """对单页渲染图（PNG 字节）调用指定模型，返回 AnalyzeResult。"""
    poller = client.begin_analyze_document(
        model_id,
        body=image_bytes,
        content_type="application/octet-stream",
    )
    return poller.result()


def _poly_min_yx(polygon: Any) -> tuple[float, float]:
    """polygon 为扁平浮点列表 [x1,y1,x2,y2,...]，返回 (min_y, min_x)。"""
    if not polygon:
        return 0.0, 0.0
    try:
        xs = [float(v) for v in polygon[0::2]]
        ys = [float(v) for v in polygon[1::2]]
    except (TypeError, ValueError):
        return 0.0, 0.0
    if not xs or not ys:
        return 0.0, 0.0
    return min(ys), min(xs)


def _poly_aabb(polygon: Any) -> tuple[float, float, float, float]:
    if not polygon:
        return 0.0, 0.0, 0.0, 0.0
    try:
        xs = [float(v) for v in polygon[0::2]]
        ys = [float(v) for v in polygon[1::2]]
    except (TypeError, ValueError):
        return 0.0, 0.0, 0.0, 0.0
    if not xs or not ys:
        return 0.0, 0.0, 0.0, 0.0
    return min(xs), min(ys), max(xs), max(ys)


def _table_first_region_polygon(table: Any) -> Any:
    regions = getattr(table, "bounding_regions", None) or []
    for r in regions:
        poly = getattr(r, "polygon", None)
        if poly:
            return poly
    return None


def _layout_table_to_html(table: Any) -> str:
    """layout 表格 cells（含 row/column span）→ 单行 HTML <table>。"""
    import html as _html

    rc = int(getattr(table, "row_count", 0) or 0)
    cc = int(getattr(table, "column_count", 0) or 0)
    cells = list(getattr(table, "cells", None) or [])
    if rc <= 0 or cc <= 0 or not cells:
        return ""

    anchor: dict[tuple[int, int], Any] = {}
    for cell in cells:
        r = int(getattr(cell, "row_index", 0) or 0)
        c = int(getattr(cell, "column_index", 0) or 0)
        anchor[(r, c)] = cell

    covered = [[False] * cc for _ in range(rc)]
    parts: list[str] = ['<table border="1">']
    for r in range(rc):
        parts.append("<tr>")
        c = 0
        while c < cc:
            if covered[r][c]:
                c += 1
                continue
            cell = anchor.get((r, c))
            if cell is None:
                parts.append("<td></td>")
                c += 1
                continue
            rs = max(1, int(getattr(cell, "row_span", 1) or 1))
            cs = max(1, int(getattr(cell, "column_span", 1) or 1))
            for rr in range(r, min(rc, r + rs)):
                for ccx in range(c, min(cc, c + cs)):
                    covered[rr][ccx] = True
            attrs = ""
            if rs > 1:
                attrs += f' rowspan="{rs}"'
            if cs > 1:
                attrs += f' colspan="{cs}"'
            content = (getattr(cell, "content", "") or "").replace("\n", " ").strip()
            parts.append(f"<td{attrs}>{_html.escape(content)}</td>")
            c += cs
        parts.append("</tr>")
    parts.append("</table>")
    return "".join(parts)


def _read_items_from_result(result: Any) -> List[PageItem]:
    items: List[PageItem] = []
    for page in getattr(result, "pages", None) or []:
        for line in getattr(page, "lines", None) or []:
            text = (getattr(line, "content", "") or "").strip()
            if not text:
                continue
            y, x = _poly_min_yx(getattr(line, "polygon", None))
            items.append((y, x, text))
    return items


def _layout_items_from_result(result: Any) -> List[PageItem]:
    """layout 结果 → 表格块 + 表外文字行，统一为 (y, x, text)。"""
    # 延迟导入主模块的表格转换工具，避免模块级循环依赖
    from convert_office_to_markdown import _html_table_to_md

    items: List[PageItem] = []
    table_bboxes: List[tuple[float, float, float, float]] = []

    for table in getattr(result, "tables", None) or []:
        poly = _table_first_region_polygon(table)
        if poly:
            x0, y0, x1, y1 = _poly_aabb(poly)
            table_bboxes.append((x0, y0, x1, y1))
            ty, tx = y0, x0
        else:
            ty, tx = 0.0, 0.0
        html = _layout_table_to_html(table)
        table_md = _html_table_to_md(html) if html else ""
        if table_md:
            items.append((ty, tx, table_md))

    def _in_any_table(cx: float, cy: float) -> bool:
        for x0, y0, x1, y1 in table_bboxes:
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                return True
        return False

    for page in getattr(result, "pages", None) or []:
        for line in getattr(page, "lines", None) or []:
            text = (getattr(line, "content", "") or "").strip()
            if not text:
                continue
            x0, y0, x1, y1 = _poly_aabb(getattr(line, "polygon", None))
            cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
            if _in_any_table(cx, cy):
                continue
            items.append((y0, x0, text))

    return items


def recognize_page_items(
    image_bytes: bytes,
    *,
    has_table: bool,
    endpoint: Optional[str] = None,
    api_key: Optional[str] = None,
) -> List[PageItem]:
    """
    单页扫描图 OCR：has_table=False → prebuilt-read；has_table=True → prebuilt-layout。
    返回 list[(y, x, text)]（坐标为渲染图像素），失败时抛出异常由调用方处理。
    """
    ep = _resolve_endpoint(endpoint)
    key = _resolve_api_key(api_key)
    if not key:
        raise ValueError(
            "缺少 Azure 密钥：请设置 AZURE_OCR_KEY、环境变量 "
            "AZURE_DOCUMENT_INTELLIGENCE_KEY，或填写模块内 DEFAULT_API_KEY。"
        )
    client = _get_client(ep, key)

    if has_table:
        result = _analyze(client, MODEL_LAYOUT, image_bytes)
        return _layout_items_from_result(result)
    result = _analyze(client, MODEL_READ, image_bytes)
    return _read_items_from_result(result)
