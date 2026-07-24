"""
ai_backend.py — 可切換的 AI 文字生成後端。

以環境變數 AI_BACKEND 選擇後端（預設 ollama）：
  - ollama：本地免費模型（qwen2.5:7b），離線、資料留本機。
  - claude：Anthropic 官方 API（原生 SDK），品質最佳、需 ANTHROPIC_API_KEY。
  - openai：OpenAI Chat Completions，需另裝 openai 套件與 OPENAI_API_KEY。

所有敘事產生都經由 generate() 呼叫；切換後端不需改動 local_ai_analyzer 的邏輯。
可用環境變數：
  AI_BACKEND    = ollama | claude | openai   （預設 ollama）
  OLLAMA_URL    = http://localhost:11434/api/generate
  OLLAMA_MODEL  = qwen2.5:7b
  CLAUDE_MODEL  = claude-opus-4-8
  OPENAI_MODEL  = gpt-4o-mini
  OPENAI_BASE_URL = （選填，相容端點用）
"""
import os

import httpx

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


def backend_name():
    return os.getenv("AI_BACKEND", "ollama").strip().lower()


# ── 簡→繁（台灣用語）保險轉換 ─────────────────────────
# 本地小模型（qwen 等）偶爾漏出簡體字；統一在輸出端過一道 OpenCC s2twp，
# 讓繁體報告不夾簡體。已是繁體者近乎不變；未安裝 opencc 時自動略過。
_cc = None
_cc_tried = False


def _to_traditional(text):
    global _cc, _cc_tried
    if os.getenv("OPENCC", "on").strip().lower() == "off":
        return text
    if not _cc_tried:
        _cc_tried = True
        try:
            from opencc import OpenCC
            _cc = OpenCC("s2twp")
        except Exception:
            _cc = None
    try:
        return _cc.convert(text) if _cc else text
    except Exception:
        return text


# ── Ollama（本地）──────────────────────────────────────
def _generate_ollama(prompt, timeout, max_tokens):
    url = os.getenv("OLLAMA_URL", "http://localhost:11434/api/generate")
    model = os.getenv("OLLAMA_MODEL", "qwen2.5:7b")
    response = httpx.post(
        url,
        json={"model": model, "prompt": prompt, "stream": False},
        timeout=timeout,
    )
    response.raise_for_status()
    return response.json()["response"]


# ── Claude（Anthropic 原生 SDK）───────────────────────
_anthropic_client = None


def _generate_claude(prompt, timeout, max_tokens):
    global _anthropic_client
    from anthropic import Anthropic

    if _anthropic_client is None:
        # 憑證由環境解析（ANTHROPIC_API_KEY 或 ant auth 設定檔）
        _anthropic_client = Anthropic()
    # 預設 claude-opus-4-8；可用 CLAUDE_MODEL 覆寫（例如 claude-sonnet-5）
    model = os.getenv("CLAUDE_MODEL") or "claude-opus-4-8"

    # 注意：目前世代模型不接受 temperature/top_p（送出會 400）；thinking 省略即為關閉。
    response = _anthropic_client.with_options(timeout=timeout).messages.create(
        model=model,
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    if getattr(response, "stop_reason", None) == "refusal":
        raise RuntimeError("Claude 因安全政策拒絕生成此內容")
    return "".join(
        b.text for b in response.content if getattr(b, "type", "") == "text"
    )


# ── OpenAI ─────────────────────────────────────────────
_openai_client = None


def _generate_openai(prompt, timeout, max_tokens):
    global _openai_client
    try:
        from openai import OpenAI
    except ImportError as e:
        raise RuntimeError(
            "使用 OpenAI 後端需先安裝套件：pip install openai"
        ) from e

    if _openai_client is None:
        kwargs = {"api_key": os.getenv("OPENAI_API_KEY")}
        base_url = os.getenv("OPENAI_BASE_URL")
        if base_url:
            kwargs["base_url"] = base_url
        _openai_client = OpenAI(**kwargs)
    model = os.getenv("OPENAI_MODEL") or "gpt-4o-mini"

    response = _openai_client.chat.completions.create(
        model=model,
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
        timeout=timeout,
    )
    return response.choices[0].message.content or ""


_BACKENDS = {
    "ollama": _generate_ollama,
    "claude": _generate_claude,
    "anthropic": _generate_claude,
    "openai": _generate_openai,
    "gpt": _generate_openai,
}


def generate(prompt, timeout=120, max_tokens=1500):
    """依 AI_BACKEND 產生文字。回傳純文字（未清理）。
    max_tokens 僅對雲端後端生效；Ollama 會生成完整回應。"""
    name = backend_name()
    fn = _BACKENDS.get(name)
    if fn is None:
        raise ValueError(
            f"未知的 AI_BACKEND：{name!r}（可用：ollama / claude / openai）"
        )
    return _to_traditional(fn(prompt, timeout, max_tokens))
