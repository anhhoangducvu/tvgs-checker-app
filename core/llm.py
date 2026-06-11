# -*- coding: utf-8 -*-
"""
llm.py — Gọi LLM đa nhà cung cấp qua REST (không SDK), cho tvgs-checker-app.
Theo sổ tay 'xay-app-ai-cho-anh-vu': không gõ cứng model, lấy danh sách model thật,
bắt buộc Base URL với loại tương thích OpenAI, vá JSON bị cắt cụt, map lỗi tiếng Việt.
Phụ thuộc: requests.
"""
import json
import re
import requests

DEFAULT_TIMEOUT = 180
MAX_OUTPUT_TOKENS = 32000

PROVIDERS = {
    "anthropic": {"label": "Anthropic (Claude)", "needs_base_url": False,
                  "note": "Key dạng sk-ant-... — lấy tại console.anthropic.com"},
    "gemini": {"label": "Google (Gemini)", "needs_base_url": False,
               "note": "Key lấy tại aistudio.google.com. Bản free hay 429 → chọn model flash/flash-lite"},
    "openai": {"label": "OpenAI (ChatGPT)", "needs_base_url": False,
               "note": "Key dạng sk-... — lấy tại platform.openai.com"},
    "openai_compat": {"label": "Tương thích OpenAI (OpenRouter...)", "needs_base_url": True,
                      "note": "BẮT BUỘC điền Base URL, vd https://openrouter.ai/api/v1"},
}

SYSTEM_PROMPT = ("Bạn là kỹ sư xây dựng cao cấp của Phòng Kỹ thuật — Công ty CP TEXO Tư vấn "
                 "và Đầu tư, chuyên kiểm tra báo cáo Tư vấn Giám sát theo Phụ lục IV NĐ "
                 "06/2021. Trả về DUY NHẤT một JSON hợp lệ theo schema yêu cầu, toàn bộ bằng "
                 "tiếng Việt, không kèm giải thích, không rào ```.")


def _hint(code, provider):
    c = str(code)
    if c in ("401", "403"):
        return ("Key sai hoặc thiếu Base URL. Loại 'tương thích OpenAI' phải điền Base URL, "
                "nếu trống sẽ gọi nhầm OpenAI.")
    if c == "429":
        return ("Hết hạn mức (quota). Đổi model nhẹ hơn, chờ quota reset, hoặc bật billing."
                + (" Gemini bản free rất hay gặp lỗi này." if provider == "gemini" else ""))
    if c == "404":
        return "Sai tên model (model đã bị gỡ) hoặc sai Base URL → bấm 'Lấy danh sách model'."
    return ""


def _call_anthropic(model, key, system, user, base_url=None):
    url = (base_url or "https://api.anthropic.com").rstrip("/") + "/v1/messages"
    r = requests.post(url, timeout=DEFAULT_TIMEOUT,
        headers={"x-api-key": key, "anthropic-version": "2023-06-01",
                 "content-type": "application/json"},
        json={"model": model, "max_tokens": 16000, "system": system,
              "messages": [{"role": "user", "content": user}]})
    r.raise_for_status()
    return "".join(b.get("text", "") for b in r.json().get("content", []))


def _call_gemini(model, key, system, user, base_url=None):
    base = (base_url or "https://generativelanguage.googleapis.com").rstrip("/")
    r = requests.post(f"{base}/v1beta/models/{model}:generateContent?key={key}",
        timeout=DEFAULT_TIMEOUT, headers={"content-type": "application/json"},
        json={"systemInstruction": {"parts": [{"text": system}]},
              "contents": [{"role": "user", "parts": [{"text": user}]}],
              "generationConfig": {"temperature": 0.2, "maxOutputTokens": MAX_OUTPUT_TOKENS}})
    r.raise_for_status()
    data = r.json()
    cands = data.get("candidates", [])
    if not cands:
        fb = (data.get("promptFeedback") or {}).get("blockReason")
        raise RuntimeError("Gemini không trả nội dung" + (f" (bị chặn: {fb})" if fb else ""))
    parts = cands[0].get("content", {}).get("parts", [])
    text = "".join(p.get("text", "") for p in parts)
    if not text:
        raise RuntimeError(f"Gemini trả rỗng (finishReason={cands[0].get('finishReason')}); "
                           "thử model không-thinking (vd gemini-2.0-flash-001).")
    return text


def _call_openai(model, key, system, user, base_url=None):
    base = (base_url or "https://api.openai.com/v1").rstrip("/")
    r = requests.post(base + "/chat/completions", timeout=DEFAULT_TIMEOUT,
        headers={"Authorization": f"Bearer {key}", "content-type": "application/json",
                 "HTTP-Referer": "https://tvgs-checker.local", "X-Title": "TVGS Checker TEXO"},
        json={"model": model, "temperature": 0.2,
              "messages": [{"role": "system", "content": system},
                           {"role": "user", "content": user}]})
    r.raise_for_status()
    data = r.json()
    if "choices" not in data:
        raise RuntimeError(f"Phản hồi không hợp lệ: {str(data)[:200]}")
    return data["choices"][0]["message"]["content"]


def call_llm(provider, model, key, system, user, base_url=None):
    if provider == "anthropic":
        return _call_anthropic(model, key, system, user, base_url)
    if provider == "gemini":
        return _call_gemini(model, key, system, user, base_url)
    if provider == "openai":
        return _call_openai(model, key, system, user, base_url)
    if provider == "openai_compat":
        if not (base_url or "").strip():
            raise ValueError("Loại 'tương thích OpenAI' BẮT BUỘC điền Base URL "
                             "(vd https://openrouter.ai/api/v1). Để trống sẽ gọi nhầm OpenAI.")
        return _call_openai(model, key, system, user, base_url)
    raise ValueError(f"Nhà cung cấp không hỗ trợ: {provider}")


def list_models(provider, key, base_url=None):
    """Trả (ok, [model...] | thông_báo_lỗi)."""
    try:
        if provider == "gemini":
            base = (base_url or "https://generativelanguage.googleapis.com").rstrip("/")
            r = requests.get(f"{base}/v1beta/models?key={key}&pageSize=200", timeout=30)
            r.raise_for_status()
            out = [m.get("name", "").split("/", 1)[-1] for m in r.json().get("models", [])
                   if "generateContent" in (m.get("supportedGenerationMethods") or [])]
            return True, sorted(set(x for x in out if x))
        if provider == "anthropic":
            base = (base_url or "https://api.anthropic.com").rstrip("/")
            r = requests.get(f"{base}/v1/models?limit=100", timeout=30,
                             headers={"x-api-key": key, "anthropic-version": "2023-06-01"})
            r.raise_for_status()
            return True, [m.get("id") for m in r.json().get("data", []) if m.get("id")]
        if provider == "openai_compat" and not (base_url or "").strip():
            return False, "Cần điền Base URL trước khi lấy danh sách model."
        base = (base_url or "https://api.openai.com/v1").rstrip("/")
        r = requests.get(f"{base}/models", timeout=30,
                         headers={"Authorization": f"Bearer {key}"})
        r.raise_for_status()
        return True, sorted(m.get("id") for m in r.json().get("data", []) if m.get("id"))
    except requests.HTTPError as e:
        code = e.response.status_code if e.response is not None else "?"
        return False, f"HTTP {code}: {_hint(code, provider)}".strip()
    except Exception as e:
        return False, str(e)[:200]


def test_connection(provider, model, key, base_url=None):
    try:
        out = call_llm(provider, model, key, "Trả lời đúng một từ.", "Trả lời: OK", base_url)
        return True, (out or "").strip()[:60] or "OK"
    except requests.HTTPError as e:
        code = e.response.status_code if e.response is not None else "?"
        return False, f"HTTP {code}: {_hint(code, provider)}"
    except Exception as e:
        return False, str(e)[:200]


# ---- JSON: vá khi bị cắt cụt ----
def balance_json(s):
    i = s.find("{")
    if i != -1:
        s = s[i:]
    out, stack, in_str, esc = [], [], False, False
    for ch in s:
        if in_str:
            out.append(ch)
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True; out.append(ch)
        elif ch in "{[":
            stack.append(ch); out.append(ch)
        elif ch in "}]":
            if stack:
                stack.pop()
            out.append(ch)
        else:
            out.append(ch)
    res = "".join(out)
    if in_str:
        res += '"'
    res = res.rstrip()
    while res and res[-1] in ",:":
        res = res[:-1].rstrip()
    for ch in reversed(stack):
        res += "}" if ch == "{" else "]"
    return res


def extract_json(text):
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.IGNORECASE).strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    a, b = text.find("{"), text.rfind("}")
    if a != -1 and b > a:
        try:
            return json.loads(text[a:b + 1])
        except Exception:
            pass
    try:
        return json.loads(balance_json(text))
    except Exception:
        return None


def ask_json(provider, model, key, user, base_url=None, system=SYSTEM_PROMPT):
    """Gọi LLM + parse JSON. Trả {ok, data|error, truncated}."""
    try:
        raw = call_llm(provider, model, key, system, user, base_url)
    except requests.HTTPError as e:
        code = e.response.status_code if e.response is not None else "?"
        return {"ok": False, "error": f"HTTP {code}: {_hint(code, provider)}"}
    except Exception as e:
        return {"ok": False, "error": str(e)[:300]}
    data = extract_json(raw)
    if not data:
        return {"ok": False, "error": "AI trả về không phải JSON hợp lệ.", "raw": raw[:4000]}
    return {"ok": True, "data": data,
            "truncated": not raw.strip().rstrip("`").rstrip().endswith("}")}
