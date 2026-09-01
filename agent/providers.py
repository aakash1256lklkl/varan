"""
Varan unified provider layer.

Supports ANY AI provider through three strategies:
  1. OpenAI-compatible Chat Completions (OpenAI, OpenRouter, Together, Groq,
     Mistral, DeepSeek, Ollama, LM Studio, vLLM, and most local servers).
  2. Anthropic Messages API (native).
  3. Google Gemini generateContent API (native).

Every provider is wrapped in the same `Provider` interface exposing:
    chat(messages: list[dict]) -> ChatMessage
where ChatMessage has: text (str) and tool_calls (list[ToolCall] | None).
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any, Optional

import httpx

# --------------------------------------------------------------------------
# Data models
# --------------------------------------------------------------------------
@dataclass
class ToolCall:
    id: str
    name: str
    arguments: dict
    raw_arguments: str = ""


@dataclass
class ChatMessage:
    role: str
    text: str = ""
    tool_calls: Optional[list[ToolCall]] = None
    content: Any = None  # raw provider message for round-trips
    raw: Any = None


class ProviderError(Exception):
    pass


# --------------------------------------------------------------------------
# Base provider
# --------------------------------------------------------------------------
class BaseProvider:
    name = "base"

    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.base_url = (cfg.get("base_url") or "").rstrip("/")
        self.model = cfg.get("model") or ""
        self.api_key = cfg.get("api_key") or ""
        try:
            self.max_tokens = int(cfg.get("max_tokens") or 0) or 4096
        except (TypeError, ValueError):
            self.max_tokens = 4096
        # Shared client with generous timeouts for big generations.
        self._client = httpx.Client(timeout=httpx.Timeout(300.0, connect=30.0))

    def chat(self, messages: list[dict], tools: Optional[list[dict]] = None) -> ChatMessage:
        raise NotImplementedError

    def close(self):
        try:
            self._client.close()
        except Exception:  # noqa: BLE001
            pass


# --------------------------------------------------------------------------
# OpenAI-compatible
# --------------------------------------------------------------------------
class OpenAICompatibleProvider(BaseProvider):
    name = "openai-compatible"

    def chat(self, messages: list[dict], tools: Optional[list[dict]] = None) -> ChatMessage:
        url = f"{self.base_url}/chat/completions"
        payload: dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "max_tokens": self.max_tokens,
        }
        if tools:
            payload["tools"] = tools
            payload["tool_choice"] = "auto"

        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        try:
            resp = self._client.post(url, json=payload, headers=headers)
        except httpx.HTTPError as exc:
            raise ProviderError(f"Network error calling {self.name}: {exc}") from exc

        if resp.status_code >= 400:
            raise ProviderError(
                f"{self.name} HTTP {resp.status_code}: {resp.text[:500]}"
            )

        data = resp.json()
        try:
            msg = data["choices"][0]["message"]
        except (KeyError, IndexError) as exc:
            raise ProviderError(f"Unexpected {self.name} response: {data}") from exc

        text = msg.get("content") or ""
        tool_calls = None
        tc = msg.get("tool_calls")
        if tc:
            tool_calls = []
            for call in tc:
                fn = call.get("function", {})
                args_raw = fn.get("arguments", "{}")
                try:
                    arguments = json.loads(args_raw) if args_raw else {}
                except json.JSONDecodeError:
                    arguments = {}
                tool_calls.append(
                    ToolCall(
                        id=call.get("id", ""),
                        name=fn.get("name", ""),
                        arguments=arguments,
                        raw_arguments=args_raw,
                    )
                )
        return ChatMessage(role="assistant", text=text, tool_calls=tool_calls, raw=msg)


# --------------------------------------------------------------------------
# Anthropic
# --------------------------------------------------------------------------
class AnthropicProvider(BaseProvider):
    name = "anthropic"

    def __init__(self, cfg: dict):
        super().__init__(cfg)
        if not cfg.get("base_url"):
            self.base_url = "https://api.anthropic.com"

    def chat(self, messages: list[dict], tools: Optional[list[dict]] = None) -> ChatMessage:
        # Convert OpenAI-style messages to Anthropic format.
        anthropic_messages = []
        system = ""
        # Strip consecutive system role handling: collect system separately.
        for m in messages:
            role = m.get("role")
            content = m.get("content", "")
            if role == "system":
                system += (content if isinstance(content, str) else json.dumps(content)) + "\n"
            elif role == "user":
                anthropic_messages.append({"role": "user", "content": content})
            elif role == "assistant":
                # Reconstruct tool-call history if present
                anthropic_messages.append({"role": "assistant", "content": content})
            elif role == "tool":
                anthropic_messages.append({
                    "role": "user",
                    "content": [{
                        "type": "tool_result",
                        "tool_use_id": m.get("tool_call_id", ""),
                        "content": content if isinstance(content, str) else json.dumps(content),
                    }],
                })

        payload: dict[str, Any] = {
            "model": self.model,
            "max_tokens": 4096,
            "messages": anthropic_messages,
        }
        if system:
            payload["system"] = system

        if tools:
            an_tools = []
            for t in tools:
                fn = t.get("function", {})
                an_tools.append({
                    "name": fn.get("name", ""),
                    "description": fn.get("description", ""),
                    "input_schema": fn.get("parameters", {"type": "object", "properties": {}}),
                })
            payload["tools"] = an_tools

        headers = {
            "Content-Type": "application/json",
            "x-api-key": self.api_key or "",
            "anthropic-version": "2023-06-01",
            "anthropic-dangerous-direct-browser-access": "true",
        }

        url = f"{self.base_url}/v1/messages"
        try:
            resp = self._client.post(url, json=payload, headers=headers)
        except httpx.HTTPError as exc:
            raise ProviderError(f"Network error calling {self.name}: {exc}") from exc

        if resp.status_code >= 400:
            raise ProviderError(f"{self.name} HTTP {resp.status_code}: {resp.text[:500]}")

        data = resp.json()
        text = ""
        tool_calls = None
        for block in data.get("content", []):
            if block.get("type") == "text":
                text += block.get("text", "")
        # Tool use blocks
        tool_uses = [b for b in data.get("content", []) if b.get("type") == "tool_use"]
        if tool_uses:
            tool_calls = []
            for tu in tool_uses:
                tool_calls.append(
                    ToolCall(
                        id=tu.get("id", ""),
                        name=tu.get("name", ""),
                        arguments=tu.get("input", {}) or {},
                        raw_arguments=json.dumps(tu.get("input", {}) or {}),
                    )
                )
        return ChatMessage(role="assistant", text=text, tool_calls=tool_calls, raw=data)


# --------------------------------------------------------------------------
# Google Gemini
# --------------------------------------------------------------------------
class GeminiProvider(BaseProvider):
    name = "gemini"

    def __init__(self, cfg: dict):
        super().__init__(cfg)
        if not cfg.get("base_url"):
            self.base_url = "https://generativelanguage.googleapis.com"

    def chat(self, messages: list[dict], tools: Optional[list[dict]] = None) -> ChatMessage:
        # Convert to Gemini contents: role user/model + function response parts.
        contents = []
        for m in messages:
            role = m.get("role")
            content = m.get("content", "")
            if role == "system":
                # Gemini has no system role; prepend as a user part is not ideal,
                # but most simple cases work with an instruction prefix.
                contents.append({
                    "role": "user",
                    "parts": [{"text": f"[System instructions]\n{content}"}],
                })
            elif role in ("user", "assistant"):
                gem_role = "model" if role == "assistant" else "user"
                contents.append({"role": gem_role, "parts": [{"text": content}]})
            elif role == "tool":
                # function response
                contents.append({
                    "role": "user",
                    "parts": [{
                        "functionResponse": {
                            "name": m.get("name", ""),
                            "response": {"result": content},
                        }
                    }],
                })

        payload: dict[str, Any] = {
            "contents": contents,
        }
        if tools:
            fn_tools = []
            for t in tools:
                fn = t.get("function", {})
                fn_tools.append({
                    "functionDeclarations": [{
                        "name": fn.get("name", ""),
                        "description": fn.get("description", ""),
                        "parameters": fn.get("parameters", {"type": "object", "properties": {}}),
                    }]
                })
            payload["tools"] = fn_tools

        url = f"{self.base_url}/v1beta/models/{self.model}:generateContent"
        if self.api_key:
            url += f"?key={self.api_key}"

        try:
            resp = self._client.post(url, json=payload)
        except httpx.HTTPError as exc:
            raise ProviderError(f"Network error calling {self.name}: {exc}") from exc

        if resp.status_code >= 400:
            raise ProviderError(f"{self.name} HTTP {resp.status_code}: {resp.text[:500]}")

        data = resp.json()
        text = ""
        tool_calls = None
        candidates = data.get("candidates", [])
        if candidates:
            parts = candidates[0].get("content", {}).get("parts", [])
            for part in parts:
                if "text" in part:
                    text += part["text"]
                if "functionCall" in part:
                    fc = part["functionCall"]
                    tc = ToolCall(
                        id="",
                        name=fc.get("name", ""),
                        arguments=fc.get("args", {}) or {},
                        raw_arguments=json.dumps(fc.get("args", {}) or {}),
                    )
                    if tool_calls is None:
                        tool_calls = []
                    tool_calls.append(tc)
        return ChatMessage(role="assistant", text=text, tool_calls=tool_calls, raw=data)


# --------------------------------------------------------------------------
# Factory
# --------------------------------------------------------------------------
def provider_factory(cfg: dict) -> BaseProvider:
    provider = (cfg.get("provider") or "openai").lower()
    if provider == "anthropic" or provider == "claude":
        return AnthropicProvider(cfg)
    if provider == "gemini" or provider == "google" or provider == "googleai":
        return GeminiProvider(cfg)
    # Everything else uses OpenAI-compatible protocol. Ollama, LM Studio,
    # OpenRouter, Groq, Together, Mistral, DeepSeek, vLLM etc.
    return OpenAICompatibleProvider(cfg)
