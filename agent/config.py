"""
Varan configuration.
Loads provider/model settings from environment or a .env file.
"""
from __future__ import annotations

import os
from pathlib import Path

import dotenv

ROOT = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ENV_PATH = ROOT / ".env"
CONFIG_PATH = ENV_PATH

DEFAULTS = {
    "provider": "openai",          # openai | anthropic | gemini | openrouter | ollama | custom
    "base_url": "",
    "model": "",
    "api_key": "",
    "max_tokens": 4096,
}

# Provider presets: name -> (default base_url, default model)
PRESETS = {
    "openai": {
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o-mini",
    },
    "openrouter": {
        "base_url": "https://openrouter.ai/api/v1",
        "model": "openai/gpt-4o-mini",
    },
    "groq": {
        "base_url": "https://api.groq.com/openai/v1",
        "model": "llama-3.3-70b-versatile",
    },
    "mistral": {
        "base_url": "https://api.mistral.ai/v1",
        "model": "mistral-small-latest",
    },
    "ollama": {
        "base_url": "http://localhost:11434/v1",
        "model": "llama3.2",
    },
    "lmstudio": {
        "base_url": "http://localhost:1234/v1",
        "model": "local-model",
    },
}


def load_config(path: Path = ENV_PATH) -> dict:
    """Load config: reads .env if present, else environment variables."""
    dotenv.load_dotenv(path)

    cfg = dict(DEFAULTS)
    # Values read from environment (which may have been seeded by .env)
    env_provider = (os.getenv("AI_PROVIDER") or os.getenv("VARAN_PROVIDER") or "").strip().lower()
    env_base = (os.getenv("AI_BASE_URL") or os.getenv("VARAN_BASE_URL") or "").strip()
    env_model = (os.getenv("AI_MODEL") or os.getenv("VARAN_MODEL") or "").strip()
    env_key = os.getenv("AI_API_KEY") or os.getenv("VARAN_API_KEY") or os.getenv("OPENAI_API_KEY") or ""

    provider = env_provider or "openai"
    preset = PRESETS.get(provider, {})

    cfg["provider"] = provider
    cfg["base_url"] = env_base or preset.get("base_url", "")
    cfg["model"] = env_model or preset.get("model", DEFAULTS["model"])
    cfg["api_key"] = env_key
    cfg["max_tokens"] = int(os.getenv("AI_MAX_TOKENS") or DEFAULTS["max_tokens"])

    return cfg


def save_prefs(provider: str, model: str = "", base_url: str = "", api_key: str = "") -> None:
    """Persist the current provider/model/base_url/api_key into .env.

    Empty values are left untouched unless reset=True is used elsewhere; here
    blank strings simply don't overwrite existing keys.
    """
    dotenv.load_dotenv(ENV_PATH)
    dotenv.set_key(ENV_PATH, "AI_PROVIDER", provider)
    if base_url:
        dotenv.set_key(ENV_PATH, "AI_BASE_URL", base_url)
    if model:
        dotenv.set_key(ENV_PATH, "AI_MODEL", model)
    if api_key:
        dotenv.set_key(ENV_PATH, "AI_API_KEY", api_key)
