"""
VARAN — AI Office Agent
Provider-agnostic AI assistant for Word (.docx), Excel (.xlsx) and
PowerPoint (.pptx). Create, edit, read and summarize Office files.
"""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

# Ensure the base project dir is on the path
os.environ.setdefault("VARAN_ROOT", ROOT)

from agent.compat import enable_utf8_stdio  # noqa: E402

enable_utf8_stdio()

from rich.console import Console  # noqa: E402
from rich.panel import Panel  # noqa: E402
from rich.markdown import Markdown  # noqa: E402

from agent.config import load_config, CONFIG_PATH  # noqa: E402
from agent.providers import provider_factory  # noqa: E402
from agent.loop import Agent  # noqa: E402

console = Console()
BANNER = r"""
██████╗  █████╗ ██████╗  █████╗ ███╗   ██╗
██╔══██╗██╔══██╗██╔══██╗██╔══██╗████╗  ██║
██████╔╝███████║██████╔╝███████║██╔██╗ ██║
██╔══██╗██╔══██╗██╔═══╝ ██╔══██║██║╚██╗██║
██║  ██║██║  ██║██║     ██║  ██║██║ ╚████║
╚═╝  ╚═╝╚═╝  ╚═╝╚═╝     ╚═╝  ╚═╝╚═╝  ╚═══╝
         AI OFFICE AGENT
"""


def print_help() -> None:
    console.print(Panel(
        "[bold]Commands[/bold]\n"
        "  [cyan]/help[/cyan]            show this help\n"
        "  [cyan]/provider <name>[/cyan]   switch provider (e.g. openai, anthropic, ollama)\n"
        "  [cyan]/model <name>[/cyan]      switch model (e.g. gpt-4o, claude-3-5-sonnet, llama3)\n"
        "  [cyan]/config[/cyan]          show current provider/model config\n"
        "  [cyan]/open <file>[/cyan]       set a target file for edits\n"
        "  [cyan]/outputs[/cyan]         list files in the outputs folder\n"
        "  [cyan]/exit[/cyan] or [cyan]quit[/cyan]  leave Varan\n"
        "\n[bold]Talk naturally[/bold]\n"
        "  create a 10-slide pitch deck about a fintech app\n"
        "  make an Excel budget tracker with a monthly bar chart\n"
        "  summarize proposal.docx and turn it into a 5-slide deck\n"
        "\nFiles are written to the [bold]outputs/[/bold] folder so originals are never destroyed. "
        "Editing an existing file writes a copy with [bold]_edited[/bold] appended.",
        title="Varan Help",
        border_style="green",
    ))


def main() -> int:
    console.print(BANNER, style="bold green")
    console.print("[dim]Provider-agnostic AI agent for Word, Excel & PowerPoint[/dim]\n")

    # Load config once so we can report it
    cfg = load_config(CONFIG_PATH)

    try:
        cfg = load_config(CONFIG_PATH)
        provider = provider_factory(cfg)
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]Failed to load config:[/red] {exc}")
        console.print("[yellow]Run: copy .env.example .env  (or set AI_PROVIDER etc.)[/yellow]")
        return 1

    agent = Agent(provider, cfg, console=console)
    console.print(f"[green]Provider:[/green] {cfg['provider']}  "
                  f"[green]Model:[/green] {cfg['model']}")

    while True:
        try:
            prompt = console.input("\n[bold cyan]varan> [/bold cyan]").strip()
        except (KeyboardInterrupt, EOFError):
            console.print("\n[dim]Goodbye.[/dim]")
            break

        if not prompt:
            continue

        low = prompt.lower()

        if low in ("exit", "quit", "/exit"):
            console.print("[dim]Goodbye.[/dim]")
            break
        if low in ("/help", "help", "-h", "--help"):
            print_help()
            continue
        if low.startswith("/provider"):
            parts = prompt.split()
            if len(parts) < 2:
                console.print("[yellow]Usage: /provider <name>[/yellow]")
                continue
            try:
                cfg = agent.change_provider(parts[1])
                console.print(f"[green]Provider →[/green] {cfg['provider']} "
                              f"[green](model:[/green] {cfg['model']}[green])[/green]")
            except Exception as exc:  # noqa: BLE001
                console.print(f"[red]{exc}[/red]")
            continue
        if low.startswith("/model"):
            parts = prompt.split()
            if len(parts) < 2:
                console.print("[yellow]Usage: /model <name>[/yellow]")
                continue
            cfg = agent.change_model(parts[1])
            console.print(f"[green]Model →[/green] {cfg['model']}")
            continue
        if low == "/config":
            c = agent.cfg
            console.print(f"[green]provider:[/green] {c['provider']}  "
                          f"[green]model:[/green] {c['model']}  "
                          f"[green]base_url:[/green] {c.get('base_url') or '(default)'}")
            continue
        if low.startswith("/open"):
            parts = prompt.split(None, 1)
            if len(parts) < 2:
                console.print("[yellow]Usage: /open <file>[/yellow]")
                continue
            agent.set_target_file(parts[1].strip())
            console.print(f"[green]Target file:[/green] {parts[1].strip()}")
            continue
        if low == "/outputs":
            agent.list_outputs()
            continue

        agent.run(prompt)

    return 0


if __name__ == "__main__":
    sys.exit(main())
