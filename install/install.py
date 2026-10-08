#!/usr/bin/env python3
"""BI-PPT installieren/aktualisieren – macOS und Windows, für VS Code (Copilot) und Claude Code.

  python3 install/install.py            installieren (Repo = Ordner über install/)
  python3 install/install.py --update   vorher git pull --ff-only im Repo
  Optionen: --dry (nur anzeigen), --ohne-venv (nur Dateien), --nur-venv (nur Python-Pakete)

Ziele
  ~/.bi-ppt/          Werkzeug (tools/), venv/, vorlagen/, beispiele/, Starter bippt / bippt.cmd, config.json
  ~/.copilot/         agents/bi-ppt.agent.md, skills/bi-ppt, skills/bi-ppt-update   (VS Code + GitHub Copilot)
  ~/.claude/          skills/bi-ppt, skills/bi-ppt-update, agents/bi-ppt.md          (Claude Code, falls vorhanden)
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
HOME = Path.home()
BIPPT = Path(os.environ.get("BIPPT_HOME", HOME / ".bi-ppt"))
COPILOT = Path(os.environ.get("BIPPT_COPILOT_HOME", HOME / ".copilot"))
CLAUDE = Path(os.environ.get("BIPPT_CLAUDE_HOME", HOME / ".claude"))
WIN = os.name == "nt"
SKILLS = ["bi-ppt", "bi-ppt-update"]


def log(msg):
    print(msg, flush=True)


def ordner_kopieren(quelle, ziel, dry):
    log(f"  {'(probe) ' if dry else ''}{ziel}")
    if dry:
        return
    if ziel.exists():
        shutil.rmtree(ziel)
    shutil.copytree(quelle, ziel, ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store", "_render_*"))


def datei_kopieren(quelle, ziel, dry):
    log(f"  {'(probe) ' if dry else ''}{ziel}")
    if not dry:
        ziel.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(quelle, ziel)


def venv_python():
    return BIPPT / "venv" / ("Scripts/python.exe" if WIN else "bin/python")


def venv_einrichten(dry):
    py = venv_python()
    if not py.exists():
        log(f"venv anlegen: {BIPPT / 'venv'} (mit {sys.executable})")
        if not dry:
            subprocess.check_call([sys.executable, "-m", "venv", str(BIPPT / "venv")])
    log("Python-Pakete installieren/aktualisieren …")
    if dry:
        return True
    r = subprocess.run([str(py), "-m", "pip", "install", "-q", "--disable-pip-version-check", "-U", "-r",
                        str(REPO / "tools" / "requirements.txt")], capture_output=True, text=True)
    if r.returncode != 0:
        log("FEHLER pip:\n" + (r.stderr or r.stdout)[-1500:])
        log("Hinweis: hinter einem Proxy `pip config set global.proxy http://…` bzw. Firmen-Index setzen und erneut.")
        return False
    return True


def starter(dry):
    sh = BIPPT / "bippt"
    cmd = BIPPT / "bippt.cmd"
    log(f"  {sh}\n  {cmd}")
    if dry:
        return
    sh.write_text('#!/bin/sh\nexec "$HOME/.bi-ppt/venv/bin/python" "$HOME/.bi-ppt/tools/bippt.py" "$@"\n', encoding="utf-8")
    sh.chmod(0o755)
    cmd.write_text('@echo off\r\nchcp 65001 >nul\r\n"%USERPROFILE%\\.bi-ppt\\venv\\Scripts\\python.exe" '
                   '"%USERPROFILE%\\.bi-ppt\\tools\\bippt.py" %*\r\n', encoding="utf-8")


def konfig(dry):
    datei = BIPPT / "config.json"
    cfg = {}
    if datei.is_file():
        try:
            cfg = json.loads(datei.read_text(encoding="utf-8"))
        except ValueError:
            cfg = {}
    cfg["quelle"] = str(REPO)
    # Vorlage aus dem Repo nutzen, wenn keine eigene gesetzt ist
    if cfg.get("vorlage") and Path(cfg["vorlage"]).resolve().parent == (REPO / "vorlagen").resolve():
        cfg["vorlage"] = ""
    log(f"  {datei} (quelle = {REPO})")
    if not dry:
        datei.write_text(json.dumps(cfg, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--update", action="store_true")
    ap.add_argument("--dry", action="store_true")
    ap.add_argument("--ohne-venv", action="store_true")
    ap.add_argument("--nur-venv", action="store_true")
    a = ap.parse_args()
    if sys.version_info < (3, 9):
        sys.exit(f"Python 3.9+ nötig, gefunden {sys.version.split()[0]}")
    if a.update and (REPO / ".git").exists() and not a.dry:
        r = subprocess.run(["git", "-C", str(REPO), "pull", "--ff-only", "-q"], capture_output=True, text=True)
        log("git pull: " + ((r.stderr or r.stdout).strip() or "ok"))
    version = (REPO / "VERSION").read_text(encoding="utf-8").strip()
    log(f"BI-PPT {version} aus {REPO}")
    ok = True
    if not a.nur_venv:
        log("Werkzeug, Vorlage, Beispiele:")
        ordner_kopieren(REPO / "tools", BIPPT / "tools", a.dry)
        datei_kopieren(REPO / "VERSION", BIPPT / "VERSION", a.dry)
        datei_kopieren(REPO / "vorlagen" / "BI_Vorlage_16x9.pptx", BIPPT / "vorlagen" / "BI_Vorlage_16x9.pptx", a.dry)
        ordner_kopieren(REPO / "skills" / "bi-ppt" / "beispiele", BIPPT / "beispiele", a.dry)
        starter(a.dry)
        konfig(a.dry)
        log("VS Code / GitHub Copilot:")
        datei_kopieren(REPO / "agents" / "bi-ppt.agent.md", COPILOT / "agents" / "bi-ppt.agent.md", a.dry)
        for s in SKILLS:
            ordner_kopieren(REPO / "skills" / s, COPILOT / "skills" / s, a.dry)
        if CLAUDE.exists():
            log("Claude Code:")
            for s in SKILLS:
                ordner_kopieren(REPO / "skills" / s, CLAUDE / "skills" / s, a.dry)
            datei_kopieren(REPO / "claude" / "agents" / "bi-ppt.md", CLAUDE / "agents" / "bi-ppt.md", a.dry)
        else:
            log(f"Claude Code: {CLAUDE} nicht vorhanden – übersprungen")
    if not a.ohne_venv:
        ok = venv_einrichten(a.dry)
    if a.dry:
        log("Probelauf – nichts geschrieben.")
        return 0
    log(f"BI-PPT {version} installiert")
    if ok and venv_python().exists():
        subprocess.run([str(venv_python()), str(BIPPT / "tools" / "bippt.py"), "umgebung"])
    log("Fertig. VS Code neu laden (Developer: Reload Window), Agent „bi-ppt“ wählen; Claude Code: neue Sitzung, /bi-ppt.")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
