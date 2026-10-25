"""Aufbau des Plugins bi-design prüfen (nur Standardbibliothek): python3 -I tests/test_plugin.py

Manifeste, Marktplatz, Version an allen Stellen, Skills und Agenten, Aufräumen alter Kopien.
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
PLUGIN = REPO / "plugin"


def frontmatter(datei):
    text = datei.read_text(encoding="utf-8")
    m = re.match(r"---\n(.*?)\n---\n", text, re.S)
    return dict(z.split(":", 1) for z in m.group(1).splitlines() if ":" in z) if m else {}


class Aufbau(unittest.TestCase):
    def test_version_ueberall_gleich(self):
        vscode = json.loads((PLUGIN / "plugin.json").read_text(encoding="utf-8"))
        claude = json.loads((PLUGIN / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))
        self.assertEqual(vscode["version"], claude["version"])
        self.assertEqual(vscode["name"], claude["name"])
        erste = re.search(r"^## (\S+)", (PLUGIN / "CHANGELOG.md").read_text(encoding="utf-8"), re.M)
        self.assertEqual(erste.group(1), vscode["version"], "CHANGELOG ohne Eintrag für die Version der Manifeste")
        sys.path.insert(0, str(PLUGIN / "tools"))
        import bippt
        self.assertEqual(bippt.__version__, vscode["version"])

    def test_marktplatz(self):
        m = json.loads((REPO / ".claude-plugin" / "marketplace.json").read_text(encoding="utf-8"))
        name = json.loads((PLUGIN / "plugin.json").read_text(encoding="utf-8"))["name"]
        self.assertEqual([p["name"] for p in m["plugins"]], [name])
        self.assertEqual((REPO / m["plugins"][0]["source"]).resolve(), PLUGIN)

    def test_skills_und_agenten(self):
        for skill in sorted((PLUGIN / "skills").iterdir()):
            if skill.is_dir():
                fm = frontmatter(skill / "SKILL.md")
                self.assertEqual(fm.get("name", "").strip(), skill.name)
                self.assertTrue(fm.get("description", "").strip())
        self.assertEqual({p.name for p in (PLUGIN / "skills").iterdir() if p.is_dir()}, {"bi-designer", "bi-ppt"})
        for agent in (PLUGIN / "com.github.copilot" / "agents" / "bi-ppt.agent.md", PLUGIN / "agents" / "bi-ppt.md"):
            self.assertEqual(frontmatter(agent).get("name", "").strip(), "bi-ppt", agent)
        self.assertTrue((PLUGIN / "vorlagen" / "BI_Vorlage_16x9.pptx").is_file())
        self.assertTrue((PLUGIN / "skills" / "bi-designer" / "marken" / "boehringer-ingelheim" / "tokens.json").is_file())

    def test_keine_alten_wege(self):
        erlaubt = {PLUGIN / "tools" / "bippt.py", PLUGIN / "tools" / "bippt" / "konfig.py",
                   PLUGIN / "skills" / "bi-ppt" / "reference" / "installation.md", PLUGIN / "CHANGELOG.md"}
        for datei in PLUGIN.rglob("*"):
            if datei.suffix in (".md", ".py", ".json") and datei not in erlaubt:
                text = datei.read_text(encoding="utf-8")
                for alt in ("~/.bi-ppt", ".bi-ppt/", "bi-ppt-update", "ppp-update", ".ppp/PROJEKT"):
                    self.assertFalse(alt in text, f"{datei.relative_to(REPO)} nennt {alt}")


class Aufraeumen(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.heim = Path(self.tmp.name)
        self.env = {**os.environ, "BIPPT_BENUTZER": str(self.heim), "BIPPT_HOME": str(self.heim / ".bi-design")}

    def tearDown(self):
        self.tmp.cleanup()

    def datei(self, rel, text):
        p = self.heim / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding="utf-8")
        return p

    def bippt(self, *args):
        return subprocess.run([sys.executable, str(PLUGIN / "tools" / "bippt.py"), *args], env=self.env,
                              capture_output=True, text=True, encoding="utf-8")

    def test_nur_eigene_kopien(self):
        agent = self.datei(".copilot/agents/bi-ppt.agent.md", "Du bist **BI-PPT**")
        skill = self.datei(".claude/skills/bi-designer/SKILL.md", "# BI-Designer")
        alt = self.datei(".bi-ppt/tools/bippt.py", "bippt")
        fremd = self.datei(".copilot/agents/bi-designer.agent.md", "ein eigener Agent")
        r = self.bippt("aufraeumen")
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertEqual(r.stdout.count("gefunden"), 3, r.stdout)
        self.assertTrue(agent.exists())
        r = self.bippt("aufraeumen", "--ja")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertFalse(agent.exists())
        self.assertFalse(skill.parent.exists())
        self.assertFalse(alt.parent.parent.exists())
        self.assertTrue(fremd.exists())
        self.assertIn("Keine alten Kopien", self.bippt("aufraeumen").stdout)


if __name__ == "__main__":
    unittest.main(verbosity=1)
