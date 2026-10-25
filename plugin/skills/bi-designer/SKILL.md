---
name: bi-designer
description: BI-Designer - Boehringer Ingelheim corporate design for any visual output, with or without a PPP project, in VS Code (Copilot) and Claude Code - colors (hex tokens, semantic roles, computed contrast rules), typography (Boehringer Forward, Segoe UI fallback), logo/symbol variants, the 895-icon set, and recipes for Power Apps canvas, Power BI, PowerPoint/Word/Excel (python-pptx, python-docx, openpyxl), charts and diagrams (matplotlib, Plotly, Vega-Lite/Deneb, Mermaid, SVG), web/HTML/CSS, HTML mail, Teams cards and model-driven apps. Load whenever something visual is designed, built or reviewed - a slide deck, report, document, spreadsheet, chart, diagram, app screen, web page, mail, card or icon - also when the user only says "im BI-Design", "Corporate Design" or "Boehringer-Farben".
---

# BI-Designer

Brand profile: **boehringer-ingelheim** – [MARKE.md](marken/boehringer-ingelheim/MARKE.md) (rules),
[tokens.json](marken/boehringer-ingelheim/tokens.json) (values), [icons.txt](marken/boehringer-ingelheim/icons.txt)
(icon names). Further profiles go into `marken/<profil>/` with the same files.

## 1. Workflow (any medium)
1. **Medium and target**: what is produced (slide deck, Word, Excel, chart, app, web, mail, card), where
   it is shown (screen, beamer, print), light or dark surface. Official template available for the
   medium (PowerPoint/Word from the Brand Hub or the company template)? → **use it**; these rules then only
   fill what the template leaves open.
2. **Recipe** for the medium (table §3) – take values from tokens.json, never type colors from memory.
3. **Assets** (§2): logo/symbol variant by background, icons from the brand set.
4. **Build** – generators/scripts where the medium allows (python-pptx, openpyxl, CSS variables, theme
   JSON), so a later colour or size change is one value, not many edits.
5. **Review** with the checklist (§5) before handing over; name every deviation.

## 2. Assets (logos, icons, fonts)
Assets are **not** part of this skill (size, licence). Asset root in this order:
1. the project's setting (PPP: `.ppp/PROJEKT.md` row „Branddesign“),
2. macOS default `~/Library/Mobile Documents/com~apple~CloudDocs/Arbeit/BrandDesign`,
3. otherwise ask once and remember the path in the project.
Relative paths inside the root: `assets` in tokens.json. Copy only used files into the deliverable or
project (`design/assets/`), never whole icon sets; font files never into shared/public places or repos.

## 3. Recipes by medium
| Medium | File |
|---|---|
| Power Apps canvas (App.Formulas theme, control recipes, logo/icons) | [reference/power-apps-canvas.md](reference/power-apps-canvas.md) |
| PowerPoint, Word, Excel (templates first; python-pptx / python-docx / openpyxl) | [reference/office.md](reference/office.md) |
| Charts and diagrams (matplotlib, Plotly, Vega-Lite/Deneb, Mermaid, SVG, draw.io) | [reference/diagramme.md](reference/diagramme.md) |
| Power BI, web/HTML/CSS, HTML mail, Teams cards, model-driven apps, documents | [reference/plattformen.md](reference/plattformen.md) |
| Power BI theme (ready to import) | [marken/boehringer-ingelheim/powerbi-theme.json](marken/boehringer-ingelheim/powerbi-theme.json) |
| Design sheet for a solution | [templates/design.md](templates/design.md) |

## 4. Non-negotiable rules
- Colors only from tokens/roles; never approximate or invent brand colors.
- Contrast: only the allowed pairs in MARKE.md. **Never Accent Green text, icons or thin lines on light
  backgrounds** (1.7:1) – Accent Green only on Dark Green or as a large fill behind Dark Green text.
- Traffic lights (Ampel): Signal Green/Yellow/Red from the master, always with label or legend; signal colors
  and Gray 5/6 never as text.
- Logo/symbol: correct variant for the background, SVG preferred, never recolor, distort, crop or rebuild.
- Icons: brand set only, variant chosen by background (PNG is not tintable), 16–32 px in apps, one icon
  style per view.
- Fonts: Boehringer Forward where the medium supports installed/embedded fonts; otherwise Segoe UI.
- Character: flat color blocks, Dark Green bands, Light Gray sections, generous white space, square
  corners (0–4 px), no gradients, no heavy shadows.
- Not defined in the package (clear space, minimum sizes, imagery, tone of voice) → „(offen, Brand Hub
  prüfen)“ instead of guessing.

## 5. Review checklist (print it with the result)
- [ ] Every color is a token/role value; no free hex values (incl. master grays/signal colors from tokens.json)
- [ ] Traffic lights in signal colors with label/legend
- [ ] Only allowed contrast pairs; no Accent Green on light backgrounds
- [ ] Logo/symbol variant fits its background; not recolored or stretched
- [ ] Icons from the brand set in one variant per background
- [ ] Fonts: Boehringer Forward or Segoe UI, sizes from the scale (28/22/18/14/12 px)
- [ ] Charts use the chart order; line charts on white without Accent Green as thin line
- [ ] Official template used where it exists

## 6. Use in other tools
- **PPP projects:** `ppp-branddesign` adds the project parts (DESIGN.md, Stellschrauben, builder DESIGN).
- **Copilot without PPP:** agent `bi-designer` in the agent picker.
- **Claude Code:** this skill is installed under `~/.claude/skills/bi-designer` by `ppp-update`.
