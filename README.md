# CloudPedagogy Word Hierarchy Renderer

Convert a structured table in an editable Microsoft Word document into six switchable interactive hierarchy layouts. Edit the supplied Word example and run one Python command; no D3.js editing is required.

## Files and demonstration

- [Editable Word example](examples/hierarchy_example.docx)
- [Renderer script](render_hierarchy.py)
- [Generated HTML example](output/hierarchy_example/index.html)
- [Normalised example data](output/hierarchy_example/data.json)
- [Example QA report](output/hierarchy_example/qa_report.md)
- [Automated tests](tests/test_render_hierarchy.py)

After enabling GitHub Pages, the live demonstration will be:

https://cloudpedagogy.github.io/cloudpedagogy-word-hierarchy-renderer/output/hierarchy_example/

## Quick start

Python 3.10 or later is recommended.

```bash
git clone https://github.com/cloudpedagogy/cloudpedagogy-word-hierarchy-renderer.git
cd cloudpedagogy-word-hierarchy-renderer

python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt

python3 render_hierarchy.py examples/hierarchy_example.docx \
  --output output/hierarchy_example --overwrite-output
```

Windows PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py render_hierarchy.py examples/hierarchy_example.docx --output output/hierarchy_example --overwrite-output
```

Open `output/hierarchy_example/index.html`.

## Create your own hierarchy

Copy [the Word example](examples/hierarchy_example.docx), replace its sample rows, and retain the `SETTINGS` and `NODES` headings. `NODES` requires `ID` and `Label`; leave `Parent ID` blank for the single root. Optional fields include Type, Value, Status, Description, Colour and Link.

Common alternatives such as `Key`, `Node ID`, `Parent`, `Name`, `Title`, `Category`, `Size`, `Weight`, `Students`, `State`, `Notes` and `URL` are accepted.

Supported layouts are `tree`, `radial-tree`, `sunburst`, `icicle`, `treemap` and `circle-pack`.

## Customisation and limits

The browser output supports layout switching, search, filters, zooming, panning, tooltips and node details. It embeds D3.js and works offline.

The renderer can represent organisations, curricula, websites, taxonomies and other parent–child structures. It requires one valid hierarchy: IDs must be unique, parent references must exist, cycles are invalid, and the data should have one root.

## Output and validation

- `index.html` — interactive offline hierarchy
- `data.json` — parsed and normalised data
- `qa_report.md` — errors and warnings

```bash
python3 render_hierarchy.py --help
python3 -m unittest discover -s tests -v
```

Use `--strict` for automated workflows that should fail on validation findings.

## Licence

MIT. See [LICENSE](LICENSE).
