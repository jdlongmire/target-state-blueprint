# Target State Blueprint

Build a target-state blueprint "field guide" for an organization function, in a fixed form: the document is a transit line, every content page is a numbered stop, and each stop answers exactly one question its audience keeps asking, named in the stop's header pill. A sustained metaphor carries the argument from the cover to the closing ask.

You write a JSON spec that describes the content; a script turns it into a print-ready PDF (US Letter, portrait) plus a PNG render of the cover. Start with `SKILL.md`; it is the entry point and walks the whole workflow: study the catalog and the design system, intake the function and its recurring questions, map the metaphor, draft the stop plan from the catalog, write the spec, build, run visual QA, deliver.

The package layout:

- `reference/design-system.md` defines the visual system: palette, typography, page furniture, block treatments.
- `reference/stop-catalog.md` is the canonical stop sequence with include/cut guidance.
- `reference/content-spec.md` is the full spec field reference, including scale behavior at 20+ stops.
- `templates/blueprint-spec-template.json` is the fill-in skeleton with the full stop sequence stubbed.
- `scripts/build_blueprint.py` is the builder.
- `examples/` holds one finished worked example on a fictional organization: spec, built PDF, and cover PNG.

Output formats: PDF (the deliverable), plus a 150 dpi PNG of the cover page.

## Author

James D. (JD) Longmire\
Senior Technical Fellow — Aerospace & Defense Sector\
Independent Researcher\
ORCID: 0009-0009-1383-7698\
Correspondence: jdlongmire@outlook.com

## License

MIT License; see `LICENSE`.
