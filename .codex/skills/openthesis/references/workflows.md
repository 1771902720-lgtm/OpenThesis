# OpenThesis workflows

## Environment

- Node.js 22 or newer
- pnpm 11.x
- An OpenThesis checkout; set `OPENTHESIS_ROOT` only when the checkout is outside the current directory tree

Prepare a fresh checkout:

```bash
git clone https://github.com/1771902720-lgtm/OpenThesis.git
cd OpenThesis
pnpm install --frozen-lockfile
pnpm build
```

Run the skill wrapper from inside the checkout:

```bash
node .codex/skills/openthesis/scripts/openthesis.mjs --help
```

## Markdown to DOCX

Direct conversion:

```bash
node .codex/skills/openthesis/scripts/openthesis.mjs build manuscript.md \
  --type thesis -o manuscript.docx
```

With a parsed template:

```bash
node .codex/skills/openthesis/scripts/openthesis.mjs build manuscript.md \
  -t university.template.json -o manuscript.docx
```

Use `--type` only when front matter does not specify `type`, or when intentionally overriding it.

## Markdown to JSON

Use this when an agent or person needs to inspect or edit the structured content:

```bash
node .codex/skills/openthesis/scripts/openthesis.mjs import manuscript.md \
  --type thesis -o manuscript.json
```

Keep the JSON beside the Markdown when it contains relative image paths. If it must move, update figure paths relative to the JSON file.

## Parse a Word template

```bash
node .codex/skills/openthesis/scripts/openthesis.mjs parse university.docx \
  --type thesis --org "University Name"
```

This writes `university.template.json` beside the source template.

## Build structured JSON

```bash
node .codex/skills/openthesis/scripts/openthesis.mjs build content.json \
  -t university.template.json -o final.docx
```

Without `-t`, OpenThesis uses its built-in USTB template.

## Create starter content

```bash
node .codex/skills/openthesis/scripts/openthesis.mjs init --type thesis
node .codex/skills/openthesis/scripts/openthesis.mjs init --type journal
node .codex/skills/openthesis/scripts/openthesis.mjs init --type official
```

## Validation and recovery

For repository changes:

```bash
pnpm check
pnpm audit --prod --audit-level high
python /root/.codex/skills/.system/skill-creator/scripts/quick_validate.py \
  .codex/skills/openthesis
```

Common failures:

- `OpenThesis CLI is not built`: run `pnpm install --frozen-lockfile && pnpm build` in the repository root.
- Missing figure: make the path relative to the Markdown/JSON source file, or use an absolute path.
- Wrong formatting: verify `-t` points to a template JSON parsed from the intended DOCX file.
- Advanced formula renders literally: native OMML currently covers fractions, roots, scripts, sums, integrals, common symbols, and Greek letters; simplify or pre-convert unsupported macros.
