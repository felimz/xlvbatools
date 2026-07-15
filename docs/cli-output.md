# Machine-first CLI output

The `xlvba` CLI is machine-first. Every non-interactive command writes exactly
one versioned JSON result envelope to stdout by default. Informational logs go
to stderr, so agents can parse stdout without stripping presentation text.

For command discovery, `xlvba help` returns the same envelope with a versioned
command catalog. `xlvba help COMMAND` narrows it to one command. Conventional
`xlvba --help` and `xlvba COMMAND --help` remain available for terminal use,
with complete option descriptions and examples. JSON discovery includes the
actual parser flags, defaults, choices, and nested subcommands.

```powershell
xlvba lint --source vba_source
xlvba extract
xlvba version
```

The default envelope is the serialized `OperationResult` contract and includes
`schema_version`, `operation`, `success`, `phase`, `data`, structured `error`,
diagnostics, artifacts, metadata, and timing fields. Failed commands still emit
the envelope and return a nonzero process exit code.

## Optional presentation formats

Presentation output must be requested explicitly:

```powershell
xlvba lint --source vba_source --text
xlvba extract --list --table
xlvba snapshot list --output-format table
xlvba version --output-format text
```

The supported values are:

| Value | Shortcut | Purpose |
|:---|:---|:---|
| `json` | none needed | Stable machine-readable default |
| `text` | `--text` | Concise prose or command-native text |
| `table` | `--table` | Aligned rows and columns where applicable |

`--help` remains terminal help. `xlvba debug` is an intentional exception: it
opens interactive Excel/VBE windows, prompts for input, and therefore uses text
throughout.

## Payload format versus presentation format

Some commands create a domain representation in addition to presenting the
result. Keep those choices separate:

```powershell
# JSON envelope containing JSON graph data (both defaults).
xlvba graph

# JSON envelope containing Mermaid content.
xlvba graph --graph-format mermaid

# Raw human-readable Mermaid presentation.
xlvba graph --graph-format mermaid --text
```

Workbook-data files are also explicit side effects:

```powershell
xlvba dump --sheets Input --data --write-json artifacts/input.json
xlvba dump --sheets Input --data --write-markdown artifacts/input.md
```

Changing stdout presentation alone never creates a dump file.
