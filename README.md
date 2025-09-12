# n8n Community Node: Markdown → DOCX

Convert Markdown text into a DOCX file inside your n8n workflows. This node uses `@mohtasham/md-to-docx` to generate a `.docx` and outputs it as binary data.

- **Purpose**: Turn Markdown content from incoming items into a Word document
- **Input sources**: JSON field (string) or Binary text file
- **Output**: Binary `.docx` on a configurable property

Powered by `@mohtasham/md-to-docx` ([repo](https://github.com/MohtashamMurshid/md-to-docx)).

## Installation

Follow n8n’s Community Nodes installation guide:
- In n8n: Settings → Community Nodes → Install
- Search for this package (or install from your local build while developing)
- Docs: https://docs.n8n.io/integrations/community-nodes/installation/

If developing locally, run:
```
npm install
npm run build
```
Then point n8n to your local community node folder per the docs.

## Node: Markdown to DOCX

- **Display name**: Markdown to DOCX
- **Name**: mdToDocx
- **Group**: transform
- **Version**: 1

### Parameters

- **Markdown Source** (options)
  - From Field: read Markdown from a JSON field
  - From Binary (Text File): read Markdown from an incoming binary property
- **Markdown Field** (string, required when source = From Field)
  - Name of the JSON field containing the Markdown string (default: `markdown`)
- **Binary Property** (string, required when source = From Binary)
  - Binary property name that holds a text/markdown file (default: `data`)
- **Filename** (string, required)
  - Output filename for the generated docx (default: `document.docx`)
- **Output Binary Property** (string, required)
  - Binary property to store the generated docx (default: `data`)

### Inputs

- Main (JSON and/or Binary)
  - Provide Markdown via the configured source

### Outputs

- Main (JSON + Binary)
  - Adds a binary property (as configured in "Output Binary Property") containing the generated DOCX
  - MIME type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`

## Usage Examples

### Example 1: Markdown from JSON field

1. Add a Set node that outputs a field `markdown` with your Markdown content, for example:
   - `# Title\n\nSome paragraph with **bold**.`
2. Add the “Markdown to DOCX” node:
   - Markdown Source: From Field
   - Markdown Field: `markdown`
   - Filename: `my-doc.docx`
   - Output Binary Property: `file`
3. Add a node that can handle binary output (e.g. “Write Binary File” or any downstream consumer).

### Example 2: Markdown from Binary

1. Ingest a text/markdown file as binary (e.g. via HTTP Request → Binary or Read Binary File)
2. Add the “Markdown to DOCX” node:
   - Markdown Source: From Binary (Text File)
   - Binary Property: `data` (or your binary key)
   - Filename: `converted.docx`
   - Output Binary Property: `docx`

## Compatibility

- Node.js: >= 20
- n8n Nodes API: 1

## Resources

- n8n Community Nodes docs: https://docs.n8n.io/integrations/#community-nodes
- md-to-docx: https://github.com/MohtashamMurshid/md-to-docx

## Maintainer

GitHub: [@elabbarw](https://github.com/elabbarw)

## Disclaimer

This project is provided as-is. The maintainer does not take responsibility for supporting this node or for any outputs or consequences resulting from its use.

## License

[MIT](https://github.com/n8n-io/n8n-nodes-starter/blob/master/LICENSE.md)
