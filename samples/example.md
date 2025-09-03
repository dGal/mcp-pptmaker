---
marp: true
theme: default
paginate: true
---

# mcp-pptmaker

Convert Marp Markdown to PPTX via an MCP stdio server.

---

## Agenda

- What this tool does
- How to configure MCP
- Generate a sample deck

---

## Slide with bullets

- Built on @marp-team/marp-cli
- Single MCP tool: `generate_pptx()`
- Returns base64 PPTX with:
  - filename: presentation.pptx
  - mimeType: application/vnd.openxmlformats-officedocument.presentationml.presentation

---

## Code sample

```js
console.log("Hello, Marp!");
```

---

## Theming

Use Marp front-matter to control theme, pagination, and other options.

```yaml
---
marp: true
theme: gaia
paginate: true
---
```

---

## Wrap-up

- Add this server to your MCP settings
- Invoke `generate_pptx()` with the markdown content
- Receive a base64-encoded PPTX file in response