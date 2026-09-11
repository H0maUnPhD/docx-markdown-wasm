# DOCX → Markdown with Rust + WebAssembly

A small client-side web app that converts `.docx` files to Markdown directly in the browser.

- Rust conversion core
- Compiled to WebAssembly with `wasm-pack`
- No backend
- DOCX stays on the user's device
- Handles paragraphs, Heading 1–3, basic tables, tabs and line breaks
- Includes a native Rust unit test

## GitHub Actions

Every push to `main` runs tests and builds the WebAssembly package. The compiled browser bundle is uploaded as the `docx-markdown-wasm-build` artifact.

## Local build

```bash
cargo test
cargo install wasm-pack --locked
wasm-pack build --target web --out-dir pkg --release
```

Then serve the repository directory with any static HTTP server and open `web/index.html`.
