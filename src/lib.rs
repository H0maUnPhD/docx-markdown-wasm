use quick_xml::events::Event;
use quick_xml::Reader;
use serde::Serialize;
use std::io::{Cursor, Read};
use wasm_bindgen::prelude::*;
use zip::ZipArchive;

#[derive(Serialize)]
struct ConversionResult {
    markdown: String,
    paragraphs: usize,
    headings: usize,
    tables: usize,
    words: usize,
}

fn escape_md(s: &str) -> String {
    s.replace('\\', "\\\\")
        .replace('*', "\\*")
        .replace('_', "\\_")
}

fn parse_document_xml(xml: &str) -> Result<ConversionResult, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut out = String::new();
    let mut paragraph = String::new();
    let mut cell = String::new();
    let mut row: Vec<String> = Vec::new();
    let mut in_text = false;
    let mut in_table = false;
    let mut in_cell = false;
    let mut p_style = String::new();
    let mut paragraphs = 0usize;
    let mut headings = 0usize;
    let mut tables = 0usize;
    let mut table_rows: Vec<Vec<String>> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let n = e.name();
                match n.as_ref() {
                    b"w:t" => in_text = true,
                    b"w:tbl" => {
                        in_table = true;
                        table_rows.clear();
                        tables += 1;
                    }
                    b"w:tc" => {
                        in_cell = true;
                        cell.clear();
                    }
                    b"w:tr" => row.clear(),
                    b"w:pStyle" => {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"w:val" {
                                p_style = String::from_utf8_lossy(&a.value).to_string();
                            }
                        }
                    }
                    b"w:tab" => {
                        if in_cell {
                            cell.push('\t');
                        } else {
                            paragraph.push('\t');
                        }
                    }
                    b"w:br" => {
                        if in_cell {
                            cell.push_str("  \n");
                        } else {
                            paragraph.push_str("  \n");
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let n = e.name();
                match n.as_ref() {
                    b"w:pStyle" => {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"w:val" {
                                p_style = String::from_utf8_lossy(&a.value).to_string();
                            }
                        }
                    }
                    b"w:tab" => {
                        if in_cell {
                            cell.push('\t');
                        } else {
                            paragraph.push('\t');
                        }
                    }
                    b"w:br" => {
                        if in_cell {
                            cell.push_str("  \n");
                        } else {
                            paragraph.push_str("  \n");
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) if in_text => {
                let txt = e.unescape().map_err(|e| e.to_string())?.into_owned();
                if in_cell {
                    cell.push_str(&txt);
                } else {
                    paragraph.push_str(&txt);
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"w:t" => in_text = false,
                b"w:p" => {
                    if !in_table {
                        let p = paragraph.trim();
                        if !p.is_empty() {
                            paragraphs += 1;
                            let clean = escape_md(p);
                            let lower = p_style.to_ascii_lowercase();
                            if lower.contains("heading1") || lower == "title" {
                                out.push_str("# ");
                                out.push_str(&clean);
                                headings += 1;
                            } else if lower.contains("heading2") {
                                out.push_str("## ");
                                out.push_str(&clean);
                                headings += 1;
                            } else if lower.contains("heading3") {
                                out.push_str("### ");
                                out.push_str(&clean);
                                headings += 1;
                            } else {
                                out.push_str(&clean);
                            }
                            out.push_str("\n\n");
                        }
                    }
                    paragraph.clear();
                    p_style.clear();
                }
                b"w:tc" => {
                    in_cell = false;
                    row.push(cell.trim().replace('|', "\\|"));
                    cell.clear();
                }
                b"w:tr" => {
                    if in_table && !row.is_empty() {
                        table_rows.push(row.clone());
                    }
                }
                b"w:tbl" => {
                    in_table = false;
                    if let Some(first) = table_rows.first() {
                        out.push_str("| ");
                        out.push_str(&first.join(" | "));
                        out.push_str(" |\n");
                        out.push_str("| ");
                        out.push_str(&vec!["---"; first.len()].join(" | "));
                        out.push_str(" |\n");
                        for r in table_rows.iter().skip(1) {
                            let mut rr = r.clone();
                            rr.resize(first.len(), String::new());
                            out.push_str("| ");
                            out.push_str(&rr.join(" | "));
                            out.push_str(" |\n");
                        }
                        out.push('\n');
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {e}")),
            _ => {}
        }
    }

    let words = out.split_whitespace().count();
    Ok(ConversionResult {
        markdown: out.trim().to_string(),
        paragraphs,
        headings,
        tables,
        words,
    })
}

#[wasm_bindgen]
pub fn convert_docx(bytes: &[u8]) -> Result<JsValue, JsValue> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor)
        .map_err(|e| JsValue::from_str(&format!("Invalid DOCX/ZIP: {e}")))?;
    let mut file = archive
        .by_name("word/document.xml")
        .map_err(|_| JsValue::from_str("word/document.xml was not found. Is this a valid .docx file?"))?;
    let mut xml = String::new();
    file.read_to_string(&mut xml)
        .map_err(|e| JsValue::from_str(&format!("Could not read document.xml: {e}")))?;
    let result = parse_document_xml(&xml).map_err(|e| JsValue::from_str(&e))?;
    serde_wasm_bindgen::to_value(&result).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn converts_heading_and_paragraph() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Hello</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>World</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
        let result = parse_document_xml(xml).expect("conversion should succeed");
        assert_eq!(result.markdown, "# Hello\n\nWorld");
        assert_eq!(result.headings, 1);
        assert_eq!(result.paragraphs, 2);
    }
}
