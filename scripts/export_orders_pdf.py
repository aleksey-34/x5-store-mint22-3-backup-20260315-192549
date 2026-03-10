from __future__ import annotations

import argparse
import tempfile
from pathlib import Path

import markdown


def build_html(markdown_text: str, title: str) -> str:
    body = markdown.markdown(markdown_text, extensions=["tables", "fenced_code"])
    title_upper = title.upper()
    is_permit = "PERMIT" in title_upper

    page_margin = "8mm 8mm" if is_permit else "10mm 10mm"
    body_font = "10.3pt" if is_permit else "11pt"
    body_line = "1.18" if is_permit else "1.26"
    heading_margin = "0.30em 0 0.15em" if is_permit else "0.42em 0 0.22em"
    h1_size = "15pt" if is_permit else "17pt"
    h2_size = "12.8pt" if is_permit else "14pt"
    h3_size = "11.2pt" if is_permit else "12pt"
    p_margin = "0.12em 0" if is_permit else "0.2em 0"
    table_margin = "0.20em 0 0.35em" if is_permit else "0.35em 0 0.55em"
    cell_padding = "3px" if is_permit else "4px"
    cell_font = "9.3pt" if is_permit else "10pt"
    cell_line = "1.12" if is_permit else "1.2"
    hr_margin = "5px 0" if is_permit else "7px 0"
    list_margin = "0.15em 0 0.22em 1.0em" if is_permit else "0.2em 0 0.35em 1.1em"

    return f"""<!doctype html>
<html lang=\"ru\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>{title}</title>
  <style>
    @page {{
      size: A4;
      margin: {page_margin};
    }}
    body {{
      font-family: "Times New Roman", serif;
      font-size: {body_font};
      line-height: {body_line};
      margin: 0;
      color: #111;
    }}
    h1, h2, h3 {{
      margin: {heading_margin};
      break-after: avoid-page;
    }}
    h1 {{ font-size: {h1_size}; }}
    h2 {{ font-size: {h2_size}; }}
    h3 {{ font-size: {h3_size}; }}
    p {{ margin: {p_margin}; }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin: {table_margin};
      table-layout: fixed;
      word-break: break-word;
      page-break-inside: auto;
    }}
    th, td {{
      border: 1px solid #222;
      padding: {cell_padding};
      vertical-align: top;
      text-align: left;
      font-size: {cell_font};
      line-height: {cell_line};
    }}
    tr {{ page-break-inside: avoid; page-break-after: auto; }}
    hr {{ border: none; border-top: 1px solid #444; margin: {hr_margin}; }}
    ul, ol {{ margin: {list_margin}; }}
    li {{ margin: 0.08em 0; }}
  </style>
</head>
<body>
{body}
</body>
</html>
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export order markdown files to PDF")
    parser.add_argument("--chrome-path", required=True, help="Absolute path to chrome.exe")
    parser.add_argument(
        "--output-dir",
        default="docflow/objects/x5-ufa-e2_logistics_park/01_orders_and_appointments/print_pdf",
        help="Directory for generated PDFs",
    )
    parser.add_argument("files", nargs="+", help="Markdown files to export")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    chrome_path = Path(args.chrome_path)
    if not chrome_path.exists():
        raise FileNotFoundError(f"Chrome not found: {chrome_path}")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    from subprocess import run

    for md_file_raw in args.files:
        md_file = Path(md_file_raw)
        if not md_file.exists():
            raise FileNotFoundError(f"Markdown not found: {md_file}")

        md_text = md_file.read_text(encoding="utf-8")
        html = build_html(md_text, md_file.stem)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as temp:
            temp_path = Path(temp.name)
            temp_path.write_text(html, encoding="utf-8")

        pdf_path = output_dir / f"{md_file.stem}.pdf"
        cmd = [
            str(chrome_path),
            "--headless=new",
            "--disable-gpu",
          "--no-pdf-header-footer",
            f"--print-to-pdf={pdf_path.resolve()}",
            str(temp_path.resolve()),
        ]
        completed = run(cmd, capture_output=True, text=True)
        if completed.returncode != 0:
            raise RuntimeError(
                "Chrome PDF export failed for "
                f"{md_file}: {completed.stdout}\n{completed.stderr}"
            )

        temp_path.unlink(missing_ok=True)
        print(f"PDF exported: {pdf_path}")


if __name__ == "__main__":
    main()
