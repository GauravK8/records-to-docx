#!/usr/bin/env python3
"""
main.py

Usage examples:
    python main.py --kvfile data.txt --template letter_template.docx --outdir output
    python main.py -k data.txt -t template.docx -o output -n "custom_name.docx"

Expected kv file format (UTF-8):
FIRSTNAME=John
LASTNAME=DOE
AGE=30
ADDRESS=USA

Placeholders in the DOCX template should match keys exactly (case-sensitive),
for example: "Dear {{ FIRSTNAME }} {{ LASTNAME }},"
"""

import argparse
from pathlib import Path
import re
import sys
from docxtpl import DocxTemplate

def parse_kv_file(path: Path):
    """
    Parse simple key=value lines into a dict.
    - Ignores blank lines and lines starting with # (comments).
    - Splits on the first '=' only.
    - Strips whitespace and surrounding quotes from values.
    """
    data = {}
    if not path.exists():
        raise FileNotFoundError(f"KV file not found: {path}")
    with path.open(encoding="utf-8") as fh:
        for lineno, raw in enumerate(fh, start=1):
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                # ignore or warn; we'll warn
                print(f"Warning: skipping invalid line {lineno}: {line}", file=sys.stderr)
                continue
            key, val = line.split("=", 1)
            key = key.strip()
            val = val.strip()
            # remove surrounding single/double quotes if present
            if (len(val) >= 2) and ((val[0] == val[-1]) and val[0] in ("'", '"')):
                val = val[1:-1]
            data[key] = val
    return data

def safe_filename(s: str, maxlen=120):
    s = str(s or "")
    s = s.strip()
    s = re.sub(r"[^\w\s-]", "", s, flags=re.UNICODE)   # remove unsafe chars
    s = re.sub(r"\s+", "_", s)
    return s[:maxlen] or "output"

def generate_doc(template_path: Path, context: dict, out_path: Path):
    tpl = DocxTemplate(str(template_path))
    tpl.render(context)
    tpl.save(str(out_path))

def main():
    p = argparse.ArgumentParser(description="Generate DOCX from KV file and DOCX template (docxtpl)")
    p.add_argument("--kvfile", "-k", required=True, help="Path to key=value text file")
    p.add_argument("--template", "-t", required=True, help="Path to DOCX template (use Jinja placeholders matching keys)")
    p.add_argument("--outdir", "-o", default="output_docs", help="Output directory (default: output_docs)")
    p.add_argument("--name", "-n", help="Optional output filename (overrides automatic naming)")
    args = p.parse_args()

    kv_path = Path(args.kvfile)
    tpl_path = Path(args.template)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    if not tpl_path.exists():
        print(f"Error: template not found: {tpl_path}", file=sys.stderr)
        sys.exit(2)

    try:
        ctx = parse_kv_file(kv_path)
    except Exception as e:
        print(f"Error reading kv file: {e}", file=sys.stderr)
        sys.exit(3)

    if not ctx:
        print("Warning: no key/value pairs found in kv file.", file=sys.stderr)

    # Build output filename
    if args.name:
        out_name = args.name
    else:
        # try to use FIRSTNAME or name-like fields
        candidate = ctx.get("FIRSTNAME") or ctx.get("name") or ctx.get("Name") or ctx.get("FULLNAME") or ctx.get("FULL_NAME") or ""
        base = safe_filename(candidate) if candidate else "document"
        out_name = f"{base}.docx"

    out_path = outdir / out_name

    try:
        generate_doc(tpl_path, ctx, out_path)
        print(f"Generated: {out_path}")
    except Exception as e:
        print(f"ERROR generating document: {e}", file=sys.stderr)
        sys.exit(4)

if __name__ == "__main__":
    main()
