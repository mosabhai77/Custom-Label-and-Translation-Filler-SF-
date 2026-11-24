#!/usr/bin/env python3
"""
Fill Salesforce Custom Label translations from an Excel sheet.

Excel layout (first row = header):

    | API Name        | fr            | de          | es_MX         |
    |-----------------|---------------|-------------|---------------|
    | Welcome_Message | Bienvenue     | Willkommen  | Bienvenido    |
    | Submit_Button   | Soumettre     | Absenden    | Enviar        |

 - First column  : custom label API name (header text is ignored)
 - Other columns : one per language. Header can be the Salesforce language code
                   (fr, de, es_MX, pt_BR, zh_CN ...) or a common language name
                   (French, German, Spanish ...).

Output: one <lang>.translation-meta.xml per language column, written into the
output folder (default: force-app/main/default/translations). If a file already
exists, its other content is kept and only the given labels are added/updated.

Usage examples:
    python fill_label_translations.py -x translations.xlsx
    python fill_label_translations.py -x translations.xlsx -n Welcome_Message Submit_Button
    python fill_label_translations.py -x translations.xlsx -f label_names.txt -o ./translations
    python fill_label_translations.py -x translations.xlsx -l fr de        # only these languages
"""
import argparse
import sys
from pathlib import Path
import xml.etree.ElementTree as ET

import pandas as pd

NS = "http://soap.sforce.com/2006/04/metadata"
ET.register_namespace("", NS)

# Common language names -> Salesforce language codes (extend as needed)
LANG_NAMES = {
    "french": "fr", "german": "de", "spanish": "es", "italian": "it",
    "portuguese": "pt_PT", "portuguese (brazil)": "pt_BR", "dutch": "nl_NL",
    "japanese": "ja", "korean": "ko", "chinese (simplified)": "zh_CN",
    "chinese (traditional)": "zh_TW", "russian": "ru", "arabic": "ar",
    "hindi": "hi", "bengali": "bn", "swedish": "sv", "danish": "da",
    "finnish": "fi", "norwegian": "no", "polish": "pl", "turkish": "tr",
    "thai": "th", "vietnamese": "vi", "indonesian": "in", "czech": "cs",
    "hungarian": "hu", "romanian": "ro", "greek": "el", "hebrew": "iw",
    "spanish (mexico)": "es_MX", "french (canada)": "fr_CA", "french (belgium)": "fr_BE", "dutch (belgium)": "nl_BE",
}


def q(tag):
    return f"{{{NS}}}{tag}"


def lang_code(header):
    h = str(header).strip()
    return LANG_NAMES.get(h.lower(), h)


def read_label_names(args):
    names = []
    if args.names:
        names += args.names
    if args.names_file:
        names += [l.strip() for l in Path(args.names_file).read_text(encoding="utf-8").splitlines()
                  if l.strip() and not l.startswith("#")]
    # de-duplicate, keep order
    return list(dict.fromkeys(names))


def build_or_load(path):
    if path.exists():
        tree = ET.parse(path)
        return tree, tree.getroot()
    root = ET.Element(q("Translations"))
    return ET.ElementTree(root), root


def upsert_labels(root, labels):
    """labels: dict api_name -> translated text"""
    existing = {}
    for el in root.findall(q("customLabels")):
        n = el.find(q("name"))
        if n is not None:
            existing[n.text] = el

    for name, text in labels.items():
        if name in existing:
            el = existing[name]
            lbl = el.find(q("label"))
            if lbl is None:
                lbl = ET.SubElement(el, q("label"))
            lbl.text = text
            continue
        el = ET.Element(q("customLabels"))
        ET.SubElement(el, q("label")).text = text
        ET.SubElement(el, q("name")).text = name
        # Keep Salesforce's alphabetical element order
        children = list(root)
        idx = len(children)
        for i, c in enumerate(children):
            tag = c.tag.split("}")[1]
            if tag > "customLabels":
                idx = i
                break
        # place after existing customLabels entries
        last_cl = max((i for i, c in enumerate(children) if c.tag == q("customLabels")), default=None)
        if last_cl is not None:
            idx = last_cl + 1
        root.insert(idx, el)
        existing[name] = el

    # sort customLabels entries by API name
    cls = root.findall(q("customLabels"))
    if cls:
        first = list(root).index(cls[0])
        for c in cls:
            root.remove(c)
        cls.sort(key=lambda e: e.find(q("name")).text.lower())
        for offset, c in enumerate(cls):
            root.insert(first + offset, c)


def write_tree(tree, path):
    ET.indent(tree, space="    ")
    path.parent.mkdir(parents=True, exist_ok=True)
    tree.write(path, encoding="UTF-8", xml_declaration=True)
    # Salesforce style: trailing newline + double-quoted declaration
    text = path.read_text(encoding="utf-8")
    text = text.replace("<?xml version='1.0' encoding='UTF-8'?>", '<?xml version="1.0" encoding="UTF-8"?>')
    path.write_text(text.rstrip() + "\n", encoding="utf-8")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("-x", "--excel", required=True, help="Excel file with translations")
    p.add_argument("-s", "--sheet", default=0, help="Sheet name or index (default: first sheet)")
    p.add_argument("-n", "--names", nargs="*", help="Custom label API names to translate")
    p.add_argument("-f", "--names-file", help="Text file with one API name per line")
    p.add_argument("-l", "--languages", nargs="*", help="Only process these language codes/headers")
    p.add_argument("-o", "--out", default="force-app/main/default/translations",
                   help="Output folder for .translation-meta.xml files")
    args = p.parse_args()

    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    df = pd.read_excel(args.excel, sheet_name=sheet, dtype=str).fillna("")
    if df.shape[1] < 2:
        sys.exit("Excel needs an API-name column plus at least one language column.")

    key_col = df.columns[0]
    df[key_col] = df[key_col].str.strip()
    df = df[df[key_col] != ""].drop_duplicates(subset=key_col, keep="last").set_index(key_col)

    wanted = read_label_names(args)
    if wanted:
        missing = [n for n in wanted if n not in df.index]
        if missing:
            print(f"WARNING: not found in Excel: {', '.join(missing)}")
        df = df.loc[[n for n in wanted if n in df.index]]
    if df.empty:
        sys.exit("No matching labels to process.")

    only = {lang_code(l) for l in args.languages} if args.languages else None
    out_dir = Path(args.out)

    for col in df.columns:
        code = lang_code(col)
        if only and code not in only:
            continue
        labels = {name: val for name, val in df[col].items() if str(val).strip()}
        skipped = [n for n in df.index if n not in labels]
        if not labels:
            print(f"[{code}] no translations, skipped")
            continue

        path = out_dir / f"{code}.translation-meta.xml"
        tree, root = build_or_load(path)
        upsert_labels(root, labels)
        write_tree(tree, path)
        msg = f"[{code}] {len(labels)} label(s) written -> {path}"
        if skipped:
            msg += f"  (blank, skipped: {', '.join(skipped)})"
        print(msg)


if __name__ == "__main__":
    main()
