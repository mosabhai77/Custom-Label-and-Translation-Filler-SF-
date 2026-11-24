"""
Streamlit app: edit Salesforce CustomLabels + translation XML files IN PLACE from an Excel sheet.

Run:  python3 -m streamlit run app.py

Excel format (first row = headers):
    API Name | Master Label | fr | de | es ...
 - Label already in CustomLabels.labels-meta.xml -> only its translations are added/updated
   (Master Label may be left blank).
 - Label NOT in the file yet -> it is added to CustomLabels.labels-meta.xml (Master Label required)
   and its translations are added.
"""
import io
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from fill_label_translations import lang_code

NONE = "(none)"
NAME_OK = re.compile(r"^[A-Za-z](?:[A-Za-z0-9]|_(?=[A-Za-z0-9]))*$")
MASTER_RE = re.compile(r"master|default|base|english|^en([_-]us)?$", re.I)
SKIP_DIRS = {"node_modules", ".git", ".sfdx", ".sf", ".vscode", ".venv", "__pycache__"}
NEW_TR = ('<?xml version="1.0" encoding="UTF-8"?>\n'
          '<Translations xmlns="http://soap.sforce.com/2006/04/metadata">\n</Translations>\n')

SAMPLE = [
    ('Welcome_Message', 'Welcome to our portal', 'Bienvenue sur notre portail', 'Willkommen auf unserem Portal', 'Bienvenido a nuestro portal', 'Tervetuloa portaaliimme', 'Bienvenue sur notre portail', 'Welkom op ons portaal'),
    ('Submit_Button', 'Submit', 'Soumettre', 'Absenden', 'Enviar', 'Lähetä', 'Envoyer', 'Verzenden'),
    ('Cancel_Button', 'Cancel', 'Annuler', 'Abbrechen', 'Cancelar', 'Peruuta', 'Annuler', 'Annuleren'),
    ('Save_Changes', 'Save changes', 'Enregistrer les modifications', 'Änderungen speichern', 'Guardar cambios', 'Tallenna muutokset', 'Enregistrer les modifications', 'Wijzigingen opslaan'),
    ('Required_Field_Error', 'Required field missing', 'Champ obligatoire manquant', 'Pflichtfeld fehlt', 'Falta un campo obligatorio', 'Pakollinen kenttä puuttuu', 'Champ obligatoire manquant', 'Verplicht veld ontbreekt'),
    ('Thank_You_Message', 'Thank you for contacting us', 'Merci de nous avoir contactés', 'Vielen Dank für Ihre Kontaktaufnahme', 'Gracias por contactarnos', 'Kiitos yhteydenotostasi', 'Merci de nous avoir contactés', 'Bedankt dat u contact met ons hebt opgenomen'),
    ('Session_Expired', 'Your session has expired', 'Votre session a expiré', 'Ihre Sitzung ist abgelaufen', 'Su sesión ha caducado', 'Istuntosi on vanhentunut', 'Votre session a expiré', 'Uw sessie is verlopen'),
    ('Search_Records', 'Search records', 'Rechercher des enregistrements', 'Datensätze suchen', 'Buscar registros', 'Hae tietueita', 'Rechercher des enregistrements', 'Records zoeken'),
    ('Contact_Us_Title', 'Contact us', 'Contactez-nous', 'Kontakt', 'Contáctenos', 'Ota yhteyttä', 'Contactez-nous', 'Neem contact met ons op'),
    ('Case_Created_Success', 'Case created successfully', 'Requête créée avec succès', 'Fall erfolgreich erstellt', 'Caso creado correctamente', 'Tapaus luotu onnistuneesti', 'Requête créée avec succès', 'Case succesvol aangemaakt'),
]


# ---------- XML text editing (touches only the block concerned, keeps formatting) ----------
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;").replace("'", "&apos;"))


def unesc(s):
    return (str(s).replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
            .replace("&apos;", "'").replace("&amp;", "&"))


def upsert(text, tag, key_tag, key, fields, upd_tag, val, overwrite, root_tag):
    """Add the <tag> block for `key` (alphabetical position) or update its <upd_tag>. Returns (text, status)."""
    nl = "\r\n" if "\r\n" in text else "\n"
    blocks = list(re.finditer(rf"<{tag}>[\s\S]*?</{tag}>", text))
    key_re = re.compile(rf"<{key_tag}>([^<]*)</{key_tag}>")

    def k_of(m):
        k = key_re.search(m.group(0))
        return unesc(k.group(1).strip()) if k else ""

    ex = next((m for m in blocks if k_of(m) == key), None)
    if ex:
        v_re = re.compile(rf"<{upd_tag}>([\s\S]*?)</{upd_tag}>")
        cur = v_re.search(ex.group(0))
        if not cur or not overwrite or unesc(cur.group(1)) == val:
            return text, "unchanged"
        nb = v_re.sub(lambda _: f"<{upd_tag}>{esc(val)}</{upd_tag}>", ex.group(0), count=1)
        return text[:ex.start()] + nb + text[ex.end():], "updated"

    if blocks:
        ind1 = re.search(r"[ \t]*$", text[:blocks[0].start()]).group(0)
        m = re.search(r"\n([ \t]+)<", blocks[0].group(0))
        inn = m.group(1) if m else ind1 + "    "
    else:
        m = re.search(r"\n([ \t]+)<[A-Za-z]", text)
        ind1 = m.group(1) if m else "    "
        inn = ind1 + "    "
    nb = (f"<{tag}>" + "".join(f"{nl}{inn}<{t}>{esc(v)}</{t}>" for t, v in fields) + f"{nl}{ind1}</{tag}>")

    nxt = next((m for m in blocks if k_of(m).lower() > key.lower()), None)
    if nxt:
        out = text[:nxt.start()] + nb + nl + ind1 + text[nxt.start():]
    elif blocks:
        e = blocks[-1].end()
        out = text[:e] + nl + ind1 + nb + text[e:]
    else:
        pos, top = -1, None
        for m in re.finditer(r"\n([ \t]+)<([A-Za-z]+)[ >/]", text):
            top = m.group(1) if top is None else top
            if m.group(1) == top and m.group(2) > tag:
                pos = m.start() + 1 + len(m.group(1))
                break
        if pos >= 0:
            out = text[:pos] + nb + nl + ind1 + text[pos:]
        else:
            c = text.rfind(f"</{root_tag}>")
            out = text[:c] + ind1 + nb + nl + text[c:]
    return out, "added"


def read(p):
    return p.read_bytes().decode("utf-8")


def save(p, t):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(t.encode("utf-8"))


def scan(root):
    found = []
    for dp, dns, fns in os.walk(root):
        dns[:] = [d for d in dns if d not in SKIP_DIRS]
        if "CustomLabels.labels-meta.xml" in fns:
            found.append(Path(dp) / "CustomLabels.labels-meta.xml")
    return sorted(found)


def apply_changes(rows, langs, labels_path, trans_dir, master_lang, protected, ow_master, ow_trans, base):
    """rows: [(api_name, master, {lang_col: text})].
    Nothing is written here. Returns (per_label_detail, per_file_report, changes) where
    changes = {Path: (original_text_or_None_if_new_file, new_text)}."""
    rel = lambda p: str(p.relative_to(base)) if base in p.parents else str(p)
    detail = {r[0]: {"API Name": r[0]} for r in rows}
    files, skipped, changes = [], set(), {}

    # 1) master labels
    text = read(labels_path)
    orig_labels = text
    existing = {unesc(m.group(1).strip()) for m in re.finditer(r"<fullName>([^<]*)</fullName>", text)}
    c = dict(added=0, updated=0, unchanged=0)
    for n, master, _ in rows:
        exists = n in existing
        if not exists and not master:
            detail[n]["CustomLabels"] = "SKIPPED: new label needs a Master Label"
            skipped.add(n)
            continue
        if exists and not (ow_master and master):
            detail[n]["CustomLabels"] = "already present"
            c["unchanged"] += 1
            continue
        text, s = upsert(text, "labels", "fullName", n,
                         [("fullName", n), ("language", master_lang), ("protected", str(protected).lower()),
                          ("shortDescription", n.replace("_", " ")), ("value", master)],
                         "value", master, True, "CustomLabels")
        c[s] += 1
        detail[n]["CustomLabels"] = {"added": "added (new label)", "updated": "master value updated",
                                     "unchanged": "already present"}[s]
    if c["added"] or c["updated"]:
        changes[labels_path] = (orig_labels, text)
    files.append({"File": rel(labels_path), "Added": c["added"], "Updated": c["updated"],
                  "Unchanged": c["unchanged"], "New file": False})

    # 2) translations
    for col in langs:
        code = lang_code(col)
        path = trans_dir / f"{code}.translation-meta.xml"
        is_new = not path.exists()
        t = NEW_TR if is_new else read(path)
        old = None if is_new else t
        c = dict(added=0, updated=0, unchanged=0)
        for n, _, tr in rows:
            v = tr[col].strip()
            if n in skipped or not v:
                detail[n][code] = "-"
                continue
            t, s = upsert(t, "customLabels", "name", n, [("label", v), ("name", n)], "label", v, ow_trans, "Translations")
            c[s] += 1
            detail[n][code] = s
        if c["added"] or c["updated"]:
            changes[path] = (old, t)
        if c["added"] or c["updated"] or not is_new:
            files.append({"File": rel(path), "Added": c["added"], "Updated": c["updated"],
                          "Unchanged": c["unchanged"], "New file": is_new})
    return list(detail.values()), files, changes


def make_backup(changes, root):
    """Zip the ORIGINAL (unchanged) contents of the files about to be edited, save it in ./backups
    next to this app, and return info for the UI. Returns None when nothing will change."""
    olds = {p: old for p, (old, _) in changes.items() if old is not None}
    created = [p for p, (old, _) in changes.items() if old is None]
    if not olds and not created:
        return None

    def entry(p):
        try:
            return f"{root.name}/{p.relative_to(root).as_posix()}"
        except ValueError:
            return p.as_posix().lstrip("/")

    now = datetime.now()
    info = ([f"Backup created {now:%Y-%m-%d %H:%M:%S} by Custom Label XML Updater", "",
             "These are the ORIGINAL, unchanged files, taken before the app edited them.", "",
             "To revert: copy each file below back to its original location (overwrite the edited file).", "",
             "Original files (zip entry -> location):"]
            + [f"  {entry(p)} -> {p}" for p in olds]
            + ["", "Files newly created by the app (delete them to revert):"]
            + ([f"  {p}" for p in created] or ["  (none)"]))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p, old in olds.items():
            z.writestr(entry(p), old.encode("utf-8"))
        z.writestr("_BACKUP_INFO.txt", "\n".join(info) + "\n")
    bdir = Path(__file__).resolve().parent / "backups"
    bdir.mkdir(exist_ok=True)
    path = bdir / f"backup_{now:%Y%m%d_%H%M%S}.zip"
    path.write_bytes(buf.getvalue())
    return dict(path=path, data=buf.getvalue(), originals={entry(p): old for p, old in olds.items()},
                created=[str(p) for p in created])


# ---------- sample / template Excel ----------
def make_xlsx(with_samples):
    wb = Workbook()
    ws = wb.active
    ws.title = "Translations"
    ws.append(["API Name", "Master Label", "fr", "de", "es", "fi", "fr_BE", "nl_NL"])
    for r in (SAMPLE if with_samples else []):
        ws.append(list(r))
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
    for col, w in zip("ABCDEFGH", (28, 32, 36, 40, 34, 30, 36, 38)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ins = wb.create_sheet("Instructions")
    for i, t in enumerate([
        "How to fill this sheet (the app reads the first sheet, first row = headers)",
        "API Name: custom label API name (case-sensitive). Required.",
        "Master Label: the default (English) text. Needed only for NEW labels that are not yet in CustomLabels.labels-meta.xml.",
        "   For labels that already exist, leave it blank (or fill it and tick 'Overwrite master value' in the app).",
        "Language columns: header = Salesforce language code (fr, de, es, pt_BR, zh_CN ...) or a name like French.",
        "   Add or delete language columns as needed. Blank cells are skipped.",
    ], 1):
        ins.cell(row=i, column=1, value=t).font = Font(name="Arial", bold=(i == 1), size=10)
    ins.column_dimensions["A"].width = 120
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------- UI ----------
def main():
    st.set_page_config(page_title="Custom Label XML Updater", layout="wide")
    st.title("Custom Label XML Updater")
    st.caption("Reads your Excel sheet and updates CustomLabels.labels-meta.xml and the translation files directly "
               "in your Salesforce project. Review the result with VS Code's diff before deploying.")

    with st.sidebar:
        st.header("Settings")
        master_lang = st.text_input("Master label language (new labels)", "en_US")
        protected = st.checkbox("Mark new labels as protected", False)
        ow_master = st.checkbox("Overwrite master value of labels that already exist", False,
                                help="Off: existing labels only get translations.")
        ow_trans = st.checkbox("Overwrite existing translations", True)

    st.subheader("1. Excel format")
    st.markdown("Columns: **API Name** | **Master Label** | one column per language (`fr`, `de`, `es` ...). "
                "For labels that already exist, Master Label can be blank and only translations are added. "
                "For new labels, Master Label is required: the label is added to CustomLabels.labels-meta.xml too.")
    c1, c2, _ = st.columns([1, 1, 3])
    c1.download_button("Download sample Excel", make_xlsx(True), "sample_custom_labels.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("Download blank template", make_xlsx(False), "translations_template.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("2. Salesforce project folder")
    folder = st.text_input("Folder that contains CustomLabels.labels-meta.xml (subfolders are searched too)",
                           placeholder="/Users/you/my-project/force-app/main/default",
                           help="In VS Code: right-click the folder > Copy Path.")
    if not folder.strip():
        st.info("Enter the folder path to continue.")
        return
    root = Path(folder.strip().strip("'\"")).expanduser()
    if not root.is_dir():
        st.error("That folder does not exist.")
        return
    found = scan(root)
    if not found:
        st.error("CustomLabels.labels-meta.xml was not found in that folder or its subfolders.")
        return
    labels_path = found[0] if len(found) == 1 else st.selectbox("CustomLabels file", found)
    trans_dir = Path(st.text_input("Translations folder", str(labels_path.parent.parent / "translations"))).expanduser()
    have = sorted(p.name.split(".")[0] for p in trans_dir.glob("*.translation-meta.xml")) if trans_dir.is_dir() else []
    st.success(f"Labels file: {labels_path}")
    st.caption("Existing translation files: " + (", ".join(have) if have else "none (new ones will be created)"))

    st.subheader("3. Excel file")
    up = st.file_uploader("Upload your Excel", type=["xlsx", "xlsm", "xls", "csv"])
    if not up:
        return
    if up.name.lower().endswith(".csv"):
        df = pd.read_csv(up, dtype=str)
    else:
        xl = pd.ExcelFile(up)
        sheet = st.selectbox("Sheet", xl.sheet_names) if len(xl.sheet_names) > 1 else xl.sheet_names[0]
        df = xl.parse(sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna("")
    df = df[df.apply(lambda r: r.str.strip().ne("").any(), axis=1)]
    cols = list(df.columns)

    a_guess = next((c for c in cols if re.search("api", c, re.I)), cols[0])
    k1, k2 = st.columns(2)
    api_col = k1.selectbox("API name column", cols, index=cols.index(a_guess))
    rest = [c for c in cols if c != api_col]
    m_opts = [NONE] + rest
    m_guess = next((c for c in rest if MASTER_RE.search(c)), NONE)
    master_col = k2.selectbox("Master label column", m_opts, index=m_opts.index(m_guess))
    lang_opts = [c for c in rest if c != master_col]
    langs = st.multiselect("Language columns", lang_opts, default=lang_opts, key=f"l_{api_col}_{master_col}")
    st.dataframe(df, hide_index=True)

    seen, blank, invalid = {}, 0, []
    for _, r in df.iterrows():
        n = r[api_col].strip()
        if not n:
            blank += 1
        elif not NAME_OK.fullmatch(n):
            invalid.append(n)
        else:
            seen[n] = (n, r[master_col].strip() if master_col != NONE else "", {c: r[c] for c in langs})
    rows = list(seen.values())
    if blank:
        st.warning(f"{blank} row(s) without an API name are skipped.")
    if invalid:
        st.error("Invalid API names skipped (letters, digits, single underscores; start with a letter): " + ", ".join(invalid))
    if langs:
        st.caption("Language files: " + ", ".join(f"{c} -> {lang_code(c)}.translation-meta.xml" for c in langs))

    b1, b2, _ = st.columns([1, 1, 3])
    preview = b1.button("Preview (writes nothing)", disabled=not rows)
    apply = b2.button("Apply changes to XML files", type="primary", disabled=not rows)
    if preview or apply:
        try:
            detail, files, changes = apply_changes(rows, langs, labels_path, trans_dir,
                                                   master_lang.strip() or "en_US", protected, ow_master, ow_trans, root)
            backup = None
            if apply:
                backup = make_backup(changes, root)  # original files are saved BEFORE anything is written
                for p, (_, new) in changes.items():
                    save(p, new)
        except Exception as e:  # noqa: BLE001
            st.error(f"Failed: {e}")
            return
        st.session_state["result"] = dict(applied=apply, detail=detail, files=files, backup=backup,
                                          changed=bool(changes))

    res = st.session_state.get("result")
    if res:
        if res["applied"] and res["changed"]:
            st.success("Files updated. Review the changes in VS Code (Source Control / diff) before deploying.")
        elif res["applied"]:
            st.info("Nothing needed to change. No files were modified.")
        else:
            st.info("Preview only. No files were changed.")
        bk = res["backup"]
        if bk:
            st.subheader("Backup of the original files")
            st.write("These are the files **as they were before this run**. Download them if you want to revert.")
            st.download_button("Download backup (zip)", bk["data"], bk["path"].name, "application/zip", type="primary")
            st.caption(f"Also saved on this computer: {bk['path']}")
            st.dataframe(pd.DataFrame([{"File": k, "Status": "original saved in backup"} for k in bk["originals"]] +
                                      [{"File": k, "Status": "new file created (delete it to revert)"} for k in bk["created"]]),
                         hide_index=True)
            for name, text in bk["originals"].items():
                with st.expander(f"Original, unchanged: {name}"):
                    st.code(text, language="xml")
            st.caption("To revert: unzip and copy the files back over the edited ones (see _BACKUP_INFO.txt in the zip).")
        st.write("**Per label**")
        st.dataframe(pd.DataFrame(res["detail"]), hide_index=True)
        st.write("**Per file**")
        st.dataframe(pd.DataFrame(res["files"]), hide_index=True)


if __name__ == "__main__":
    main()
