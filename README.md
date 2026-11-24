# Salesforce Custom Label Translation Filler

Reads translations from an Excel sheet and generates Salesforce
`<lang>.translation-meta.xml` files for your Custom Labels, one file per
language. The files can then be deployed to your org with the Salesforce CLI.

The script works only on local files. It never connects to Salesforce and needs
no credentials. Org access is only needed at the deploy step (see below).

## Features

- Fills translations for multiple languages in one run
- Accepts language codes (`fr`, `pt_BR`) or common names (`French`, `German`) as column headers
- Optionally limits the run to specific custom label API names
- Merges into existing translation files: other content is kept, only the given labels are added or updated
- Skips blank cells and warns about API names missing from the Excel sheet
- Keeps labels sorted alphabetically, matching Salesforce's metadata format

## Repository contents

```
.
├── app.py                          # Streamlit app: edits CustomLabels + translation XML in place
├── label_sheet_builder.html        # browser alternative (Chrome/Edge), also edits in place
├── fill_label_translations.py      # the script
├── requirements.txt                # Python dependencies
├── sample/
│   ├── translations_template.xlsx  # blank template (API Name, Master Label, fr, de, es)
│   └── sample_custom_labels.xlsx   # same format with 10 sample labels (fr, de, es, fi, fr_BE, nl_NL)
├── .gitignore
└── README.md
```

Start from `sample/translations_template.xlsx`: add your label API names in column A and translations under each language column. Add or remove language columns as needed.

## Requirements

- Python 3.8+
- `pandas` and `openpyxl`

```bash
pip install -r requirements.txt
```

To deploy: [Salesforce CLI](https://developer.salesforce.com/tools/salesforcecli) (`npm install -g @salesforce/cli`)

## Streamlit app (recommended): update your project's XML in place

`app.py` opens a local web page. You give it your Salesforce project folder and an Excel file, and it edits `CustomLabels.labels-meta.xml` and the translation files directly. Review the result with VS Code's diff, then deploy.

### Run

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

The browser opens at `http://localhost:8501`. (Next time, only `source .venv/bin/activate` and the last command are needed.)

### Excel format

First row = headers. A sample and a blank template can be downloaded from the app itself (and are in `sample/`).

| API Name | Master Label | fr | de | es |
|---|---|---|---|---|
| Welcome_Message | Welcome to our portal | Bienvenue sur notre portail | Willkommen auf unserem Portal | Bienvenido a nuestro portal |

- **Label already in `CustomLabels.labels-meta.xml`:** only its translations are added or updated. Master Label can be left blank.
- **New label (not in the file yet):** Master Label is required. The label is added to `CustomLabels.labels-meta.xml` and its translations are added. A new label without a Master Label is skipped and reported, because translating a label that does not exist would fail to deploy.
- **Language columns:** header = Salesforce language code (`fr`, `de`, `pt_BR`, `zh_CN`) or a common name such as French. Add or delete columns as needed. Blank cells are skipped.

### Steps in the app

1. Paste the **project folder path** (in VS Code: right-click the folder, Copy Path). The folder may be `force-app/main/default` or the whole project; subfolders are searched for `CustomLabels.labels-meta.xml`. The translations folder defaults to `translations` next to the `labels` folder and can be changed.
2. **Upload the Excel file.** Check the detected API name, master label and language columns.
3. Click **Preview (writes nothing)** to see, per label and per file, what will be added or updated.
4. Click **Apply changes to XML files**. The originals are backed up first (zip, downloadable), then the files are edited in place.
5. Review in VS Code, then deploy.

Settings in the sidebar: master label language (default `en_US`), protected flag for new labels, whether to overwrite the master value of existing labels (off by default), and whether to overwrite existing translations (on by default).

### Backup before changes

When you click **Apply**, the app first saves the **original, unchanged** version of every file it is about to edit, and only then writes the changes.

- A zip named `backup_YYYYMMDD_HHMMSS.zip` is saved in a `backups/` folder next to `app.py` (ignored by git).
- The same zip is offered as a **Download backup (zip)** button right after the run, and the originals are shown in the app (one collapsed section per file).
- The zip contains the original files with their folder paths, plus `_BACKUP_INFO.txt` that lists where each file belongs and which files were newly created.
- **To revert:** unzip and copy the files back over the edited ones. Delete any files listed as newly created (for example a new `de.translation-meta.xml`).
- Only files that actually change are backed up. Preview never writes or backs up anything.

### What changes in the files

- Only the label concerned is edited, as text. Indentation, line endings and everything else stay as they were, so the diff shows only real changes.
- New labels and translations are inserted in alphabetical position. New `CustomLabels` entries get `fullName`, `language`, `protected`, `shortDescription` (the API name with underscores as spaces) and `value`.
- If a language file does not exist, it is created.
- Apostrophes and quotes are written as `&apos;` and `&quot;`, as Salesforce does.
- Running the same sheet twice changes nothing the second time.

## Alternative: browser page (no Python needed)

`label_sheet_builder.html` does the same in-place editing without Python. Open it in Chrome or Edge, load the Excel, pick the project folder and click Apply. It does not use the format above: the master labels go in one column and, if there is no API name column, you type the API names into its table.

## Alternative: command-line script

### Excel format

The first row is the header.

| API Name        | fr        | German     | es_MX      |
|-----------------|-----------|------------|------------|
| Welcome_Message | Bienvenue | Willkommen | Bienvenido |
| Submit_Button   | Soumettre | Absenden   | Enviar     |

- **First column:** custom label API name (the header text is ignored)
- **Other columns:** one per language. The header can be a Salesforce language
  code or a common language name. Names that aren't recognised are used as-is
  as the code, so use the code if in doubt.

Common names are mapped in the `LANG_NAMES` dictionary at the top of the script. Add your own entries if needed.

## Usage

```bash
python fill_label_translations.py -x translations.xlsx [options]
```

| Option | Description |
|---|---|
| `-x`, `--excel` | Excel file with translations (required) |
| `-s`, `--sheet` | Sheet name or index (default: first sheet) |
| `-n`, `--names` | Custom label API names to process, space separated |
| `-f`, `--names-file` | Text file with one API name per line (lines starting with `#` are ignored) |
| `-l`, `--languages` | Only process these language codes or headers |
| `-o`, `--out` | Output folder (default: `force-app/main/default/translations`) |

If neither `-n` nor `-f` is given, every label in the sheet is processed.

### Examples

```bash
# All labels, all languages
python fill_label_translations.py -x translations.xlsx

# Specific labels only
python fill_label_translations.py -x translations.xlsx -n Welcome_Message Submit_Button

# Labels listed in a file, French and German only, custom output folder
python fill_label_translations.py -x translations.xlsx -f label_names.txt -l fr de -o ./translations
```

### Output

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Translations xmlns="http://soap.sforce.com/2006/04/metadata">
    <customLabels>
        <label>Bienvenue</label>
        <name>Welcome_Message</name>
    </customLabels>
</Translations>
```

Console output reports what was written per language, for example:

```
[fr] 2 label(s) written -> force-app/main/default/translations/fr.translation-meta.xml
[de] 1 label(s) written -> force-app/main/default/translations/de.translation-meta.xml  (blank, skipped: Submit_Button)
```

## Deploying to Salesforce

### 1. Prerequisites in the org

- The custom labels must already exist.
- Translation Workbench is enabled (Setup → Translation Language Settings).
- Each language you are deploying is added and active there.
- Your user has "Modify Metadata Through Metadata API Functions" (or Modify All Data / System Administrator), plus "Customize Application" and "Manage Translation".

### 2. Log in (once)

```bash
sf org login web --alias myOrg
# Sandbox:
sf org login web --alias mySandbox --instance-url https://test.salesforce.com
```

This opens a browser for you to sign in. The CLI stores an OAuth token locally and reuses it.

### 3. Deploy

```bash
sf project deploy start -d force-app/main/default/translations -o myOrg
```

Run this from inside an SFDX project (a folder containing `sfdx-project.json`).

### CI/CD (no browser login)

Authenticate non-interactively instead:

```bash
# JWT with a Connected App
sf org login jwt --client-id <consumer_key> --jwt-key-file server.key \
  --username you@company.com --instance-url https://login.salesforce.com

# or an SFDX auth URL stored as a pipeline secret
sf org login sfdx-url --sfdx-url-file authUrl.txt --alias myOrg
```

## Alternative: Translation Workbench import

Admins who prefer the UI can use Setup → Translation Workbench → Export / Import,
which uses `.stf` files. This script does not generate `.stf` files.

## Troubleshooting

| Problem | Likely cause |
|---|---|
| `WARNING: not found in Excel: ...` | The API name is misspelled or missing from the first column. Names are case-sensitive. |
| Language skipped with "no translations" | The column is empty for the selected labels. |
| Deploy error about the language | The language isn't enabled in Translation Workbench. |
| Deploy error about a label | The custom label doesn't exist in the target org yet. Deploy the labels first. |
| Insufficient access error on deploy | Your user lacks the Metadata API or Manage Translation permissions listed above. |

## Notes

- Duplicate API names in the Excel sheet: the last row wins.
- Existing translations for the same label and language are overwritten.
- Empty cells never overwrite existing translations.
