# wan-ocr
## Booklist Factory

Install and start the local dashboard from this directory.

```sh
python3 -m pip install "python-docx>=1.2" "pandas>=2"
python3 booklists.py --self-test
python3 factory.py
```

Open http://127.0.0.1:8765 on the same computer. Start with 3–5 files. Choose a Booklists folder containing one folder per region, or select individual Word files and enter their region.

The factory reads the 50-column booklists, checks quantities and duplicate school names, then publishes a cumulative analytics bucket. A later upload with the same region and filename replaces that source in subsequent totals. Files with other names remain separate submissions. Check flagged entries and duplicate candidates before using the totals. Download all audit and summary files from the bucket.

Source versions and completed reports persist under `~/.wan-booklists`. Set `WAN_FACTORY_DATA` to choose a different archive (for example, a separate pilot archive), or `WAN_FACTORY_PORT` to choose another local port. Run one factory process per archive.
