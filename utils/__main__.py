import json
import logging

from utils.vba_doc_parser import VbaDocs

logging.basicConfig(
    level = logging.INFO,
    format = "%(asctime)s [%(levelname)s] %(message)s",
    handlers = [logging.FileHandler("utils.log", mode="w")],
)

docs = VbaDocs()
docs.read_directory("VBA-Docs/api")
docs.process_pages()
json.dump(docs.to_dict(), open("docs.json", "wt"), indent=4)
with open("pyvba/Excel.py", "wt") as f:
    f.write(docs.to_python())
