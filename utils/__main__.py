import json
import logging

from utils.vba_doc_parser import VbaDocs

logging.basicConfig(
    level = logging.INFO,
    format = "%(asctime)s [%(levelname)s] %(message)s",
    handlers = [logging.FileHandler("utils.log", mode="w")],
)

docs = VbaDocs()
docs.read_directory("office-vba-reference/api")
docs.process_pages()
json.dump(docs.to_dict(), open("office-vba-api.json", "wt"), indent=4)
