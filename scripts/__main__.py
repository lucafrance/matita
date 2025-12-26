import json
import logging
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from pyvba_parser.parser import VbaDocs

logging.basicConfig(
    level = logging.INFO,
    format = "%(asctime)s [%(levelname)s] %(message)s",
    handlers = [logging.FileHandler("utils.log", mode="w")],
)

docs = VbaDocs()
docs.read_directory("office-vba-reference/api")
docs.process_pages()
json.dump(docs.to_dict(), open("office-vba-api.json", "wt"), indent=4)