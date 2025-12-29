import json
import logging
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from vipera.reference.parser import VbaDocs

def main():
    os.makedirs("logs", exist_ok=True)
    logging.basicConfig(
        level = logging.INFO,
        format = "%(asctime)s [%(levelname)s] %(message)s",
        handlers = [logging.FileHandler("logs/utils.log", mode="w")],
    )

    docs = VbaDocs()
    docs.read_directory("office-vba-reference/api")
    docs.process_pages()
    os.makedirs("data", exist_ok=True)
    json.dump(docs.to_dict(), open("data/office-vba-api.json", "wt"), indent=4)

if __name__ == "__main__":
    main()
