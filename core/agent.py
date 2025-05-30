import logging
from core.generator_word import generate_word_doc
from core.generator_ppt import generate_ppt_doc


def execute_commands(command_json, doc_type):
    logging.info(f"Executing commands for {doc_type}...")

    try:
        if doc_type == "word":
            return generate_word_doc(command_json)
        elif doc_type == "ppt":
            return generate_ppt_doc(command_json)
        else:
            raise ValueError("Unknown document type")
    except Exception as e:
        logging.exception("Command execution failed")
        raise
