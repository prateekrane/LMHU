import logging
from .doc_builder import WordBuilder
from .ppt_builder import PPTBuilder

class CommandExecutor:
    def __init__(self):
        self.word_builder = WordBuilder()
        self.ppt_builder = PPTBuilder()

    def execute_commands(self, commands, doc_type, output_path):
        try:
            if doc_type == "word":
                for command in commands:
                    self.word_builder.run_command(command)
                self.word_builder.save(output_path)
            elif doc_type == "ppt":
                for command in commands:
                    self.ppt_builder.run_command(command)
                self.ppt_builder.save(output_path)
        except Exception as e:
            logging.error(f"Command execution error: {e}", exc_info=True)
            raise