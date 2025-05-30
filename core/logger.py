import logging
import os
from datetime import datetime


def setup_logging():
    logs_dir = os.path.join(os.getcwd(), "logs")
    os.makedirs(logs_dir, exist_ok=True)

    logging.basicConfig(
        filename=os.path.join(logs_dir, f"lmhu_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        filemode='w',
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logging.getLogger().addHandler(logging.StreamHandler())

