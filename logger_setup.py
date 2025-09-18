import logging
import sys
import os

def setup_logging(log_file="daily_report.log"):
    # Ensure directory exists
    os.makedirs(os.path.dirname(log_file), exist_ok=True) if os.path.dirname(log_file) else None

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file, mode="a", encoding="utf-8", errors="replace"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logger = logging.getLogger("global_logger")

    # Handle uncaught exceptions
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

    sys.excepthook = handle_exception
    return logger

# Initialize once when imported
logger = setup_logging()