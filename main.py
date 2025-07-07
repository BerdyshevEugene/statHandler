import asyncio

from app.config import EXCEL_PATH
from app.excel_processor import ExcelProcessor
from app.mq_consumer import process_messages
from logger.logger import setup_logger

logger = setup_logger(__name__)

if __name__ == "__main__":
    processor = ExcelProcessor(EXCEL_PATH)
    need_save = True
    try:
        asyncio.run(process_messages(processor))
    except KeyboardInterrupt:
        logger.info("application stopped by user")
        need_save = False
    except Exception as e:
        logger.error(f"critical error: {str(e)}")
        import traceback

        traceback.print_exc()
    finally:
        if need_save:
            try:
                processor.wb.save(EXCEL_PATH)
                logger.info("excel file saved before exit")
            except Exception as e:
                logger.error(f"error while saving excel file: {str(e)}")
