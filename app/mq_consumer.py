import aio_pika
import json

from app.config import RABBITMQ_URL, QUEUE_NAME
from app.excel_processor import ExcelProcessor
from logger.logger import setup_logger

logger = setup_logger(module_name=__name__)


async def process_messages(processor: ExcelProcessor):
    connection = await aio_pika.connect_robust(RABBITMQ_URL)
    queue_name = QUEUE_NAME

    async with connection:
        channel = await connection.channel()
        queue = await channel.declare_queue(queue_name, durable=True)

        logger.info(
            f"waiting for messages in the queue {queue_name}. To exit, press CTRL+C"
        )

        async for message in queue:
            async with message.process():
                try:
                    data = json.loads(message.body.decode())
                    logger.info(f"📩 received message: {data['Date']} {data['City']}")
                    processor.process_message(data)
                except Exception as e:
                    logger.error(f"❌ message processing error: {str(e)}")
                    import traceback

                    traceback.print_exc()
