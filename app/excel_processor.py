import openpyxl

from datetime import datetime
from openpyxl.utils import get_column_letter

from app.config import (
    CITY_ORDER,
    CITY_MAPPING,
    MONTH_NAMES,
    WEEKDAYS,
    HEADERS,
    HEADER_FILL,
    HEADER_FONT,
    HEADER_ALIGNMENT,
    METRIC_FILL,
    METRIC_ALIGNMENT,
    DATA_ALIGNMENT,
    THICK_BORDER,
)
from logger.logger import setup_logger

logger = setup_logger(module_name=__name__)


class ExcelProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        try:
            self.wb = openpyxl.load_workbook(file_path)
        except FileNotFoundError:
            self.wb = openpyxl.Workbook()
            if "Sheet" in self.wb.sheetnames:
                del self.wb["Sheet"]
            self.wb.save(file_path)

        # устанавливаем ширину столбцов
        self.column_widths = {
            "A": 54.14,
            "B": 18.29,
            "C": 11.43,
            "D": 14.57,
            "E": 17.86,
            "F": 19.71,
        }

    def find_data_section(self, ws, target_date: datetime) -> int:
        date_str = target_date.strftime("%d.%m.%Y")
        weekday = WEEKDAYS[target_date.weekday()]
        search_value = f"{date_str} ({weekday})"

        for row in range(1, ws.max_row + 1):
            cell_value = ws.cell(row, 1).value
            if cell_value == search_value:
                return row
        return None

    def get_city_column(self, ws, row: int, city: str) -> int:
        city = city.lower().strip()
        for col in range(2, ws.max_column + 1):
            cell_value = ws.cell(row, col).value
            if cell_value:
                cell_city = str(cell_value).lower().strip()
                if city in cell_city:
                    return col
        return None

    def create_new_data_section(self, ws, date: datetime) -> int:
        last_row = ws.max_row

        # добавляет ТОЛЬКО ОДНУ пустую строку перед новой секцией
        if last_row > 0:
            # проверяем, что последняя строка не пустая
            if any(ws.cell(last_row, col).value for col in range(1, ws.max_column + 1)):
                ws.append([])
                last_row += 1

        date_row = last_row + 1
        date_str = date.strftime("%d.%m.%Y")
        weekday = WEEKDAYS[date.weekday()]
        ws.cell(date_row, 1).value = f"{date_str} ({weekday})"
        ws.cell(date_row, 1).fill = HEADER_FILL
        ws.cell(date_row, 1).font = HEADER_FONT
        ws.cell(date_row, 1).alignment = HEADER_ALIGNMENT
        ws.cell(date_row, 1).border = THICK_BORDER

        for col, city in enumerate(CITY_ORDER, start=2):
            ws.cell(date_row, col).value = city
            ws.cell(date_row, col).fill = HEADER_FILL
            ws.cell(date_row, col).font = HEADER_FONT
            ws.cell(date_row, col).alignment = HEADER_ALIGNMENT
            ws.cell(date_row, col).border = THICK_BORDER

        for i, metric in enumerate(HEADERS, start=1):
            cell = ws.cell(date_row + i, 1)
            cell.value = metric
            cell.fill = METRIC_FILL
            cell.alignment = METRIC_ALIGNMENT
            cell.border = THICK_BORDER

        return date_row

    def create_new_sheet(
        self, sheet_name: str
    ) -> openpyxl.worksheet.worksheet.Worksheet:
        ws = self.wb.create_sheet(sheet_name)
        logger.info(f"new list created {sheet_name}")

        for col_letter, width in self.column_widths.items():
            ws.column_dimensions[col_letter].width = width
        return ws

    def format_data_cell(self, cell, value, index: int):
        cell.border = THICK_BORDER
        cell.alignment = DATA_ALIGNMENT

        if index in [4, 5, 6]:
            cell.value = f"{value} сек."
        elif index == 7:
            cell.value = f"{value:.2f}%".replace(".", ",")
        else:
            cell.value = value

    def process_message(self, data: dict):
        try:
            date_str = data["Date"]
            city_key = data["City"]
            date = datetime.strptime(date_str, "%Y-%m-%d")
            city_name = CITY_MAPPING.get(city_key)

            if not city_name:
                logger.warning(f"⚠️ unknown city: {city_key}")
                return

            values = [
                int(data["ВСЕГО:"]),
                int(data["Потеряно:"]),
                int(data["Переведено:"]),
                int(data["Успешно завершено:"]),
                int(data["Клиенты, не дождавшиеся ответа, ждали в среднем:"]),
                int(data["В среднем клиенты ждут:"]),
                int(data["В среднем разговор длится:"]),
                (int(data["Потеряно:"]) / int(data["ВСЕГО:"])) * 100
                if int(data["ВСЕГО:"]) > 0
                else 0,
            ]

            sheet_name = f"{MONTH_NAMES[date.month]} {date.year}"

            if sheet_name not in self.wb.sheetnames:
                ws = self.create_new_sheet(sheet_name)
            else:
                ws = self.wb[sheet_name]

            logger.info(
                f"🔍 process data for {date_str} ({city_name}) on sheet {sheet_name}"
            )
            data_row = self.find_data_section(ws, date)

            if data_row is None:
                logger.info(f"⚠️ date {date_str} not fund, create new section")
                data_row = self.create_new_data_section(ws, date)
                logger.success(f"✅ new section created at row: {data_row}")

            city_col = self.get_city_column(ws, data_row, city_name)

            if not city_col:
                logger.warning(f"⚠️ col for city '{city_name}' not found, add new")
                # ищем последний занятый столбец в строке с датой
                last_col = 1
                for col in range(2, ws.max_column + 1):
                    if ws.cell(data_row, col).value:
                        last_col = col

                city_col = last_col + 1

                cell = ws.cell(data_row, city_col)
                cell.value = city_name
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = HEADER_ALIGNMENT
                cell.border = THICK_BORDER
                logger.info(
                    f"add new column {get_column_letter(city_col)} for city: {city_name}"
                )

            start_row = data_row + 1
            for i, value in enumerate(values):
                cell = ws.cell(row=start_row + i, column=city_col)
                self.format_data_cell(cell, value, i)

            logger.success(
                f"✅ data was add {get_column_letter(city_col)}{start_row + i}"
            )
            self.wb.save(self.file_path)
            logger.success("💾 Excel was saved")

        except Exception as e:
            logger.error(f"❌ error: {str(e)}")
            import traceback

            traceback.print_exc()
