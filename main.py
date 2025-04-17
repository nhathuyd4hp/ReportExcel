import os
import re
import logging
import pandas as pd
from src.bot import (
    Excel,
    SharePoint,
    WebAccess,
)
from datetime import datetime
from src.common.decorator import handle_error_func
from openpyxl.utils import column_index_from_string, get_column_letter

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    encoding="utf-8",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler("bot.log", mode="a", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)
SP_DOWNLOAD_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "SharePoint"
)


@handle_error_func()
def main(**kwargs):
    data = WebAccess(
        url="https://webaccess.nsk-cad.com",
        username="hanh0704",
        password="159753",
    ).get_information(
        builder_name="009300",
        delivery_date=["2025/03/21", "2025/04/20"],
        fields=[
            "案件番号",
            "得意先名",
            "物件名",
            "確未",
            "確定納期",
            "曜日",
            "追加不足",
            "配送先住所",
            "階",
            "資料リンク",
        ],
    )
    if not isinstance(data, pd.DataFrame):
        return
    # Xóa những dòng 不足
    data = data[data["追加不足"] != "不足"].head(2)
    # Tải file báo giá
    SP = SharePoint(
        url="https://nskkogyo.sharepoint.com/",
        username="vietnamrpa@nskkogyo.onmicrosoft.com",
        password="Robot159753",
        download_directory=SP_DOWNLOAD_PATH,
    )
    prices = []
    for url in data["資料リンク"].to_list():
        statuses: list[tuple[str, str]] = SP.download_file(
            site_url=url,
            file_pattern="見積書/.*.(xlsm|xlsx|xls)$",
        )
        status = statuses[0]
        if status[1]:
            prices.append(status[1])
        else:
            price = Excel(
                file_path=os.path.join(SP_DOWNLOAD_PATH, status[0])
            ).search(
                sheetname="見積書 (3)",
                keyword="税抜金額",
                axis=1,
            )
            price = re.sub(r"[^\d.,]", "", price) # Convert to numberic
            prices.append(price)
    data["金額（税抜）"] = prices
    data["金額（税抜）"] = pd.to_numeric(data["金額（税抜）"], errors='coerce')
    del data["資料リンク"]
    # Save 
    excelFile = f"{datetime.today().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    logger.info(f"Excel File: {excelFile}")
    data.to_excel(excelFile, index=False)
    # ------------ # 
    # Setup
    excel = Excel(
        file_path=excelFile,
        timeout=10,
        retry_interval=0.5,
    )
    last_column: str = re.findall(r"[A-Z]+", excel.shape[1])[0]
    last_row: int = int(re.findall(r"\d+", excel.shape[1])[0])
    excel.edit(
        cells=[
            "A1","B1","C1","D1","E1","F1","G1","H1","I1","J1"
        ],
        background_colors=["A6A6A6","A6A6A6","A6A6A6","A6A6A6","A6A6A6","A6A6A6","A6A6A6","A6A6A6","A6A6A6","A6A6A6"],
    )
    excel.edit(
        cells=[
            "{column}{row}".format(
            column=get_column_letter(column_index_from_string(last_column) - 2),
            row=last_row + 3,
            ),
            "{column}{row}".format(column=last_column, row=last_row + 3)
        ],
        contents=[
            "合計",
            "=SUM({from_cell}:{to_cell})".format(
            from_cell=f"{last_column}2",
            to_cell=f"{last_column}{last_row}",
            )
        ],
        background_colors=["A6A6A6"],
    )
    excel.page_setup(
        orientation="Landscape",
        header="さくら建設　鋼製野縁納材報告（2025/03/21-2025/04/20）　",
    )
    excel.format(
        AutoFitColumnWith=True,
        Border="All Borders",
    )
    excel.save()
    file = excel.export()
    logger.info(f"PDF File: {file}")

    

if __name__ == "__main__":
    main(
        logger=logger,
    )
