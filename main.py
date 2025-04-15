import os
import re
import logging
import pandas as pd
from datetime import datetime
from src.bot import WebAccess
from src.bot import SharePoint
from src.bot import Excel

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
SP_DOWNLOAD_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),"SharePoint")
def main():
    data = WebAccess(
        url="https://webaccess.nsk-cad.com",
        username="hanh0704",
        password="159753",
        headless=True,
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
    data = data[data["追加不足"] != "不足"]
    # Tải file báo giá
    SP = SharePoint(
        url="https://nskkogyo.sharepoint.com/",
        username="vietnamrpa@nskkogyo.onmicrosoft.com",
        password="Robot159753",
        download_directory=SP_DOWNLOAD_PATH,
    )
    prices  = []
    for url in data["資料リンク"].to_list():
        statuses:list[tuple[str,str]] = SP.download_file(
            site_url=url, file_pattern="見積書/.*.(xlsm|xlsx|xls)$"
        )
        status = statuses[0]
        if status[1]:
            prices.append(status[1])
        else:
            value = Excel(
                file_path=os.path.join(SP_DOWNLOAD_PATH,status[0])
            ).search_keyword(
                sheetname="見積書 (3)",
                keyword="税抜金額",
                axis=1,
            )    
            prices.append(value)
    data['金額（税抜）'] = prices
    del data['資料リンク']
    # Process data
    data['金額（税抜）'] = data['金額（税抜）'].astype(str).str.replace(r"[^\d.,]", "", regex=True)
    # To excel
    resultFile = f"{datetime.today().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx"
    data.to_excel(resultFile, index=False)
    logger.info(f"Kết quả: {resultFile}")
    # Setup Excel File
    excel = Excel(
        file_path=os.path.join(SP_DOWNLOAD_PATH,resultFile),
        timeout=10,
        retry_interval=0.5,
    )
    excel.page_setup(
        orientation="Landscape",
        header="さくら建設　鋼製野縁納材報告（2025/03/21-2025/04/20）　",
    )
if __name__ == "__main__":
    main()
