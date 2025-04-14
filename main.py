import logging
import pandas as pd
from uuid import uuid4
from src.bot import WebAccess
from src.bot import SharePoint

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


def main():
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
    data = data[data["追加不足"] != "不足"]
    data = data.head(1)
    # Tải file báo giá
    SP = SharePoint(
        url="https://nskkogyo.sharepoint.com/",
        username="vietnamrpa@nskkogyo.onmicrosoft.com",
        password="Robot159753",
        download_directory="D:/VanNgocNhatHuy/RPA/Report/SharePoint",
    )
    statuses = []
    for url in data["資料リンク"].to_list():
        status = SP.download_file(
            site_url=url, file_pattern="見積書/.*.(xlsm|xlsx|xls)$"
        )
        if status[1]:
            statuses.append(status[1])
        else:
            statuses.append(status[0])

    # To excel
    data["File"] = statuses
    data.to_excel(f"{uuid4()}.xlsx", index=False)


if __name__ == "__main__":
    main()
