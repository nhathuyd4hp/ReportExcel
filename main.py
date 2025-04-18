import os
import re
import logging
import pandas as pd
from src.bot import (
    Excel,
    SharePoint,
    WebAccess,
    MailDealer,
)
import datetime as dt
from datetime import datetime
from src.common.decorator import HandleExceptionFunc
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


@HandleExceptionFunc()
def main(**kwargs):    
    # ---- #
    today = datetime.now()
    to_date = today.replace(day=20)
    from_date = None
    if today.month == 1:
        from_date = today.replace(year=today.year - 1, month=12, day=21)
    else:
        from_date = today.replace(month=today.month - 1, day=21)
    to_date = to_date.strftime("%Y/%m/%d")
    from_date = from_date.strftime("%Y/%m/%d")
    # ---- #
    logger: logging.Logger = kwargs.get("logger", logging.getLogger(__name__))
    download_path = kwargs.get(
        "download_path", os.path.dirname(os.path.abspath(__file__))
    )
    # ----#
    data = WebAccess(
        url="https://webaccess.nsk-cad.com",
        username="hanh0704",
        password="159753",
    ).get_information(
        builder_name="009300",
        delivery_date=[from_date, to_date],
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
        download_directory=download_path,
    )
    prices = []
    for url in data["資料リンク"].to_list():
        statuses: list[tuple[str, str]] = SP.download_file(
            site_url=url,
            file_pattern="見積書/.*.(xlsm|xlsx|xls)$",
        )
        if not statuses:
            prices.append(None)
        else:
            status = statuses[0]
            if status[1]:
                prices.append(status[1])
            else:
                price = Excel(file_path=os.path.join(download_path, status[0])).search(
                    sheetname="見積書 (3)",
                    keyword="税抜金額",
                    axis=1,
                )
                price = re.sub(r"[^\d.,]", "", price)  # Convert to numberic
                prices.append(price)
    data["金額（税抜）"] = prices
    data["金額（税抜）"] = pd.to_numeric(data["金額（税抜）"], errors="coerce").fillna(0)  # convert to numberic
    del data["資料リンク"]
    # Save
    excelFile = f"{datetime.today().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    data.to_excel(os.path.abspath(excelFile), index=False)
    logger.info(f"Excel File: {os.path.abspath(excelFile)}")
    # ------------ #
    # Setup
    excel = Excel(
        file_path=os.path.abspath(excelFile),
        timeout=10,
        retry_interval=0.5,
    )
    last_column: str = re.findall(r"[A-Z]+", excel.shape[1])[0]
    last_row: int = int(re.findall(r"\d+", excel.shape[1])[0])
    excel.edit(
        cells=[
            "A1",
            "B1",
            "C1",
            "D1",
            "E1",
            "F1",
            "G1",
            "H1",
            "I1",
            "J1",
            "{column}{row}".format(
                column=get_column_letter(column_index_from_string(last_column) - 2),
                row=last_row + 3,
            ),
            "{column}{row}".format(column=last_column, row=last_row + 3),
        ],
        contents=[
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            "合計",
            "=SUM({from_cell}:{to_cell})".format(
                from_cell=f"{last_column}2",
                to_cell=f"{last_column}{last_row}",
            ),
        ],
        background_colors=[
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
            "A6A6A6",
        ],
    )
    excel.format_cells(
        range=f"{excel.shape[1][0]}2:{excel.shape[1]}",
        tab="Number",
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
    # pdfFile = excel.export(
    #     file_name=f"さくら建設　鋼製野縁納材報告（{from_date} - {to_date}).pdf"
    # )
    # logger.info(f"PDF File: {os.path.abspath(pdfFile)}")
    # Send mail
    # MailDealer(
    #     username="vietnamrpa",
    #     password="nsk159753",
    # ).send_mail(
    #     from_email="kantou@nsk-cad.com",
    #     to_email="ikeda.k@jkenzai.com",
    #     content="""ジャパン建材　池田様		
		
    #     いつもお世話になっております。		
		
    #     さくら建設　鋼製野縁納材報告書（2025,0221～0320）		
    #     を送付致しましたので、ご査収の程よろしくお願い致します。		
                
    #     ----・・・・・----------・・・・・----------・・・・・-----		
                
    #     　エヌ・エス・ケー工業㈱　横浜営業所		
    #     中山　知凡		
    #     　		
    #     　〒222-0033		
    #     　横浜市港北区新横浜２-４-６　マスニ第一ビル８F-B		
    #     　TEL:(045)595-9165 / FAX:(045)577-0012		
    #     　		
    #     -----・・・・・----------・・・・・----------・・・・・-----		
    #     """,
    #     attachments=os.path.abspath(pdfFile),
    # )


if __name__ == "__main__":
    main(
        logger=logging.getLogger(__name__),
        download_path=os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "SharePoint"
        ),
    )
