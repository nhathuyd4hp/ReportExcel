import os
import time
import logging
from typing import Literal
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from pywinauto import Application,Desktop
from pywinauto.application import WindowSpecification
from pywinauto.timings import wait_until_passes
from pywinauto.controls.uiawrapper import UIAWrapper
from src.common.decorator import safe_window

class Excel:
    @safe_window()
    def __init__(
        self,
        file_path:str,
        timeout:int=10,
        retry_interval:float=0.5,
        logger_name:str=__name__,
    ):
        self.logger = logging.getLogger(logger_name)
        if not os.path.exists(file_path):
            self.logger.error(f"{file_path} không tồn tại")
            raise FileNotFoundError(f"{file_path} không tồn tại")
        self.file_path = file_path
        self.timeout= timeout
        self.retry_interval=retry_interval
        Application(backend="uia").start(
            r'"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" "{}"'.format(file_path)
        )
        while True:
            titles = [w.window_text() for w in Desktop(backend="uia").windows()]
            if found := [title for title in titles if title.startswith(os.path.basename(self.file_path))]:
                self.title:str = found[0]
                break
            continue
        self.App: Application = Application(backend="uia").connect(title=self.title)
        if self.title.find("[Protected View]") != -1:
            ExcelApp:WindowSpecification = self.App.window(title=self.title)
            EnableEditingButton:WindowSpecification = ExcelApp.child_window(title="Enable Editing", control_type="Button")
            wait_until_passes(
                timeout=self.timeout,
                retry_interval=self.retry_interval,
                func=lambda: EnableEditingButton.exists(timeout=self.timeout),
            )
            EnableEditingButton.wrapper_object().click_input(double=True)
        time.sleep(5)
        self.title = self.App.window().window_text()
        pass


    def __wait_for_exists(self,window:WindowSpecification) -> UIAWrapper | None:
        try:
            wait_until_passes(
                timeout=self.timeout,
                retry_interval=self.retry_interval,
                func=lambda: window.exists(timeout=self.timeout),
            )
            return window.wrapper_object()
        except Exception as e:
            self.logger.error(e)
            return None
        
    @safe_window()
    def page_setup(
        self,
        orientation:Literal["Landscape","Portrait"] = "Portrait",
        page_size:Literal["A3","A4","A5","B4 (JIS)","Statement","Executive"]="A4",
        header:str=None,
        footer:str=None,
    ) -> bool:
        ExcelApp:WindowSpecification = self.App.window(title=self.title)
        self.__wait_for_exists(
            ExcelApp.child_window(title="Page Layout", auto_id="TabPageLayoutExcel", control_type="TabItem")
        ).select()
        self.__wait_for_exists(
            ExcelApp.child_window(title="Print Titles", auto_id="PrintTitles", control_type="Button")
        ).click_input()
        while not self.App.window(title="Page Setup").exists():
            continue
        dialog = self.App.window(title="Page Setup")

        self.__wait_for_exists(
            dialog.child_window(title="Page", control_type="TabItem")
        ).select()
        self.__wait_for_exists(
            dialog.child_window(title=orientation, control_type="RadioButton")
        ).click_input() 
        self.__wait_for_exists(
            dialog.child_window(title="Paper size:", control_type="ComboBox")
        ).select(page_size).click_input()
        # 
        if header or footer:
            self.__wait_for_exists(
                dialog.child_window(title="Header/Footer", control_type="TabItem")
            ).select()
            try:
                self.__wait_for_exists(
                    dialog.child_window(title="Header:", control_type="ComboBox")
                ).select(header).click_input()
            except IndexError:
                self.__wait_for_exists(
                    dialog.child_window(title="Header/Footer", control_type="TabItem")
                ).select()
                self.__wait_for_exists(
                    dialog.child_window(title="Custom Header...", control_type="Button")
                ).click_input(double=True)
                self.__wait_for_exists(
                    dialog.child_window(title="Left section:", control_type="Edit")
                ).type_keys(header)
                pass
        self.__wait_for_exists(
            dialog.child_window(title="OK", control_type="Button")
        ).click_input()
        return True

    def search_keyword(self,sheetname:str,keyword:str,axis:int=0) -> str | None:
        workbook = load_workbook(
            filename=self.file_path,
            data_only=True,
        )
        worksheet = workbook[sheetname]
        target_cell:Cell|None = None
        cells: tuple[Cell, ...] = None
        if axis == 0:  # Tìm kiếm theo cột dọc (column)
            for col in worksheet.iter_cols():
                for cell in col:
                    if cell.value == keyword:
                        target_cell = cell
                        break
                if target_cell:
                    cells = [
                        row[0] for row in worksheet.iter_rows(min_col=target_cell.column,max_col=target_cell.column)
                    ]
        if axis == 1:  # Tìm kiếm theo hàng ngang (row)
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value == keyword:
                        target_cell = cell
                        break
                if target_cell:
                    cells = worksheet[target_cell.row]
                    break
        if not(target_cell and cells):
            self.logger.error(f"Không tìm thấy keyword: {keyword}")
            return None
        cells = [cell.value for cell in cells if cell.value is not None]
        values: list = cells[cells.index(keyword)+1:]
        return " ".join(str(value) for value in values)