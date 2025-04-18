import os
import re
from pathlib import Path
import time
import logging
from typing import Union
import pandas as pd
from functools import wraps
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from pywinauto.application import WindowSpecification
from pywinauto import Application
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.timings import wait_until_passes
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from src.common.decorator import retry
from src.common.decorator import require_authentication


def switch_to_default_content(func):
    """
    Decorator để tự động chuyển về default_content của trình duyệt.
    """

    @wraps(func)
    def wrapper(self, *args, **kwargs):
        self.browser.switch_to.default_content()
        result = func(self, *args, **kwargs)
        return result

    return wrapper


class MailDealer:
    def __init__(
        self,
        username: str,
        password: str,
        timeout: int = 5,
        headless: bool = False,
        logger_name: str = __name__,
    ):
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-notifications")
        # Disable log
        options.add_argument("--disable-logging")
        options.add_argument("--log-level=3")  #
        options.add_argument("--silent")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        if headless:
            options.add_argument("--headless=new")
        # Attribute
        self.logger = logging.getLogger(logger_name)
        self.browser = webdriver.Chrome(options=options)
        self.browser.maximize_window()
        self.timeout = timeout
        self.wait = WebDriverWait(self.browser, timeout)
        self.username = username
        self.password = password
        # Trạng thái đăng nhập
        self.authenticated = self.__authentication(username, password)
        # Root Window
        self.root_window = self.browser.window_handles[0]

    def __del__(self):
        if hasattr(self, "browser") and isinstance(self.browser, WebDriver):
            self.browser.quit()

    @retry(
        exceptions=(StaleElementReferenceException, ElementClickInterceptedException)
    )
    @switch_to_default_content
    def __authentication(self, username: str, password: str) -> bool:
        time.sleep(0.5)
        self.browser.get("https://mds3310.maildealer.jp/")
        username_field = self.wait.until(
            EC.presence_of_element_located((By.ID, "fUName")),
        )
        username_field.send_keys(username)

        password_field = self.wait.until(
            EC.presence_of_element_located((By.ID, "fPassword")),
        )
        password_field.send_keys(password)

        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input[value='ログイン']"),
            ),
        )
        login_btn = self.wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[value='ログイン']"),
            ),
        )
        login_btn.click()
        try:
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div[class='d_error_area ']"),
                ),
            )
            self.logger.error(
                "❌ Xác thực thất bại! Kiểm tra thông tin đăng nhập.",
            )
            return False
        except TimeoutException:
            if self.browser.current_url.find("app") != -1:
                self.logger.info("✅ Xác thực thành công!")
                return True
            return False

    @require_authentication
    @switch_to_default_content
    def __open_mail_box(self, mail_box: str, tab: Union[str, None] = None) -> bool:
        # --------------#
        if not self.wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.CSS_SELECTOR, "iframe[id='ifmSide']"),
            ),
        ):
            self.logger.error("Không tìm thấy Frame MailBox")
            return False
        mail_boxs: list[str] = mail_box.split("/")
        for box in mail_boxs:
            try:
                span_box = self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, f"span[title='{box}']"),
                    ),
                )
                span_box.click()
                time.sleep(1)
            except TimeoutException:
                self.logger.error(f"Không tìm thấy hộp thư {box}")
                return False
        self.browser.switch_to.default_content()
        # --------------#
        if not self.wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.CSS_SELECTOR, "iframe[id='ifmMain']"),
            ),
        ):
            self.logger.error("Không thể tìm thấy mailbox")
            return False
        # --------------#
        if tab:
            try:
                self.wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            f".//span[@class='olv-c-tab__name' and text()='{tab}']",
                        ),
                    ),
                ).click()
            except TimeoutException:
                self.logger.error(f"❌ Không tìm thấy Tab {tab}")
                return False
        self.browser.switch_to.default_content()
        return True

    @retry(exceptions=(StaleElementReferenceException, ValueError))
    @require_authentication
    @switch_to_default_content
    def mailbox(
        self, mail_box: str, tab_name: Union[str, None] = None
    ) -> pd.DataFrame | None:
        if not self.__open_mail_box(mail_box, tab_name):
            return None
        time.sleep(2)
        if not self.wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.CSS_SELECTOR, "iframe[id='ifmMain']"),
            ),
        ):
            self.logger.error("❌ Không thể tìm thấy Content Iframe!.")
            return None
        thead = self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "thead")))
        labels = thead.find_elements(By.TAG_NAME, "th")
        # Lọc lấy các thẻ label
        columns = []
        index_value = []
        for index, label in enumerate(labels):
            if label.find_elements(By.XPATH, "./*") and label.text:
                columns.append(label.text)
                index_value.append(index)
        df = pd.DataFrame(columns=columns)
        try:
            self.wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[text()='条件に一致するデータがありません。']")
                )
            )
            self.logger.info(f"✅ Hộp thư: {mail_box} rỗng")
            return df
        except Exception:
            tbodys = self.wait.until(
                EC.presence_of_all_elements_located((By.TAG_NAME, "tbody"))
            )
            for tbody in tbodys:
                row = []
                values: list[WebElement] = tbody.find_elements(By.TAG_NAME, "td")
                values: list[WebElement] = [values[index] for index in index_value]
                for value in values:
                    row.append(value.text)
                df.loc[len(df)] = row
            self.logger.info(f"✅ Lấy hộp thư: {mail_box}, tab: {tab_name}: thành công")
            return df

    @retry()
    @require_authentication
    def read_mail(self, mail_box: str, mail_id: str, tab_name: str = None) -> str:
        content = ""
        if not self.browser.current_url.startswith(
            "https://mds3310.maildealer.jp/app/"
        ):
            self.__authentication(self.username, self.password)
        self.__open_mail_box(
            mail_box=mail_box,
            tab=tab_name,
        )
        self.wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.CSS_SELECTOR, "iframe[id='ifmMain']"),
            ),
        )
        email_span = self.wait.until(
            EC.presence_of_element_located((By.XPATH, f"//span[text()='{mail_id}']"))
        )
        email_span.click()
        try:
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.CSS_SELECTOR, "iframe[id='html-mail-body-if']")
                )
            )
            ps = self.wait.until(
                EC.presence_of_all_elements_located((By.TAG_NAME, "p"))
            )
            for p in ps:
                content += p.text + "\n"
        except TimeoutException:
            body = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div[class='olv-p-mail-view-body']")
                )
            )
            content = body.find_element(By.TAG_NAME, "pre").text
        self.logger.info(
            f"✅ Đã đọc được nội dung mail:{mail_id}. tab: {tab_name} ở box:{mail_box}"
        )
        return content

    @retry(
        exceptions=(
            StaleElementReferenceException,
            ElementClickInterceptedException,
        )
    )
    @require_authentication
    def 一括操作(
        self,
        案件ID: any,
        このメールと同じ親番号のメールをすべて関連付ける: bool = False,
    ) -> tuple[bool, str]:
        try:
            popup = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div[class='pop-panel__content']")
                )
            )
            input = popup.find_element(By.ID, "fMatterID_add")
            button = input.find_element(By.XPATH, "./ancestor::*[1]//button")
            このメールと同じ親番号のメールをすべて関連付ける_div = popup.find_element(
                By.XPATH,
                "//div[text()='このメールと同じ親番号のメールをすべて関連付ける']",
            )
            div_checkbox = (
                このメールと同じ親番号のメールをすべて関連付ける_div.find_element(
                    By.XPATH, "./ancestor::*[1]//div"
                )
            )
            div_input = (
                このメールと同じ親番号のメールをすべて関連付ける_div.find_element(
                    By.XPATH, "./ancestor::*[1]//input"
                )
            )
            time.sleep(0.5)
            if (
                div_input.is_selected()
                != このメールと同じ親番号のメールをすべて関連付ける
            ):
                time.sleep(0.5)
                div_checkbox.click()
            time.sleep(1)
            input.clear()
            input.send_keys(案件ID)
            time.sleep(1)
            button.click()
            time.sleep(2)
            # Check Result
            snackbar_div = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div[class='snackbar__msg']")
                )
            )
            if snackbar_div.text == "案件との関連付けを行いました。":
                self.logger.info(f"Liên kết {案件ID}: {snackbar_div.text}")
                return True, snackbar_div.text
            else:
                self.logger.info(f"Liên kết {案件ID}: {snackbar_div.text}")
                return False, snackbar_div.text
        except TimeoutException:
            button = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "button[title='一括操作']")
                )
            )
            button = self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "button[title='一括操作']")
                )
            )
            button.click()
            return self.一括操作(
                案件ID=案件ID,
                このメールと同じ親番号のメールをすべて関連付ける=このメールと同じ親番号のメールをすべて関連付ける,
            )
        except NoSuchElementException as e:
            self.logger.error(
                f"❌ Liên kết {案件ID} thất bại: {e.msg.split('(Session info')[0].strip()}"
            )
            return False, e
        except Exception as e:
            self.logger.error(
                f"❌ Liên kết {案件ID} thất bại: {e.msg.split('(Session info')[0].strip()}"
            )
            return False, e

    @retry(exceptions=())
    @require_authentication
    @switch_to_default_content
    def send_mail(
        self,
        from_email: str,
        to_email: str,
        content: str,
        attachments: str | list[str] | None = None,
    ) -> bool:
        self.logger.info(f"Send Mail: From {from_email} to {to_email}")
        self.wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.CSS_SELECTOR, "iframe[id='ifmSide']"),
            )
        )
        self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[title='メール作成']"))
        ).click()
        while len(self.browser.window_handles) == 1:
            continue
        work_window = self.browser.window_handles[-1]
        self.browser.switch_to.window(work_window)
        self.browser.maximize_window()
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='次へ']"))
        ).click()
        fromInput = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='fFrom']"))
        )
        toInput = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='fTo[]']"))
        )
        fromInput.clear()
        toInput.clear()
        fromInput.send_keys(from_email)
        toInput.send_keys(to_email)
        self.wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
        )
        body = self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        self.browser.execute_script(
            "arguments[0].innerText = arguments[1];", body, content
        )
        self.browser.switch_to.default_content()
        if attachments:
            if isinstance(attachments, str):
                attachments = [attachments]
            # ------#
            for attach in attachments:
                self.wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "button[title='添付ファイル']")
                    )
                ).click()
                OpenDialog: WindowSpecification = wait_until_passes(
                    timeout=10,
                    retry_interval=0.5,
                    func=lambda: Application()
                    .connect(title="Open")
                    .window(title="Open"),
                )
                if not os.path.exists(attach):
                    self.logger.error(f"{attach} không tồn tại")
                    continue
                filename: UIAWrapper = wait_until_passes(
                    timeout=self.timeout,
                    retry_interval=1,
                    func=lambda: OpenDialog.child_window(class_name="ComboBoxEx32")
                    .child_window(class_name="Edit")
                    .wrapper_object(),
                )
                filename.type_keys(re.sub(r"([{}^%~()])", r"{\1}", str(Path(attach))))
                open_button: UIAWrapper = wait_until_passes(
                    timeout=self.timeout,
                    retry_interval=1,
                    func=lambda: OpenDialog.child_window(
                        title="&Open", class_name="Button"
                    ).wrapper_object(),
                )
                open_button.click_input()
                self.logger.info(f"Đính kèm file: {attach}")
                time.sleep(1)
        self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='menu']"))
        )
        while True:
            menu = self.browser.find_element(By.CSS_SELECTOR, "div[class='menu']")
            if len(menu.find_elements(By.XPATH, "./*")) == 2:
                break
            else:
                menu.click()
                time.sleep(1)
        menu = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='menu']"))
        )
        for button in menu.find_elements(By.TAG_NAME, "button"):
            if button.text == "一時保存":
                self.wait.until(EC.element_to_be_clickable(button)).click()
                break
        time.sleep(5)
        self.browser.close()


__all__ = [MailDealer]
