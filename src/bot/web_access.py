import os
import time
import logging
import pandas as pd
from datetime import date
from selenium import webdriver
from src.common.decorator import retry
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from src.common.decorator import require_authentication
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException


class WebAccess:
    def __init__(
        self,
        url:str,
        username: str,
        password: str,
        timeout: int = 10,
        headless: bool = False,
        download_directory: str = os.path.dirname(os.path.abspath(__file__)),
        logger_name: str = __name__,
    ):
        os.makedirs(download_directory, exist_ok=True)
        options = webdriver.ChromeOptions()
        options.add_argument("--disablenotifications")  # Tắt thông báo
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": download_directory,
                "download.prompt_for_download": False,
                "safebrowsing.enabled": True,
            },
        )

        if headless:
            options.add_argument("--headless=new")
        # Disable log
        options.add_argument("--disable-logging")
        options.add_argument("--log-level=3")  #
        options.add_argument("--silent")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        # Attribute
        self.url = url
        self.download_directory = download_directory
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

    def __get_latest_downloaded_file(self) -> str:
        self.browser.execute_script("window.open('');")
        # Wait for open
        while len(self.browser.window_handles) == 1:
            continue
        time.sleep(1)
        new_window = next(w for w in self.browser.window_handles if w != self.root_window)
        self.browser.switch_to.window(new_window)
        time.sleep(1)
        self.browser.get("chrome://downloads/")
        time.sleep(1)
        download_items: list[WebElement] = self.browser.execute_script("""
            return document.
                querySelector("downloads-manager").shadowRoot
                .querySelector("#mainContainer")
                .querySelector("#downloadsList")
                .querySelector("#list")
                .querySelectorAll("downloads-item")
        """)
        if download_items:
            item = download_items[0]
            downloaded_file: str =  self.browser.execute_script(f"""
                return document
                    .querySelector("downloads-manager")
                    .shadowRoot.querySelector("#downloadsList")
                    .querySelector("#list")
                    .querySelector("#{item.get_attribute("id")}")
                    .shadowRoot.querySelector("#details")
                    .querySelector("#name")
                    .textContent
                    """
            )
            self.browser.close()
            self.browser.switch_to.window(self.root_window)
            return downloaded_file
        return None
    
    def __get_status_download(self,file_name:str) -> tuple[str,str]:
        download_file = []
        self.browser.execute_script("window.open('');")
        # Wait for open
        while len(self.browser.window_handles) == 1:
            continue
        time.sleep(1)
        new_window = next(w for w in self.browser.window_handles if w != self.root_window)
        self.browser.switch_to.window(new_window)
        time.sleep(1)
        self.browser.get("chrome://downloads/")
        time.sleep(1)
        download_items: list[WebElement] = self.browser.execute_script("""
            return document.
            querySelector("downloads-manager").shadowRoot
            .querySelector("#mainContainer")
            .querySelector("#downloadsList")
            .querySelector("#list")
            .querySelectorAll("downloads-item")
        """)
        for item in download_items:
            name = self.browser.execute_script(f"""
                return document
                    .querySelector("downloads-manager").shadowRoot
                    .querySelector("#downloadsList")
                    .querySelector("#list")
                    .querySelector("#{item.get_attribute("id")}").shadowRoot
                    .querySelector("#content")
                    .querySelector("#details")
                    .querySelector("#title-area")
                    .querySelector("#name")
                    .getAttribute("title")
                """
            )
            tag = self.browser.execute_script(f"""
                return document
                    .querySelector("downloads-manager").shadowRoot
                    .querySelector("#downloadsList")
                    .querySelector("#list")
                    .querySelector("#{item.get_attribute("id")}").shadowRoot
                    .querySelector("#content")
                    .querySelector("#details")
                    .querySelector("#title-area")
                    .querySelector("#tag")
                    .textContent.trim();
                """
            )
            if name == file_name:
                return name,tag
        self.browser.close()
        self.browser.switch_to.window(self.root_window)
        return None,None


    @retry()
    def __authentication(self, username: str, password: str) -> bool:
        self.browser.get(self.url)
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input[type='text']")
            )
        ).send_keys(username)
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input[type='password']")
            )
        ).send_keys(password)
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "button[class='btn login']")
            )
        ).click()
        try:
            error_box = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR,"div[id='f-error-box']")
                )
            )
            data = error_box.find_element(By.CSS_SELECTOR,"div[class='data']")
            self.logger.info(f"Xác thực thất bại!: {data.text}")
            return False
        except TimeoutException:
            self.logger.info("Xác thực thành công!")
            return True

    def __switch_tab(self, tab: str) -> bool:
        a = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f"a[title='{tab}']"))
        )
        href = a.get_attribute("href")
        self.browser.get(href)
        return True
    
    @retry(
        exceptions=(ElementClickInterceptedException)
    )
    @require_authentication
    def get_information(
        self,
        builder_name: str = None,
        drawing: list[str] = None,
        delivery_date: list[str] = [date.today().strftime("%Y/%m/%d"), ""],
        fields: list[str] = None,
    ) -> pd.DataFrame | None:
        self.__switch_tab("受注一覧")
        # Clear
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "button[type='reset']")
            )
        )
        self.wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[type='reset']")
            )
        ).click()
        # Filter
        if builder_name:
            self.wait.until(
                EC.presence_of_element_located(
                    (By.ID, "select2-search_builder_cd-container")
                )
            )
            self.wait.until(
                EC.element_to_be_clickable(
                    (By.ID, "select2-search_builder_cd-container")
                )
            ).click()
            time.sleep(1)
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[class="select2-search__field"]')
                )
            ).send_keys(builder_name)
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[class="select2-search__field"]')
                )
            ).send_keys(Keys.ENTER)
            time.sleep(1)
            builder_field = self.wait.until(
                EC.presence_of_element_located(
                    (By.ID, "select2-search_builder_cd-container")
                )
            )
            builder_name = builder_field.text
        if delivery_date:  # Delivery date
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_from"]')
                )
            ).clear()
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_from"]')
                )
            ).send_keys(delivery_date[0])
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]')
                )
            ).clear()
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]')
                )
            ).send_keys(delivery_date[1])
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]')
                )
            ).send_keys(Keys.ESCAPE)
        if drawing:  # Drawing
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "button[id='search_drawing_type_ms']")
                )
            )
            self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "button[id='search_drawing_type_ms']")
                )
            ).click()
            for e in drawing:
                xpath = f"//span[text()='{e}']"
                self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                ).click()
            self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "button[id='search_drawing_type_ms']")
                )
            ).send_keys(Keys.ESCAPE)
        # Search
        time.sleep(1)
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "button[type='submit']")
            )
        )
        self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
        ).click()
        time.sleep(2)
        # Download File
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "a[class='button fa fa-download']")
            )
        )
        self.wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[class='button fa fa-download']")
            )
        ).click()
        time.sleep(5)
        file_name = self.__get_latest_downloaded_file()
        file_name,tag = self.__get_status_download(file_name)
        if tag:
            self.logger.info(f"Tải {file_name} thất bại: {tag}")
            return None
        self.logger.info(f"Tải {file_name} thành công")
        filePath = os.path.join(self.download_directory,file_name)
        df = pd.read_csv(filePath, encoding="CP932")
        df = df[fields] if fields else df
        os.remove(filePath)
        self.logger.info(f"Lấy thông tin '{builder_name}' thành công!")
        return df



__all__ = [WebAccess]