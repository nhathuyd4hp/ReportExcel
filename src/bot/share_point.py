import os
import re
import time
import logging
from selenium import webdriver
from src.common.decorator import retry
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from src.common.decorator import require_authentication
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException


class SharePoint:
    def __init__(
        self,
        url: str,
        username: str,
        password: str,
        timeout: int = 10,
        headless: bool = False,
        download_directory: str = os.path.dirname(os.path.abspath(__file__)),
        logger_name: str = __name__,
    ):
        os.makedirs(download_directory, exist_ok=True)
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-notifications")
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": download_directory,
                "download.prompt_for_download": False,
                "safebrowsing.enabled": True,
            },
        )
        # Disable log
        options.add_argument("--disable-logging")
        options.add_argument("--log-level=3")
        options.add_argument("--silent")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        if headless:
            options.add_argument("--headless=new")
        # Attribute
        self.url = url
        self.logger = logging.getLogger(logger_name)
        self.browser = webdriver.Chrome(options=options)
        self.browser.maximize_window()
        self.timeout = timeout
        self.wait = WebDriverWait(self.browser, timeout)
        self.username = username
        self.download_directory = download_directory
        self.password = password
        # Trạng thái đăng nhập
        self.authenticated = self.__authentication(username, password)
        # Root Window
        self.root_window = self.browser.window_handles[0]

    def __get_latest_downloaded_file(self):
        self.browser.execute_script("window.open('');")
        # Wait for open
        while len(self.browser.window_handles) == 1:
            continue
        time.sleep(1)
        new_window = next(
            w for w in self.browser.window_handles if w != self.root_window
        )
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
            file_name = self.browser.execute_script(f"""
                return document
                    .querySelector("downloads-manager")
                    .shadowRoot.querySelector("#downloadsList")
                    .querySelector("#list")
                    .querySelector("#{item.get_attribute("id")}")
                    .shadowRoot.querySelector("#details")
                    .querySelector("#name")
                    .textContent
                    """)
        self.browser.close()
        self.browser.switch_to.window(self.root_window)
        return file_name
    def __get_status_download(self, files: list[str]) -> list[tuple[str, str]]:
        statuses = []
        self.browser.execute_script("window.open('');")
        # Wait for open
        while len(self.browser.window_handles) == 1:
            continue
        time.sleep(1)
        new_window = next(
            w for w in self.browser.window_handles if w != self.root_window
        )
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
                """)
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
                """)
            if name in files:
                statuses.append((name,tag))
        self.browser.close()
        self.browser.switch_to.window(self.root_window)
        return statuses

    @retry(exceptions=(TimeoutException))
    def __authentication(self, username: str, password: str) -> bool:
        time.sleep(0.5)
        self.browser.get("https://login.microsoftonline.com/")
        # -- Wait load page
        while self.browser.execute_script("return document.readyState") != "complete":
            continue
        if self.browser.current_url.startswith("https://m365.cloud.microsoft/?auth="):
            self.logger.info("Xác thực thành công!")
            return True
        # -- Email, phone or Skype
        self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="email"]'))
        ).send_keys(username)
        # -- Next
        btn = self.wait.until(EC.presence_of_element_located((By.ID, "idSIButton9")))
        self.wait.until(EC.element_to_be_clickable(btn)).click()
        # -- Check usernameError
        try:
            usernameError = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'div[id="usernameError"]')
                )
            )
            self.logger.error(usernameError.text)
            return False
        except TimeoutException:
            pass
        # -- Password
        self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="password"]'))
        ).send_keys(password)
        # -- Sign in
        btn = self.wait.until(EC.presence_of_element_located((By.ID, "idSIButton9")))
        self.wait.until(EC.element_to_be_clickable(btn)).click()
        # -- Check stay signed in
        try:
            passwordError = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'div[id="passwordError"]')
                )
            )
            self.logger.error(passwordError.text)
            return False
        except TimeoutException:
            pass
        # -- Stay signed in?
        self.wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div[class='row text-title']")
            )
        )
        btn = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='idSIButton9']"))
        )
        self.wait.until(EC.element_to_be_clickable(btn)).click()
        time.sleep(1)

        while self.browser.execute_script("return document.readyState") != "complete":
            continue

        if not self.browser.current_url.startswith(
            "https://m365.cloud.microsoft/?auth="
        ):
            self.logger.info(" Xác thực thất bại!")
            return False

        self.logger.info("Xác thực thành công!")
        return True

    @retry(
        exceptions=(StaleElementReferenceException),
    )
    @require_authentication
    def get_link_file(self, site_url: str, file: str) -> str:
        self.logger.info(f"Lấy link {file}: {site_url}")

    @retry(
        exceptions=(StaleElementReferenceException),
    )
    @require_authentication
    def download_file(
        self, site_url: str, file_pattern: str
    ) -> list[tuple[bool, list]]:
        self.logger.info(f"Tải {file_pattern}: {site_url}")
        download_files = []
        time.sleep(0.5)
        if not self.browser.current_url == site_url:
            self.browser.get(site_url)
        while self.browser.execute_script("return document.readyState") != "complete":
            continue
        # Access Denied
        try:
            ms_error_header = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div#ms-error-header h1")
                )
            )
            if ms_error_header.text == "Access Denied":
                SignInWithTheAccount = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div#ms-error a"))
                )
                SignInWithTheAccount.click()
            else:
                return (False, None)
        except TimeoutException:
            pass
        # -- Folder --
        found_folder = False
        folders = file_pattern.split("/")[:-1]
        for folder in folders:
            # Lấy tất cả các dòng
            rows = []
            try:
                ms_DetailsList_contentWrapper = self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "div[class='ms-DetailsList-contentWrapper']")
                    )
                )
                rows = ms_DetailsList_contentWrapper.find_elements(
                    by=By.CSS_SELECTOR,
                    value="div[class^='ms-DetailsRow-fields fields-']",
                )
                for row in rows:
                    icon_gridcell = row.find_element(
                        By.CSS_SELECTOR,
                        "div[role='gridcell'][data-automationid='DetailsRowCell']",
                    )
                    if icon_gridcell.find_elements(By.TAG_NAME, "svg"):  # Folder
                        name_gridcell = row.find_element(
                            By.CSS_SELECTOR,
                            "div[role='gridcell'][data-automation-key^='displayNameColumn_']",
                        )
                        button = name_gridcell.find_element(By.TAG_NAME, "button")
                        if button.text == folder:
                            button.click()
                            found_folder = True
                            time.sleep(5)  # Có thể tối ưu ở đây
                            break
            except TimeoutException:
                rows = self.browser.find_elements(
                    By.CSS_SELECTOR, "div[id^='virtualized-list_'][id*='_page-0_']"
                )
                for row in rows:
                    icon_gridcell = row.find_element(
                        By.CSS_SELECTOR,
                        "div[role='gridcell'][data-automationid='field-DocIcon']",
                    )
                    name_gridcell = row.find_element(
                        By.CSS_SELECTOR,
                        "div[role='gridcell'][data-automationid='field-LinkFilename']",
                    )
                    if icon_gridcell.find_elements(By.TAG_NAME, "svg"):  # Folder
                        span = name_gridcell.find_element(By.TAG_NAME, "span")
                        if span.text == folder:
                            span.click()
                            found_folder = True
                            time.sleep(5)  # Có thể tối ưu ở đây
                            break
        if not found_folder:
            self.logger.info(f"Không tìm thấy {file_pattern}")
            return (False, [])
        # -- File --
        pattern = file_pattern.split("/")[-1]
        pattern = re.compile(pattern)
        try:
            ms_DetailsList_contentWrapper = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div[class='ms-DetailsList-contentWrapper']")
                )
            )
            rows = ms_DetailsList_contentWrapper.find_elements(
                by=By.CSS_SELECTOR, value="div[class^='ms-DetailsRow-fields fields-']"
            )
            # Lấy tất cả các dòng
            if not rows:
                self.logger.info("Không tìm thấy file phù hợp")
                return (False, [])
            for row in rows:
                icon_gridcell = row.find_element(
                    By.CSS_SELECTOR,
                    "div[role='gridcell'][data-automationid='DetailsRowCell']",
                )
                if icon_gridcell.find_elements(By.TAG_NAME, "svg"):  # Folder
                    continue
                else:
                    name_gridcell = row.find_element(
                        By.CSS_SELECTOR,
                        "div[role='gridcell'][data-automation-key^='displayNameColumn_']",
                    )
                    button = name_gridcell.find_element(By.TAG_NAME, "button")
                    display_name = button.text
                    # Nếu display_name match với file là được
                    if pattern.match(display_name):
                        ActionChains(self.browser).context_click(button).perform()
                        if self.wait.until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "button[name='Download']"))
                        ):
                            time.sleep(1)
                            download_btn = self.browser.find_element(
                                By.CSS_SELECTOR, "button[name='Download']"
                            )
                            self.wait.until(
                                EC.element_to_be_clickable(download_btn)
                            ).click()
                            lastest_downloaded_file = (
                                self.__get_latest_downloaded_file()
                            )
                            self.logger.info(f"Tải {lastest_downloaded_file}")
                            download_files.append(lastest_downloaded_file)
        except TimeoutException:
            rows = self.browser.find_elements(
                By.CSS_SELECTOR, "div[id^='virtualized-list_'][id*='_page-0_']"
            )
            for row in rows:
                icon_gridcell = row.find_element(
                    By.CSS_SELECTOR,
                    "div[role='gridcell'][data-automationid='field-DocIcon']",
                )
                if icon_gridcell.find_elements(By.TAG_NAME, "svg"):  # Folder
                    continue
                else:
                    # Rewrite Here
                    name_gridcell = row.find_element(
                        By.CSS_SELECTOR,
                        "div[role='gridcell'][data-automationid='field-LinkFilename']",
                    )
                    span_name_gridcell = name_gridcell.find_element(By.TAG_NAME, "span")
                    if pattern.match(span_name_gridcell.text):
                        if row.find_elements(
                            By.CSS_SELECTOR, "div[class^='rowSelectionCell_']"
                        ):
                            time.sleep(1)
                            rowSelectionCell_ = row.find_element(
                                By=By.CSS_SELECTOR,
                                value="div[class^='rowSelectionCell_']",
                            )
                            self.wait.until(
                                EC.element_to_be_clickable(rowSelectionCell_)
                            ).click()
                            ActionChains(self.browser).context_click(
                                rowSelectionCell_
                            ).perform()
        time.sleep(5)
        if not download_files:
            self.logger.error(f"Không tìm thấy file nào khớp {file_pattern}")
        statuses = self.__get_status_download(download_files)
        for status in statuses:
            if status[1]:
                self.logger.info(f"{status[0]}: {status[1]}")
        return statuses


__all__ = [SharePoint]
