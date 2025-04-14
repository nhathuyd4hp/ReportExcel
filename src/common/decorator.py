import os
import uuid
import logging
import functools
from selenium.webdriver.chrome.webdriver import WebDriver


def require_authentication(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        instance = args[0]
        logger = getattr(instance, "logger", None)
        if getattr(instance, "authenticated", False):
            return func(*args, **kwargs)
        if logger:
            logger.error(f"{func.__name__} yêu cầu xác thực")
        return None

    return wrapper


def retry(exceptions=()):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            instance = args[0]
            logger = None
            browser = None
            if hasattr(instance, "logger") and isinstance(
                instance.logger, logging.Logger
            ):
                logger: logging.Logger = instance.logger
            if hasattr(instance, "browser") and isinstance(instance.browser, WebDriver):
                browser: WebDriver = instance.browser
            try:
                return func(*args, **kwargs)
            except exceptions:
                if logger:
                    logger.info(f"Retry: {func.__name__}")
                return func(*args, **kwargs)
            except Exception as e:
                saved_image = f"./screenshots/{uuid.uuid4()}.png"
                os.makedirs(os.path.dirname(saved_image), exist_ok=True)
                browser.save_screenshot(saved_image)
                if logger:
                    if hasattr(e, "msg"):
                        logger.error(f"{func.__name__}: {e.msg} - {saved_image}")
                    else:
                        logger.error(f"{func.__name__}: {e} - {saved_image}")
                return None

        return wrapper

    return decorator
