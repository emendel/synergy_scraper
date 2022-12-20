import auto_application_helpers
import synergy
from selenium.webdriver.common.by import By
import time
from decouple import config


def test_string_equal():
    assert 62 == 62


def test_open_browser():
    b = auto_application_helpers.init("https://google.com")
    elem = b.find_elements(
        by=By.XPATH, value="//*[contains(text(), 'Privacy')]")
    time.sleep(2)
    assert (len(elem) > 0)
    b.quit()
