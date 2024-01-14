from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


class SeleniumManager:
    """Manages interactions with WhatsApp via Selenium."""

    # X PATH
    SEARCH_CSS = '#side > div._3gYev > div'
    SENDBUTTON_CSS = '#main > footer > div._2lSWV._3cjY2.copyable-area > div > span:nth-child(2) > div > div._1VZX7 > div._2xy_p._3XKXx'
    THREEDOTS_CSS = '#app > div > div.two._1jJ70 > div._2Ts6i._3RGKj > header > div._604FD > div > span > div:nth-child(6) > div > span'
    LOGOUTBUTTON_CSS = '#app > div > div.two._1jJ70 > div._2Ts6i._3RGKj > header > div._604FD > div > span > div._3OtEr._2Qn52 > span > div > ul > li:nth-child(6) > div'
    CONFIRMBUTTON_CSS = '#app > div > span:nth-child(3) > div > div > div > div > div > div > div.p357zi0d.ns59xd2u.kcgo1i74.gq7nj7y3.lnjlmjd6.przvwfww.mc6o24uu.e65innqk.le5p0ye3 > div > button.emrlamx0.aiput80m.h1a80dm5.sta02ykp.g0rxnol2.l7jjieqr.hnx8ox4h.f8jlpxt4.l1l4so3b.le5p0ye3.m2gb0jvt.rfxpxord.gwd8mfxi.mnh9o63b.qmy7ya1v.dcuuyf4k.swfxs4et.bgr8sfoe.a6r886iw.fx1ldmn8.orxa12fk.bkifpc9x.rpz5dbxo.bn27j4ou.oixtjehm.hjo1mxmu.snayiamo.szmswy5k > div > div'
    FINAL_CSS = '#app > div > div.landing-wrapper > div.landing-window > div.landing-main > div > div > div._3rDmx > div > span'

    def __init__(self):
        """Initializes the SeleniumManager with a new Chrome browser instance."""
        self.browser = webdriver.Chrome()

    def open_whatsapp(self):
        """Opens WhatsApp Web and waits for the user to log in."""

        # Get the screen size
        screen_width = self.browser.execute_script(
            "return window.screen.availWidth")
        screen_height = self.browser.execute_script(
            "return window.screen.availHeight")

        # Set the window size to half the screen width and full screen height
        self.browser.set_window_size(screen_width // 2, screen_height)

        # Set the window position to the right side of the screen
        self.browser.set_window_position(screen_width // 2, 0)

        self.browser.get("https://web.whatsapp.com")
        WebDriverWait(self.browser, 600).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, self.SEARCH_CSS)))
        time.sleep(1)

    def send_message(self, message: str, phoneNumber: str, country_code: str = '+91'):
        """
        Sends a message to the specified phone number.

        Args:
            message (str): The message to send.
            phoneNumber (str): The phone number to send the message to.
            country_code (str): The country code for the phone number. Defaults to '+91'.

        Returns:
            None
        """
        walink = f"https://web.whatsapp.com/send?phone={country_code}{phoneNumber}&text={message}&app_absent=1"
        self.browser.get(walink)
        try:
            sendButton = WebDriverWait(self.browser, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, self.SENDBUTTON_CSS)))
            time.sleep(1)
            sendButton.click()
            time.sleep(4)
        except Exception as error:
            print(phoneNumber, "error")
            print('Error while sending message', error, '\n')

    def logout(self):
        """ Logout from the previous logged in account """
        menu_button = WebDriverWait(self.browser, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, self.THREEDOTS_CSS)))
        time.sleep(1)
        menu_button.click()

        logout_button = WebDriverWait(self.browser, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, self.LOGOUTBUTTON_CSS)))
        time.sleep(1)
        logout_button.click()

        confirm_button = WebDriverWait(self.browser, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, self.CONFIRMBUTTON_CSS)))
        time.sleep(1)
        confirm_button.click()

        final = WebDriverWait(self.browser, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, self.FINAL_CSS)))
        time.sleep(1)

        self.browser.close()
        self.browser.quit()
