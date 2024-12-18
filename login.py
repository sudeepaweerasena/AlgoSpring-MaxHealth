import undetected_playwright as playwright
import time


class Login:
    def __init__(self, page):
        self.page = page

    async def perform_login(self, username: str, password: str):
        await self.page.wait_for_load_state('networkidle')
        await self.page.goto("https://portal.maxhealth.ae/auth/login")
        await self.page.wait_for_load_state('networkidle')

        # Enter login credentials
        await self.page.locator("#mui-1").fill(username)
        await self.page.locator("#mui-2").fill(password)

        # Click login-in button
        await self.page.locator("#mui-3").click()
        

        time.sleep(5)
