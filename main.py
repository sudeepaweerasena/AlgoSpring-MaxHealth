import asyncio
import pandas as pd
from playwright.async_api import async_playwright
from login import Login
from newCase import NewCase

async def main():
   
    file_path = "D:\\AlgoSpring\\python\\MaxHealth\\MaxHealth_Field Mapping.xlsx" 
    df1 = pd.read_excel(file_path, sheet_name="Sheet1")  # Key-Value
    df2 = pd.read_excel(file_path, sheet_name="Sheet2") 

    browser = None  # Define `browser` outside the try block
    try:
        async with async_playwright() as playwright:
            # Launch the browser
            browser = await playwright.chromium.launch(headless=False)
            context = await browser.new_context()
            context.set_default_timeout(60000)
            page = await context.new_page()

            # Login and process form
            login = Login(page)
            await login.perform_login("irfan.p@gargashinsurance.com", "November@123")
            
            new_case = NewCase(page, df1, df2)
            await new_case.create_new_case()

            # Wait for completion before closing the browser
            await asyncio.sleep(1)  # Wait for any pending operations

    finally:
        # Ensure the browser closes properly
        if browser:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
