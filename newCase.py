import asyncio
import os
import time
from datetime import datetime
from O365 import Account
# from mailSend import send_email_with_attachment
import csv

class NewCase:
    def __init__(self, page, df1, df2):
        self.page = page
        self.df1 = df1
        self.df2 = df2

    def fetch_value(self, df, key):
        return df[df['KEY'] == key]['VALUE'].values[0]



    def get_value(self, column_name):
        try:
            # Fetch value from the specific column and convert to string
            return str(self.df2[column_name].values[0])
        except KeyError:
            print(f"KeyError: '{column_name}' column not found in DataFrame.")
            return None
        except IndexError:
            print(f"IndexError: No data found in column '{column_name}'.")
            return None
        
    async def extract_and_save_benefits(self):
        try:
            modal = await self.page.wait_for_selector('div.MuiDialog-container div.MuiPaper-root', state='visible', timeout=120000)
            await asyncio.sleep(5)  # Ensure all data is loaded
            rows = await modal.query_selector_all('table tbody tr')
            print(f"Debug: Found {len(rows)} rows in the table.")

            benefits = {}
            for index, row in enumerate(rows):
                # Adapt the script to target 'th' elements
                cells = await row.query_selector_all('th')
                print(f"Debug: Row {index} has {len(cells)} cells containing data.")
                if len(cells) >= 2:
                    key = await cells[0].evaluate('(element) => element.textContent')
                    value = await cells[1].evaluate('(element) => element.textContent')
                    benefits[key.strip()] = value.strip()
                    print(f"Extracted: {key.strip()} -> {value.strip()}")

            if benefits:
                with open('D:\\AlgoSpring\\python\\MaxHealth\\benefits_list.csv', 'w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(['Benefit', 'Coverage'])
                    for key, value in benefits.items():
                        writer.writerow([key, value])
                print("Benefits saved to CSV file.")
                await self.page.get_by_role("button").click()
                
            else:
                print("No benefits data extracted.")

        except Exception as e:
            print(f"Failed to extract and save benefits due to error: {e}")

    async def extract_and_save_benefitsB(self):
        try:
            modal = await self.page.wait_for_selector('div.MuiDialog-container div.MuiPaper-root', state='visible', timeout=120000)
            await asyncio.sleep(5) 
            rows = await modal.query_selector_all('table tbody tr')
            print(f"Debug: Found {len(rows)} rows in the table.")

            benefits = {}
            for index, row in enumerate(rows):
                cells = await row.query_selector_all('th')
                print(f"Debug: Row {index} has {len(cells)} cells containing data.")
                if len(cells) >= 2:
                    key = await cells[0].evaluate('(element) => element.textContent')
                    value = await cells[1].evaluate('(element) => element.textContent')
                    benefits[key.strip()] = value.strip()
                    print(f"Extracted: {key.strip()} -> {value.strip()}")

            if benefits:
                with open('D:\\AlgoSpring\\python\\MaxHealth\\benefits_list_B.csv', 'w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(['Benefit', 'Coverage'])
                    for key, value in benefits.items():
                        writer.writerow([key, value])
                print("Benefits saved to CSV file.")
                await self.page.get_by_role("button").click()
            else:
                print("No benefits data extracted.")

        except Exception as e:
            print(f"Failed to extract and save benefits due to error: {e}")

    async def extract_and_save_benefitsC(self):
        try:
            modal = await self.page.wait_for_selector('div.MuiDialog-container div.MuiPaper-root', state='visible', timeout=120000)
            await asyncio.sleep(5)
            rows = await modal.query_selector_all('table tbody tr')
            print(f"Debug: Found {len(rows)} rows in the table.")

            benefits = {}
            for index, row in enumerate(rows):
                cells = await row.query_selector_all('th')
                print(f"Debug: Row {index} has {len(cells)} cells containing data.")
                if len(cells) >= 2:
                    key = await cells[0].evaluate('(element) => element.textContent')
                    value = await cells[1].evaluate('(element) => element.textContent')
                    benefits[key.strip()] = value.strip()
                    print(f"Extracted: {key.strip()} -> {value.strip()}")

            if benefits:
                with open('D:\\AlgoSpring\\python\\MaxHealth\\benefits_list_C.csv', 'w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(['Benefit', 'Coverage'])
                    for key, value in benefits.items():
                        writer.writerow([key, value])
                print("Benefits saved to CSV file.")
                await self.page.get_by_role("button").click()
            else:
                print("No benefits data extracted.")

        except Exception as e:
            print(f"Failed to extract and save benefits due to error: {e}")

    async def create_new_case(self):
        screenshot_path = "D:\\AlgoSpring\\python\\MaxHealth"
        # Click on the 'New Case' button 
        await self.page.locator("text=New Case").click()
        await self.page.wait_for_load_state('networkidle')  

        # Uplaod Censuse file
        file_input_locator = self.page.locator("input[type='file'][accept*='.xlsx']")
        await file_input_locator.wait_for(state="attached", timeout=10000) 

        # Set the file
        await file_input_locator.set_input_files("D:\\AlgoSpring\\python\\MaxHealth\\MaxHealth.xlsx")
        
        await asyncio.sleep(0.5)

        # print(self.df2.columns)
        # print(self.df2.head())   
        # print(self.df1.columns) 
        # print(self.df1.head())   


        # Claim Handling By
        TPA = self.get_value("TPA")
        if not TPA:
            raise ValueError("TPA value is missing or invalid.")
        TPA = TPA.strip()
        print(f"Selected TPA value: {TPA}")

        # Selector for the dropdown
        dropdown_selector = '//*[@id="mui-39"]'  
        await self.page.wait_for_selector(dropdown_selector)
        TPA = TPA.strip().upper() 
        print("TPA ", TPA)


        # Step 1: Click the dropdown to open it
        await self.page.click(dropdown_selector)

        # Step 2: Wait for the dropdown options to appear
        option_selector = f"//li[contains(@class, 'MuiMenuItem-root') and contains(text(), '{TPA}')]"
        await self.page.wait_for_selector(option_selector)
        options = await self.page.locator("//li[contains(@class, 'MuiMenuItem-root')]").all_text_contents()
        print(f"Available options: {options}")
        await self.page.click(option_selector)



        # Product Line
        network = self.get_value("Network")
        network = network.strip().upper()
        product_line = network.split()[0] 
        await self.page.locator('//*[@id="mui-40"]').click()      
        await self.page.locator(f'li[role="option"]:has-text("{product_line}")').click()
        await asyncio.sleep(0.5) 


        # Name of Client
        company_name = self.fetch_value(self.df1, "Company Name")
        await self.page.locator("//input[@id='clientName']").fill(company_name)
        print(company_name)


        #client location 
        location = self.fetch_value(self.df1, "Emirates")
        print(location)
        await self.page.locator('//*[@id="mui-42"]').click()      
        await self.page.locator(f'li[role="option"]:has-text("{location}")').click()
        await asyncio.sleep(0.5) 

        # Premium Frequency
        premium_frequency = self.get_value("Premium Frequency")
        await self.page.locator('//*[@id="mui-44"]').click()    
        print(premium_frequency)  
        await self.page.locator(f'li[role="option"]:has-text("{premium_frequency}")').click()
        await asyncio.sleep(0.5) 

        # Account Handling Person Name  
        contact_person = self.fetch_value(self.df1, "Contatct Person")
        await self.page.locator("//input[@id='accountHandlingPersonName']").fill(contact_person)
        await asyncio.sleep(0.5) 
        # Fetch the policy start date
        policy_str_date = self.fetch_value(self.df1, "Effective from")
        print(policy_str_date)
        # Check if the value is already a datetime object
        if isinstance(policy_str_date, datetime):
            # Format directly if it's already a datetime object
            policy_str_date = policy_str_date.strftime("%d/%m/%Y")
        else:
            # Parse the string into a datetime object and reformat it
            policy_str_date = datetime.strptime(policy_str_date, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y")
            print(policy_str_date)

        # Use the reformatted date
        await self.page.locator("input[placeholder='dd/mm/yyyy']").click()
        await self.page.locator("input[placeholder='dd/mm/yyyy']").fill(policy_str_date)
        await asyncio.sleep(0.5)

        # Account Handling Person Email 
        email = self.fetch_value(self.df1, "Email")
        await self.page.locator('input[name="brokerSecondaryEmail"]').fill(email)
        await asyncio.sleep(0.5) 

        # Current Insurer-not available(equal to Contact person value): mandotory field
        current_insurer = self.fetch_value(self.df1, "Contatct Person")
        await self.page.locator('input[name="currentInsurer"]').fill(current_insurer)
        print("current insurer")
        await asyncio.sleep(0.5) 

        # Account Handling Person Mobile Number 
        contact_number = str(self.fetch_value(self.df1, "Contact  Number"))
        await self.page.locator('input[name="brokerMobile"]').fill(contact_number)
        await asyncio.sleep(0.5) 

        # Policy Holder Type (mui-50)
        policy_holder_type = "SME Group: 10 members (minimum of 3 employees) up to 300 members"
        await self.page.locator('//*[@id="mui-49"]').click()
        await self.page.locator(f'li[role="option"]:has-text("{policy_holder_type}")').click()
        await asyncio.sleep(0.5)

        # Target premium
        target_premium = self.get_value("Target Premium")  
        await self.page.locator('//*[@id="mui-51"]').click()
        print(target_premium)
        await self.page.locator('//*[@id="mui-51"]').type(target_premium)
        await asyncio.sleep(0.5) 

        # Quotation For
        quotation_for = self.fetch_value(self.df1, "New-Renew")
        # Check New-Renew (New/ New - Virgin)
        # New - New client - Need to upload TOP
        # New-Virgin - New Client(virigin Group) 
        if quotation_for == "New":
            await self.page.get_by_label("Quotation For").click()
            await self.page.get_by_role("option", name="New Client", exact=True).click()
            await asyncio.sleep(0.5)
            print("ok1")
            
            # Upload TOP
            try:
                # TOB file path
                file_path = "D:\\AlgoSpring\\python\\MaxHealth\\CS-170524-B65751-V1.pdf"
                # Provide the XPath to the input element where files are uploaded
                input_xpath = '//*[@id="root"]/div/main/div/div/div[2]/div[1]/form/div/div/div[2]/div[2]/div[16]/div/div/input'
                # Set the input element's visibility to true before uploading (since it's hidden)
                await self.page.eval_on_selector(input_xpath, "input => input.style.display = 'block'")
                # Upload the file
                await self.page.set_input_files(input_xpath, file_path)

            except Exception as e:
                print(f"File upload failed: {e}")

            await asyncio.sleep(0.5) 
            print("ok2")

        else:
            await self.page.get_by_label("Quotation For").click()
            await self.page.get_by_role("option", name="New Client (Virgin Group)").click()
            await asyncio.sleep(0.5)
            print("ok1")        

        # Click Next Button
        await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[3]/button[2]').click()

        # Click Next Button
        await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[3]/button[3]/span').click()

        await asyncio.sleep(0.5)


        #----------------ASSIGN PLANS-------------------------------

        # Dynamically process categories and their plans
        categories_present = self.df2['Category'].dropna().unique()  # Get unique categories (A, B, C)

        for category in categories_present:
            category = category.strip().upper()  
            category_data = self.df2[self.df2['Category'] == category]  

            if category_data.empty:
                print(f"No data found for category {category}. Skipping.")
                continue

            select_plan = category_data['Select Plan'].values[0]  

            print(f"Processing category {category} with dynamic plan: {select_plan}")

            # Each category Benefits
            if category == 'A':
                await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[1]/div/div[1]/div[2]/div[2]/div/div').click()
                await asyncio.sleep(0.5) 
                await self.page.locator(f'li[role="option"]:has-text("{select_plan}")').click()
                await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[1]/div/div[1]/div[2]/div[3]/button').click()
                print("click1")
                await self.extract_and_save_benefits()
                print("click2")
                await asyncio.sleep(5) 

            if category == 'B':
                await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[1]/div/div[1]/div[3]/div[2]/div/div').click()
                await asyncio.sleep(0.5)
                await self.page.locator(f'li[role="option"]:has-text("{select_plan}")').click()
                await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[1]/div/div[1]/div[3]/div[3]/button').click()
                await self.extract_and_save_benefitsB()
                await asyncio.sleep(0.5)
                
            if category == 'C':
                await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[1]/div/div[1]/div[4]/div[2]/div/div').click()
                await asyncio.sleep(0.5)
                await self.page.locator(f'li[role="option"]:has-text("{select_plan}")').click()
                await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[1]/div/div[1]/div[4]/div[3]/button').click()
                await self.extract_and_save_benefitsC()
                await asyncio.sleep(0.5)

            await asyncio.sleep(0.5)



        # Click Next Button
        await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[3]/button[3]').click()

        # Save and Exit Button
        await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[3]/button[2]').click()
        await asyncio.sleep(5)



#       ---------------------------------------Download Quatation--------------------------------------
############################### Haven't enough data for the successful download ######################

        # # Extract the value "CS-201124-B73859-V1" from the page
        # element = self.page.locator('//*[@id="root"]/div/main/div/div/div[1]/div[1]/div[1]/div/div/span[1]')
        # extracted_value = await element.text_content()
        # print("Extracted Value:", extracted_value)


        # # Take a screenshot
        # screenshot_path = f"D:\\AlgoSpring\\python\\MaxHealth\\{extracted_value}.png"
        # await self.page.screenshot(path=screenshot_path)
        # print(f"Screenshot saved at {screenshot_path}")

        # await asyncio.sleep(5)

        # # Send email with the screenshot as attachment
        # email_sent = send_email_with_attachment("sudeepa.w@thealtria.com", screenshot_path, extracted_value)
        # if email_sent:
        #     print("Email with screenshot sent successfully.")
        # else:
        #     print("Failed to send email with screenshot.")


        # # Click the "SEND FOR REVIEW" button
        # button = await self.page.wait_for_selector('button:has-text("SEND FOR REVIEW")', timeout=5000)
        # await button.click()
        # print("Button clicked successfully.")

        

        # # print("download")
        # # download
        # download_path = 'D:\\AlgoSpring\\python\\MaxHealth'
        
        # # Expect download and handle the timeout
        # try:
        #     async with self.page.expect_download(timeout=120000) as download_info:
        #         # Ensure the right button triggers the download
        #         await self.page.locator('//*[@id="root"]/div/main/div/div/div[2]/div[3]/button[2]').click()

        #         # Save the downloaded file to the specified path
        #         download = await download_info.value
        #         download_path = os.path.join(download_path, "quotation1.pdf")
        #         await download.save_as(download_path)
        #         print(f"Downloaded quotation PDF to: {download_path}")

        # except asyncio.TimeoutError:
        #     print("Error: Timeout exceeded while waiting for the download.")
        # except Exception as e:
        #     print(f"Unexpected error during download: {e}")

        # await asyncio.sleep(10)  # Wait a little for final processing