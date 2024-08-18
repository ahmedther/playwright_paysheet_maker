import asyncio
import random
from env.env import URL, APP_USERNAME, PASSWORD
from playwright.async_api import async_playwright
from datetime import datetime, timedelta


class PlayWriterHandler:
    def __init__(self):
        self.browser = None

    async def launch_browser(self, headless=True):
        self.playwright = await async_playwright().start()
        # self.browser = await self.playwright.chromium.launch_persistent_context(
        #     user_data_dir=r"C:\Users\AHMED\AppData\Local\Google\Chrome\User Data\Default",
        #     headless=False,
        #     args=[
        #         "--disable-blink-features=AutomationControlled",
        #         "--no-sandbox",
        #         "--disable-setuid-sandbox",
        #         "--ignore-certificate-errors",
        #         "--disable-infobars",
        #     ],
        # )
        self.browser = await self.playwright.chromium.launch(
            headless=headless,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--ignore-certificate-errors",
                "--disable-infobars",
            ],
        )
        user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        ]
        self.page = await self.browser.new_page()
        await self.page.set_extra_http_headers(
            {
                "User-Agent": random.choice(user_agents),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9",
                "Referer": "https://www.google.com",
            }
        )
        # await stealth_async(self.page)
        await self.page.evaluate(
            "() => { Object.defineProperty(navigator, 'webdriver', { get: () => false }) }"
        )
        await self.page.goto(URL)

    async def login_and_go_to_schedule(self):
        await self.page.fill('input[name="auth_key"]', APP_USERNAME)
        await self.page.fill('input[name="password"]', PASSWORD)
        await self.page.click("button#log_in")
        await self.page.click('a[data-e2e-id="schedule-nav"]')

    async def change_page(self):
        await self.page.click('button[aria-label="Previous"]')

    async def extract_data(self, week_offset=0):
        data_list = []
        await self.page.wait_for_selector("ul.appointments li.color-none")
        full_week = await self.page.query_selector_all("ul.appointments")

        start_date = datetime.now() - timedelta(
            days=(datetime.now().weekday() + 1) % 7, weeks=week_offset
        )

        for i, appointments in enumerate(full_week):

            pt_data = await appointments.query_selector_all(
                'div.event-inner.calendar-x-event.align-left.appointment:not(.cancelled):not(.break) > div.inner-padding > div.details[data-testid="appointment-details"]'
            )
            for pt in pt_data:
                data = await pt.inner_text()
                date_of_session = (start_date + timedelta(days=i)).strftime("%#d-%b-%y")
                # date_of_session = (start_date + timedelta(days=i)).strftime("%Y-%m-%d")
                data_list.append([data, date_of_session])

        return data_list

    async def close_browser(self):
        await self.browser.close()
        await self.playwright.stop()

    async def run(self):
        try:
            await self.launch_browser()
            await self.login_and_go_to_schedule()
            data_week2 = await self.extract_data()
            await self.change_page()
            data_week1 = await self.extract_data(1)
            await self.close_browser()

            return data_week1, data_week2
            # self.save_data(data_week2, "data_week2.pkl")
            # self.save_data(data_week1, "data_week1.pkl")

            # Load data_week2 and data_week1 when needed
            # data_week2 = self.load_data("data_week2.pkl")
            # data_week1 = self.load_data("data_week1.pkl")

            # dataframe2 = self.convert_to_dataframe(data_week2)
            # dataframe1 = self.convert_to_dataframe(data_week1)
            # dataframe = self.concat_dataframes([dataframe1, dataframe2])
            # name = "formatted_output.xlsx"
            # self.save_to_excel(dataframe, name)
            # os.startfile(name)

            # input("Press Enter to close Excel and generate PDF with password...")

            # pdf_name = "formatted_output.pdf"
            # self.convert_excel_to_pdf(os.path.abspath(name), os.path.abspath(pdf_name))
            # self.set_pdf_password(pdf_name, "test")
            # os.startfile(pdf_name)

        except ResourceWarning:
            pass


if __name__ == "__main__":

    pw = PlayWriterHandler()
    data1, data2 = asyncio.run(pw.run())
