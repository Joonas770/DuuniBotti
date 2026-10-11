from robocorp.tasks import task
from robocorp import browser
from RPA.Excel.Application import Application
from playwright.sync_api import TimeoutError

from datetime import datetime, date, timedelta
import sys 



@task
def main():
    """Ohjelman pääfunktio"""
    sys.stdout.reconfigure(encoding="utf-8")

    browser.configure(
        slowmo=0,
    )

    # Määritellään Robotille hakusanat
    keywords = [
        "rpa-asiantuntija",
        "Ohjelmistokehittäjä",
        "Fullstack-kehittäjä",
        "Pilvipalveluasiantuntija"

    ]

    search_duunitori(keywords)




def search_duunitori(keywords):
    """Menee sivustolle ja syöttää hakusanat sekä muut suodatukset"""

    rows = []
    browser.goto("https://duunitori.fi/")

    page = browser.page()

    page.get_by_role("button", name="Hyväksy Evästeet").click()

    page.get_by_role("searchbox", name="Sijainti").fill("Uusimaa")

    for word in keywords:
        # Syötetään hakupalkkiin kaikki hakusanat
        page.get_by_role("searchbox", name="Etsi").fill(word)
        page.get_by_role("button", name="Etsi").click()


    # Järjestetään löydetyt työpaikkailmoitukset uusimmasta vanhimpaan
    page.get_by_role("link", name="Järjestä").click()
    page.get_by_role("button", name="Uusimmat").click()

    jobs = page.locator(".job-box.job-box--lg")

    today = date.today()
    links = []

    for job in jobs.all():
        link = job.locator("a.job-box__hover").get_attribute("href")
        published = job.locator(".job-box__job-posted").inner_text()
        published_text = published.split(" ")[1]

        if published_text.endswith("."):
            published_text += str(today.year)

        published = datetime.strptime(
            published_text,
        "%d.%m.%Y"
        ).date()
        if published >= today - timedelta(days=1):
            links.append(link)
    
    for link in links:
        page.goto("https://duunitori.fi" + link)
        rows.append(get_data())

    add_data_to_excel(rows)

def get_data():
    page = browser.page()
    title = page.locator("h1.text--break-word").inner_text()
    try:
        company = page.locator('a[href*="/tyopaikat/yritys"]').first.inner_text(timeout=5000)
    except TimeoutError:
        company = "ei tietoa"

    # Määrittelyt sijainnit työpaikoille
    locationsConfig = [
        "Helsinki",
        "Vantaa",
        "Espoo"
    ]

    locations = page.locator(
        'p.header__info a[href*="/tyopaikat/alue/"]'
    ).all_inner_texts()

    company_location = [
        location.strip().removeprefix("-").strip()
        for location in locations
        if location.strip().removeprefix("-").strip() in locationsConfig
    ]

    return [title, company, ", ".join(company_location), page.url]

    
def add_data_to_excel(rows):
    app = Application()

    app.open_application()
    try:
        app.open_workbook("output/template/template.xlsx")
        today = date.today().strftime(
            "%d.%m.%Y"
        )
        if rows:
            app.write_data_to_range(
                f"A2:D{len(rows) + 1}",
                rows
        )
        app.save_excel_as(f"output/history/{today}_jobs.xlsx")
    finally:
        app.quit_application()
