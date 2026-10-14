from robocorp.tasks import task
from robocorp import browser
from robocorp import vault
from RPA.Excel.Application import Application
from RPA.Email.ImapSmtp import ImapSmtp
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
        "DevOps-insinööri",
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

    # Poistetaan kaksoiskappaleet listasta
    unique_rows = [list(x) for x in set(tuple(row) for row in rows)]

    #Suodatetaan ja haetaan top 5 työpaikat
    top_5_jobs = filter_and_get_top_5(unique_rows)

    #Jos työpaikkoja löytyy, lähetetään yhteenveto sähköpostitse
    if top_5_jobs:
        send_email_summary(top_5_jobs)

    #Päivitetään rows muuttuja Excel tallennusta varten
    rows = unique_rows

    add_data_to_excel(rows)

def get_data():
    """Hakee työpaikkojen tiedot ja palauttaa ne listana"""
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

def filter_and_get_top_5(rows):
    scored_jobs = []
    bonus_keywords = ["Python", "RPA", "etätyö", "Fullstack", "Java"]

    for row in rows:
        # row[0] = title, row[1] = company, row[2] = location, row[3] = url
        title_lower = row[0].lower()
        score = 0

        # Suodatus: Poistetaan Senior ja lead työpaikat
        if "senior" in title_lower or "lead" in title_lower:
            continue

        # Lasketaan bonuspisteet hakusanoista
        for word in bonus_keywords:
            if word.lower() in title_lower:
                score += 1

        # Lisätään pisteet listaan
        scored_jobs.append((score, row))

    # Järjestetään työpaikat pisteiden mukaan laskevasti
    scored_jobs.sort(key=lambda x: x[0], reverse=True)

    # Palautetaan top 5 työpaikat
    return [job[1] for job in scored_jobs[:5]]

# Kirjaudutaan sähköpostipalvelimelle
def send_email_summary(top_5_jobs):
    mail_credentials = vault.get_secret("EmailCredentials")
    mail = ImapSmtp(smtp_server="smtp-mail.outlook.com", smtp_port=587)
    mail.authorize_smtp(
        account=mail_credentials["username"],
        password=mail_credentials["password"],  
    )

    # Luodaan sähköpostin sisältö
    body = "<h2>Päivän Top 5 kiinnostavimmat työpaikat</h2><ul>"
    for i, job in enumerate(top_5_jobs, start=1):
        body += f"<li><b>{i}. {job['title']}</b> - {job['company']} ({job['location']})<br>"
        body += f"<a href='{job['url']}'>Katso ilmoitus tästä</a></li><br>"
    body += "</ul>"

    # Sähköpostin lähetys
    mail.send_message(
        sender=mail_credentials["username"],
        recipients=mail_credentials["recipient"],
        subject=f"Uudet Top 5 työpaikat - {date.today().strftime('%d.%m.%Y')}",
        body=body,
        html=True
    )
