from robocorp.tasks import task
from robocorp import browser
from robocorp import vault
from RPA.Excel.Application import Application
from RPA.Email.ImapSmtp import ImapSmtp
from playwright.sync_api import TimeoutError

from datetime import datetime, date, timedelta
from html import escape
import sys
import traceback

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

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
    try:
        tyomarkkintori_jobs = search_tyomarkkinatori(keywords)
        duunitori_jobs = search_duunitori(keywords)

        all_jobs = duunitori_jobs + tyomarkkintori_jobs
        if all_jobs:
            add_data_to_excel(all_jobs)

        #Suodatetaan ja haetaan top 5 työpaikat
        top_5_jobs = filter_and_get_top_5(all_jobs)

        if top_5_jobs:
            send_email_summary(top_5_jobs)
        else:
            print("Uusia työpaikkailmoituksia ei löytynyt.")
            send_notification_email(
                f"Ei uusia työpaikkoja - {date.today().strftime('%d.%m.%Y')}",
                "<p>Uusia työpaikkailmoituksia ei löytynyt.</p>"
            )
    except Exception as e:
        error_msg = f"Työpaikkahaun suoritus epäonnistui: {e}\n\n{traceback.format_exc()}"
        print(error_msg)
        error_body = escape(error_msg).replace("\n", "<br>")
        send_notification_email("Työpaikkahaun suoritus epäonnistui", f"<p>{error_body}</p>")


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
        rows.append(get_duunitori_data())

    # Poistetaan kaksoiskappaleet listasta
    unique_rows = [list(x) for x in set(tuple(row) for row in rows)]


    #Päivitetään rows muuttuja Excel tallennusta varten
    rows = unique_rows

    return rows

def search_tyomarkkinatori(keywords):
    page = browser.page()
    jobs = []

    page.goto("https://tyomarkkinatori.fi/henkiloasiakkaat/avoimet-tyopaikat?pa=MAX_1_DAY&r=01")
    page.get_by_role("button", name="Salli kaikki").click()

    for word in keywords:
        page.get_by_role("textbox").fill(word)
        page.get_by_role("button", name="Hae").click()
        page.wait_for_timeout(1500)

        jobs.extend(get_tyomarkkinatori_data())


        page.get_by_role("textbox").fill("")

    return jobs
def get_duunitori_data():
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

def get_tyomarkkinatori_data():
    rows = []
    page = browser.page()
    results = page.locator("#tmt-haku-search-results-scroll-anchor")

    jobs = results.locator(":scope > div").filter(has=page.locator("h3"))


    for job in jobs.all():
        try:
            title = job.locator("h3").inner_text()
            publisher = job.locator("ul > li").first.locator("span").inner_text()
            location_texts = job.locator(".lpkjRD span").all_inner_texts()
            location = location_texts[0] if location_texts else "ei tietoa"
            href = job.locator("h3 a").get_attribute("href")
            link = "https://tyomarkkinatori.fi" + href
            rows.append([title, publisher, location, link])
        except Exception as e:
            print(f"Virhe työpaikan käsitelyssä: {e}")
            continue
    return rows

def add_data_to_excel(rows):
    app = Application()

    app.open_application()
    try:
        app.open_workbook("output/template/template.xlsx")
        today = date.today().strftime("%d.%m.%Y")
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

def send_email_summary(top_5_jobs):
    # Luodaan sähköpostin sisältö
    body = "<h2>Päivän Top 5 kiinnostavimmat työpaikat</h2><ul>"
    for i, job in enumerate(top_5_jobs, start=1):
        body += f"<li><b>{i}. {job[0]}</b> - {job[1]} ({job[2]})<br>"
        body += f"<a href='{job[3]}'>Katso ilmoitus tästä</a></li><br>"
    body += "</ul>"

    send_notification_email(
        f"Uudet Top 5 työpaikat - {date.today().strftime('%d.%m.%Y')}",
        body
    )


def send_notification_email(subject, body):
    mail_credentials = vault.get_secret("EmailCredentials")
    mail = ImapSmtp(smtp_server=SMTP_SERVER, smtp_port=SMTP_PORT)
    mail.authorize_smtp(
        account=mail_credentials["username"],
        password=mail_credentials["password"],
    )

    # Sähköpostin lähetys
    mail.send_message(
        sender=mail_credentials["username"],
        recipients=mail_credentials["recipient"],
        subject=subject,
        body=body,
        html=True
    )
