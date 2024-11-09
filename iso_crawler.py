from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from docx import Document
from css_to_word import apply_css_to_word

def ISO_Crawler(iso_url, output_path):
    # Set up Chrome Driver
    driver_path = './chromedriver.exe'

    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Create Service instance
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        doc = Document()
        driver.get(iso_url)
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "sts-standard"))
        )

        sections = driver.find_elements(By.CLASS_NAME, "sts-standard")
        for section in sections:

            process_element(doc, section)

        doc.save(output_path)
    finally:
        # Close the browser
        driver.quit()
        print(f"Crawl Complete! OutputFile: {output_path}")

def process_element(doc, element):
    # Get the text content of the element
    if 'sts-p' in element.get_attribute("class"):
        # For elements with 'sts-label', get the full text content (including children)
        text = element.text.strip()
    else:
        text = element.parent.execute_script(
            "return arguments[0].childNodes[0] ? arguments[0].childNodes[0].nodeValue : null", element
        )
    if text and text != '—':
        css_styles = {}
        css_styles['font-size'] = element.value_of_css_property("font-size")
        css_styles['font-weight'] = element.value_of_css_property("font-weight")
        css_styles['text-align'] = element.value_of_css_property("text-align")

        has_styles = any(css_styles.values())  # True if any style is not None or empty
        if has_styles:
            apply_css_to_word(doc, text, css_styles)

    children = element.find_elements(By.XPATH, "./*")
    for child in children:
        process_element(doc, child)
