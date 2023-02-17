import time
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
import xlwings as xw
URL = "https://careers.google.com/jobs/results/?page=1"
PAGE = 50


def main():
    filename = "Google_Job_Scrap_" + str(date.today()) + ".xlsx"
    print("Start scraping all google jobs....")
    print('Please, Do not close chrome browser. '
          'It will be closed automatically after finished.')
    print('This process maybe take 5 - 10 minutes')
    wb = xw.Book()
    sheet = wb.sheets[0]
    sheet.range("A1").value = ['Title', 'Company', 'Remote Eligible', 'Location', 'Update Time', 'Minimum qualifications',
                               'Preferred qualifications', 'Responsibilities', 'About Job', 'Link']
    chrome_options = webdriver.ChromeOptions()
    browser = webdriver.Chrome(options=chrome_options)
    scrape(URL, browser, sheet)
    wb.save(filename)


def scrape(ini_url, browser, sheet):
    count_pages = 1
    while count_pages <= PAGE:
        url_page = ini_url.replace("page=1", f"page={count_pages}")
        job_links = []
        get_job_link(url_page, browser, job_links, sheet)
        count_pages += 1


def get_job_link(url_page, browser, job_links, sheet):
    global total
    browser.get(url_page)
    time.sleep(2)
    total_pages = browser.find_element(By.XPATH, "//p[@class='gc-h-flex gc-sidebar__pagination--page']").text
    print(f"Progress: {total_pages}")
    print(".....Scraping.....")
    job_box = browser.find_elements(By.XPATH, '//ol[@id="search-results"]/li/div[@class="gc-card__container"]/a')
    for job in job_box:
        link = job.get_attribute("href")
        job_links.append(link)
    parse_jobs(browser, job_links, sheet)


def parse_jobs(browser, job_links, sheet):
    i = 0
    for link in job_links:
        browser.execute_script(f"window.open('{link}', 'new_window')")
        browser.switch_to.window(browser.window_handles[1])
        time.sleep(3)
        try:
            title = browser.find_element(By.XPATH,
                                     '''//h1[@class=
                                     "gc-card__title gc-job-detail__title gc-heading gc-heading--beta"]''').text
        except:
            title = " "
        try:
            company = browser.find_element(By.XPATH,
                                       '//li[@itemprop="hiringOrganization"]/span').text
        except:
            company = " "
        try:
            locations = browser.find_elements(By.XPATH,
                                          '''//div[@class="gc-card__header gc-job-detail__header"]/
                                          ul[@class="gc-job-tags gc-job-detail__tags"]/li[@itemprop="jobLocation"]/
                                          div[@itemprop="address"]''')
        except:
            locations = []
        try:
            browser.find_element(By.XPATH, '//ul[@class="gc-job-tags gc-job-detail__tags"]/'
                                           'li [@class="gc-job-tags__remote gc-job-tags--meta"]')
            remote = True
        except:
            remote = False
        try:
            min_q = browser.find_elements(By.XPATH,
                                      '//div[@itemprop="qualifications"]/ul')[0].text
        except:
            min_q = " "
        try:
            pre_q = browser.find_elements(By.XPATH,
                                      '//div[@itemprop="qualifications"]/ul')[1].text
        except:
            pre_q = " "
        try:
            about_job = browser.find_element(By.XPATH,
                                         '//div[@itemprop="description"]').text
        except:
            about_job = " "
        try:
            rsb = browser.find_element(By.XPATH,
                                   '//div[@itemprop="responsibilities"]').text
        except:
            rsb = " "
        try:
            update_time = browser.find_elements(By.XPATH, '//meta[@itemprop="datePosted"]')[i].get_attribute("content")
        except:
            update_time = "NA"
        i += 1
        loca = ""
        for location in locations:
            loca += location.text + "\n"
        row = [title, company, remote, loca, update_time, min_q, pre_q, rsb, about_job, link]
        if not sheet.range('A2').value:
            sheet.range('A2').value = row
        else:
            sheet.range("A" + str(sheet.range("A1").end('down').row+1)).value = row
        browser.close()
        browser.switch_to.window(browser.window_handles[0])


if __name__ == "__main__":
    main()
