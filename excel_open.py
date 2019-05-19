# -*- encoding=utf8 -*-
from automagica import *

def cinii_search(excel, search):
    print(excel, search)

    def get_next_count(browser)->int:
        header = browser.find_element_by_class_name('heading')
        texts = header.text.split()
        hit_num = int(texts[1][:-2])
        per_page = int(texts[2].split('-')[1])
        next_count = hit_num // per_page
        return next_count
    
    browser = ChromeBrowser()
    url = 'https://ci.nii.ac.jp'
    browser.get(url)
    browser.find_element_by_name('q').send_keys(search)
    button = browser.find_element_by_xpath('//*[@id="article_form"]/div/div[1]/div[2]/div[2]/button')
    button.click()
    next_click_count = get_next_count(browser)
    
    page = 0
    done = 0

    ExcelCreateWorkbook(path=excel)

    while page <= next_click_count:

        articles = browser.find_elements_by_class_name('paper_class')

        for article in articles:

            done += 1

            a = article.find_element_by_tag_name('a')
            title = a.text
            link = a.get_attribute('href')
            authors = article.text.split('\n')[1]
            jounal = article.find_element_by_class_name('journal_title').text
            
            # print(title, authors, jounal, link)
            ExcelWriteRowCol(excel, r=done, c=1, write_value=done)
            ExcelWriteRowCol(excel, r=done, c=2, write_value=title)
            ExcelWriteRowCol(excel, r=done, c=3, write_value=authors)
            ExcelWriteRowCol(excel, r=done, c=4, write_value=link)
            ExcelWriteRowCol(excel, r=done, c=5, write_value=jounal)

        if page != next_click_count:
            browser.find_element_by_xpath('//*[@class="paging_next btn pagingbtn"]').click()
            page += 1
        else: 
            print('正常に終了しました')
            break


if __name__ == '__main__':
    cinii_search(
        'output.xlsx', 
        '伊集院　利明' 
    )