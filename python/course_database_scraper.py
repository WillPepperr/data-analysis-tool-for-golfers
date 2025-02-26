from bs4 import BeautifulSoup
import requests
import pandas as pd
import random
import time
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

columns = ['Course'] + [str(i) for i in range(1, 10)] + ['Front'] + [str(i) for i in range(10, 19)] + ['Back', 'Total', 'State', 'City', 'Country']
df = pd.DataFrame(columns=columns)

letters = ['c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'y', 'z']
for letter in letters:
    letter_url = f"https://sourse.com.php?slet={letter}"
    page = requests.get(letter_url)
    soup = BeautifulSoup(page.text, 'html.parser')

    courses = soup.find_all('tbody', id="courselist")[0].find_all('tr')

    base_url =  "https://source.com/courses/"
    href_values = []

    for course in courses:
        link_tag = course.find('a')
        if link_tag and 'href' in link_tag.attrs:
            href_values.append(link_tag['href'])
    print(f"getting URL {letter_url}")
    time.sleep(5)
    count = 0
    for href in href_values:
        url = base_url + href
        count += 1
        try:
            page = requests.get(url, headers=headers)
            page.raise_for_status()
            soup = BeautifulSoup(page.text, 'html.parser')
            print(f"Scraping {url}")
            # Get Course Name
            html_name = soup.find_all('div', class_="headline moveup")[0]

            course_tag = html_name.find('h3')
            course_name = course_tag.get_text()
            course_name = course_name.replace('&nbsp', '').strip()
            print("Scraped Name")
            # Get Par Values

            front_nine = soup.find_all('table', class_='table table-bordered table-condensed centertext')[0]
            front_numbers = front_nine.find_all('td')

            front_values = [title.text for title in front_numbers]
            front_par_index = front_values.index('Par:')
            front_pars = front_values[front_par_index + 1: front_par_index + 11]
            print("Scraped Pars")
            is_back_nine = True
            is_course_data = False
            try:
                back_nine = soup.find_all('table', class_='table table-bordered table-condensed centertext')[1]
                back_numbers = back_nine.find_all('td')
                back_values = [title.text for title in back_numbers]
                back_par_index = back_values.index('Par:')
                back_pars = back_values[back_par_index + 1: back_par_index + 12]
            except:
                print("No back nine for this course")
                is_back_nine = False

            if is_back_nine is False:
                back_pars = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
                back_pars[-1] = front_pars[-1]

            html_name = soup.find_all('table', class_="table table-striped")[0]
            city = html_name.find_all('tr')[3]
            city_name = city.find_all('td')[1]
            city_name = city_name.get_text()

            state = html_name.find_all('tr')[4]
            state_name = state.find_all('td')[1]
            state_name = state_name.get_text()

            country = html_name.find_all('tr')[6]
            country_name = country.find_all('td')[1]
            country_name = country_name.get_text()

            new_row = [course_name] + front_pars + back_pars + [state_name, city_name, country_name]

            df.loc[len(df)] = new_row
            print(f"Successfully scrapped course:{count}/{len(href_values)}")

            wait_time = random.uniform(.5, 1)
            time.sleep(wait_time)

        except:
            print("No course data")
            wait_time = random.uniform(.5, 1)
            time.sleep(wait_time)


    df.to_csv(f'golf_courses_{letter}.csv', index=False)