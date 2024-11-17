from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd

URL = "https://steamdb.info/"

# Keep Chrome browser open after program finishes
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
driver.get(URL)

most_played_games = driver.find_element(By.XPATH, value='/html/body/div[4]/div[1]/div[2]/div[1]/div[1]/table')
trending_games = driver.find_element(By.XPATH, value='/html/body/div[4]/div[1]/div[2]/div[1]/div[2]/table')
popular_releases = driver.find_element(By.XPATH, value='/html/body/div[4]/div[1]/div[2]/div[2]/div[1]/table')
hot_releases = driver.find_element(By.XPATH, value='/html/body/div[4]/div[1]/div[2]/div[2]/div[2]/table')

tables = [most_played_games, trending_games, popular_releases, hot_releases]
most_played_dict = {}
trending_dict = {}
popular_releases_dict = {}
hot_releases_dict = {}

for table in tables:
    title = table.find_element(By.CLASS_NAME, value='table-title').text
    columns_title = table.find_elements(By.CLASS_NAME, value='text-center')

    games_list = []
    column_1_list = []
    column_2_list = []

    if title == 'Most Played Games':
        games = table.find_elements(By.CLASS_NAME, value='css-truncate')
        column_1 = table.find_elements(By.CLASS_NAME, value='tabular-nums')
        column_2 = table.find_elements(By.CLASS_NAME, value='tabular-nums')

        for item in games:
            games_list.append(item.text)

        for item in column_1:
            column_1_list.append(item.text)

        for item in column_2:
            column_2_list.append(item.text)

        most_played_dict[title] = games_list
        most_played_dict[columns_title[0].text] = column_1_list[0:30:2]
        most_played_dict[columns_title[1].text] = column_2_list[1:30:2]

    elif title == 'Trending Games':
        games = table.find_elements(By.CLASS_NAME, value='css-truncate')
        column_1 = []
        column_2 = table.find_elements(By.CLASS_NAME, value='tabular-nums')

        for item in games:
            games_list.append(item.text)

        for item in column_2:
            column_2_list.append(item.text)

        trending_dict[title] = games_list[:-2]
        trending_dict[columns_title[1].text] = column_2_list

    elif title == 'Popular Releases':
        games = table.find_elements(By.CLASS_NAME, value='css-truncate')
        column_1 = table.find_elements(By.CLASS_NAME, value='text-center')
        column_2 = table.find_elements(By.CLASS_NAME, value='text-center')

        for item in games:
            games_list.append(item.text)

        for item in column_1:
            column_1_list.append(item.text)

        for item in column_2:
            column_2_list.append(item.text)

        popular_releases_dict[title] = games_list
        popular_releases_dict[columns_title[0].text] = column_1_list[2:32:2]
        popular_releases_dict[columns_title[1].text] = column_2_list[3:32:2]

    elif title == 'Hot Releases':
        games = table.find_elements(By.CLASS_NAME, value='css-truncate')
        column_1 = table.find_elements(By.CLASS_NAME, value='text-center')
        column_2 = table.find_elements(By.CLASS_NAME, value='text-center')

        for item in games:
            games_list.append(item.text)

        for item in column_1:
            column_1_list.append(item.text)

        for item in column_2:
            column_2_list.append(item.text)

        hot_releases_dict[title] = games_list
        hot_releases_dict[columns_title[0].text] = column_1_list[2:32:2]
        hot_releases_dict[columns_title[1].text] = column_2_list[3:32:2]

print(most_played_dict)
print(trending_dict)
print(popular_releases_dict)
print(hot_releases_dict)

most_played_table = pd.DataFrame(most_played_dict)
trending_table = pd.DataFrame(trending_dict)
popular_releases_table = pd.DataFrame(popular_releases_dict)
hot_releases_table = pd.DataFrame(hot_releases_dict)

with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    # To write each DataFrame to a different sheet, must use pd.ExcelWriter
    # Alone .to_excel() works for only 1 sheet/ 1 file.
    most_played_table.to_excel(writer, sheet_name='Most Played Games', index=False)
    trending_table.to_excel(writer, sheet_name='Trending Games', index=False)
    popular_releases_table.to_excel(writer, sheet_name='Popular Released Games', index=False)
    hot_releases_table.to_excel(writer, sheet_name='Hot Released Games', index=False)

print("Dictionaries saved to different sheets in output.xlsx")

