import time
from auto_application_helpers import init
from selenium.webdriver.chrome.options import Options
# from fake_useragent import UserAgent
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import pandas as pd
import os
from user_agent import generate_user_agent, generate_navigator
from zipfile import ZipFile
import xlsxwriter
from decouple import config

USERNAME = config('USERNAME')
PASSWORD = config('PASSWORD')

team = "iowa"
year = "2022"
URL = "https://apps.synergysports.com/basketball/teams/54457ddd300969b132fd0615/overall?seasonId=630530d30a41b857ff3c1501"
dest_path = os.getcwd() + "/{}/{}".format(year, team)

if not os.path.exists(dest_path):
    os.makedirs(dest_path)


def init(link):
    options = Options()
    # ua = UserAgent()
    # user_agent = ua.random
    user_agent = generate_user_agent()
    options.add_argument(f'user-agent={user_agent}')

    browser = webdriver.Firefox()
    browser.get(link)
    time.sleep(2)
    return browser


def login():
    b = init(URL)
    time.sleep(2)
    b.maximize_window()
    username = b.find_element(by=By.ID, value="Username")
    username.click()
    action = ActionChains(b)
    action.send_keys(USERNAME).perform()
    time.sleep(0.5)
    b.find_element(by=By.ID, value="Password").click()
    action2 = ActionChains(b)
    action2.send_keys(PASSWORD, Keys.TAB, Keys.TAB, Keys.ENTER).perform()
    time.sleep(10)
    shot_types(b)
    play_types(b)
    cumulative_box(b)
    overall(b)
    zip_files()


def get_headers(b, team, single_line):
    headers = b.find_elements(by=By.TAG_NAME, value="th")
    for h in headers:
        single_line.append(h.text)
    team.append(single_line)


def retrieve_data(columns, b):
    single_line = []
    team = []
    get_headers(b, team, single_line)
    stats = b.find_elements(by=By.TAG_NAME, value="td")
    count = 0
    for stat in stats:
        count += 1
        res = stat.text
        res = res.replace(',', '')
        single_line.append(res)
        if count == columns:
            try:
                team.append(single_line)
            except:
                team.append(single_line)
            count = 0
            single_line = []
    return team


def shot_types(b):
    shot_types = b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'Shot Types')]")
    shot_types.click()
    time.sleep(5)
    b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'Player Breakdown')]").click()
    time.sleep(10)
    b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'At Rim')]").click()
    b.find_element(by=By.XPATH, value="//*[contains(text(), 'Hook')]").click()
    b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'Runner')]").click()
    b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'Jump Shot')]").click()
    b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'All Field Goal Attempts')]").click()

    time.sleep(5)
    write_data_play_type(retrieve_data(16, b), "shot_type", 16)


def cumulative_box(b):
    boxscore = b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'Cumulative Box')]")
    boxscore.click()
    time.sleep(5)
    write_data_cumulative_box(retrieve_data(43, b), "cumulative_box")


def write_data_cumulative_box(data, type):
    headers = data[0]
    del data[0]
    data[0] = data[0][-43:len(data[0])]
    headers = headers[0:43]

    df = pd.DataFrame(data, columns=headers)

    writer = pd.ExcelWriter(
        os.getcwd() + "/{}/{}/{}_{}.xlsx".format(year, team, type, team), engine='xlsxwriter')

    df = df.drop('CHG COM', axis=1)
    df = df.drop('CHG TKN', axis=1)
    df = df.drop('FG MADE', axis=1)
    df = df.drop('FG MISS', axis=1)
    df = df.drop('2 FG MADE', axis=1)
    df = df.drop('2 FG MISS', axis=1)
    df = df.drop('3 FG MADE', axis=1)
    df = df.drop('3 FG MISS', axis=1)
    df = df.drop('FT MADE', axis=1)
    df = df.drop('FT MISS', axis=1)
    df = df.drop('+1', axis=1)
    df = df.drop('SF', axis=1)
    df = df.drop('%SF', axis=1)
    df.to_excel(writer, sheet_name='whole_team', index=False)
    workbook = writer.book
    ws = writer.sheets['whole_team']
    for row in range(0, len(data)):
        for col in range(0, len(df.columns)):
            if '%' in df.iloc[row, col]:
                try:
                    df.iloc[row, col] = df.iloc[row, col].replace('%', '')
                    df.iloc[row, col] = float(df.iloc[row, col])
                except:
                    pass
            try:
                ws.write_number(
                    row+1, col, float(df.iloc[row, col]))
            except:
                ws.write(
                    row+1, col, df.iloc[row, col])
    highlighted_columns = ['E', 'I', 'N', 'P']
    for i in highlighted_columns:
        ws.conditional_format('{}2:{}{}'.format(
            i, i, len(data)), {'type': '3_color_scale'})

    writer.close()


def play_types(b):
    b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'Play Types')]").click()
    time.sleep(10)
    carets = b.find_elements(
        by=By.CLASS_NAME, value="la-caret-right")

    for c in reversed(carets):
        c.click()
    time.sleep(2)
    write_data_play_type(retrieve_data(25, b), "play_type", 25)


def write_data_play_type(data, type, num_columns):
    junk = len(data)
    new_data = []
    n = False
    for row in range(0, len(data)):
        if data[row][0] == 'Long 3 pts' and type == "shot_type":
            junk = row
        if "P&R Including Passes" in data[row] and type == "play_type":
            index = data[row].index('P&R Including Passes')
            junk = row
            n = True
        if n:
            for col in range(0, len(data[row])):
                if junk == row and col >= index:
                    new_data.append(data[row][col])
                if row > junk:
                    new_data.append(data[row][col])

    headers = data[0]
    del data[0]
    data[0] = data[0][-num_columns:len(data[0])]
    headers = headers[0:num_columns]

    if type == "play_type":
        junk -= 3
    data = data[0:junk]
    my_list_of_lists = [new_data[i:i + num_columns]
                        for i in range(0, len(new_data), num_columns)]
    for elem in my_list_of_lists:
        data.append(elem)

    df = pd.DataFrame(data, columns=headers)
    writer = pd.ExcelWriter(
        os.getcwd() + "/{}/{}/{}_{}.xlsx".format(year, team, type, team), engine='xlsxwriter')
    df = drop_columns(type, df)
    df.to_excel(writer, sheet_name='whole_team', index=False)
    workbook = writer.book
    ws = writer.sheets['whole_team']

    format_yellow = workbook.add_format({'bg_color': 'yellow'})
    players = {}
    playtype = ''
    for row in range(0, len(data)):
        if '#' in df.iloc[row, 0]:
            player_name = data[row][0]
            data[row][0] = '#'+playtype
            if player_name not in players:
                players[player_name] = [data[row]]
            else:
                players[player_name].append(data[row])
        else:
            playtype = data[row][0]
    format_excel(data, headers, df, ws, format_yellow, type)

    for p in players:
        new_df = pd.DataFrame(players[p], columns=headers)
        new_df.to_excel(writer, sheet_name=p, index=False)
        workbook = writer.book
        ws = writer.sheets[p]
        format_excel(players[p], headers, new_df, ws, format_yellow, type)

    writer.close()


def drop_columns(type, df):
    if type == "play_type":
        df = df.drop('FG MADE', axis=1)
        df = df.drop('FG MISS', axis=1)
        df = df.drop('2 FG MADE', axis=1)
        df = df.drop('2 FG MISS', axis=1)
        df = df.drop('3 FG MADE', axis=1)
        df = df.drop('3 FG MISS', axis=1)
        df = df.drop('%TIME RANK', axis=1)
        df = df.drop('PPP RATING', axis=1)
        df = df.drop('PPP RANK', axis=1)

    if type == "shot_type":
        df = df.drop('TO%', axis=1)
        df = df.drop('SCORE%', axis=1)
    # if type == "overall" or type == "overall_defense":
    #     df = df.drop('%SF', axis=1)
    return df


def overall(b):
    b.find_element(
        by=By.XPATH, value="//*[contains(text(), 'Overall')]").click()
    time.sleep(10)
    carets = b.find_elements(
        by=By.CLASS_NAME, value="la-caret-right")

    for c in reversed(carets):
        c.click()
    time.sleep(2)
    write_data_play_type(retrieve_data(17, b), "overall", 17)
    b.find_elements(
        by=By.XPATH, value="//*[contains(text(), 'Offense')]")[1].click()
    time.sleep(10)
    carets = b.find_elements(
        by=By.CLASS_NAME, value="la-caret-right")

    for c in reversed(carets):
        c.click()
    time.sleep(2)
    write_data_play_type(retrieve_data(17, b), "overall_defense", 17)


def format_excel(data, headers, df, ws, format_yellow, type):
    df.replace('%', '')
    df.replace(',', '')
    for row in range(0, len(data)):
        for col in range(0, len(df.columns)):
            try:
                if '%' in df.iloc[row, col]:
                    df.iloc[row, col] = df.iloc[row, col].replace('%', '')
                    df.iloc[row, col] = float(df.iloc[row, col])
            except:
                pass
            if df.iloc[row, 0] != '':
                try:
                    if float(df.iloc[row, 1]) >= 3:
                        if '#' in df.iloc[row, 0]:
                            try:
                                ws.write_number(
                                    row+1, col, float(df.iloc[row, col]))
                            except:
                                ws.write(
                                    row+1, col, df.iloc[row, col])
                        else:
                            try:
                                ws.write_number(
                                    row+1, col, float(df.iloc[row, col]), format_yellow)
                            except:
                                ws.write(
                                    row+1, col, df.iloc[row, col], format_yellow)
                    else:
                        ws.set_row(row+1, None, None, {'hidden': True})
                except:
                    pass

    highlighted_columns = []
    if type == "play_type":
        highlighted_columns = ['E', 'H', 'N', 'P']

    if type == "shot_type":
        highlighted_columns = ['E', 'H', 'N', 'P']

    for i in highlighted_columns:
        ws.conditional_format('{}2:{}{}'.format(
            i, i, len(data)), {'type': '3_color_scale', 'criteria': '=$B$2>260', })


def zip_files():
    # create a ZipFile object
    destination = os.getcwd() + '/{}/{}'.format(year, team)
    zipObj = ZipFile(destination + '/{}.zip'.format(team), 'w')
    # Add multiple files to the zip
    zipObj.write(destination + '/cumulative_box_{}.xlsx'.format(team),
                 arcname="cumulative_box_{}.xlsx".format(team))
    zipObj.write(destination + '/shot_type_{}.xlsx'.format(team),
                 arcname="shot_type_{}.xlsx".format(team))
    zipObj.write(destination + '/play_type_{}.xlsx'.format(team),
                 arcname="play_type_{}.xlsx".format(team))
    # zipObj.write(destination + '/overall_{}.xlsx'.format(team),
    #              arcname="overall_{}.xlsx".format(team))
    # zipObj.write(destination + '/overall_defense_{}.xlsx'.format(team),
    #              arcname="overall_defense_{}.xlsx".format(team))
    # close the Zip File
    zipObj.close()


login()
