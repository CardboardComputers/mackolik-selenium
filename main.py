from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException

import os
import time
import datetime
import json

import probability_machine
import email_machine

# see https://stackoverflow.com/a/3850271
import atexit
# close the driver if you can, when you close the program
def exit_handler():
    global wd
    if wd is not None:
        wd.quit()
atexit.register(exit_handler)

with open('_config.json', 'r') as f:
    config = json.loads(f.read())
    
    FIREFOX_BINARY_PATH = config.get('firefoxBinaryPath')
    TIME_OFFSET_HOURS = config.get('timeOffsetHours')
    
    del config

LOOP_SLEEP_TIME_SECONDS = 60*30
MATCH_TIMEDELTA_MAX = datetime.timedelta(seconds=60*60*6)

with open('page_cleaner.js', 'r') as f:
    JS_PAGE_CLEANER = f.read()


def get_team_names(wd: webdriver.Firefox, node: WebElement) -> list:
    wh_main = wd.current_window_handle
    node.find_element(By.TAG_NAME, 'a').click()
    time.sleep(3)
    wd.switch_to.window(wd.window_handles[-1])
    team_names = list()
    team_names.append(
        wd.find_element(By.XPATH, '//a[@class="left-block-team-name"]')
        .text.strip().lower())
    team_names.append(
        wd.find_element(By.XPATH, '//a[@class="r-left-block-team-name"]')
        .text.strip().lower())
    wd.close()
    wd.switch_to.window(wh_main)
    return team_names


def get_league_table(wd: webdriver.Firefox, node: WebElement) -> list:
    wh_main = wd.current_window_handle
    node.click()
    time.sleep(3)
    wd.switch_to.window(wd.window_handles[-1])

    data_rows = list()
    try:
        try:
            # make sure it's not in the finals league, e.g. KOSTA Apertural Final
            # just choose the shortest name, that should do it (?)
            # no groups no finals etc.
            league_select: WebElement = wd.find_element(By.ID, 'Select2')
            league_names = [e.text for e in league_select.find_elements(By.TAG_NAME, 'option')]
            shortest_name = min(league_names, key=len)
            if len(shortest_name) < len(Select(league_select).first_selected_option.text):
                Select(league_select).select_by_visible_text(shortest_name)
                time.sleep(3)
        except NoSuchElementException:
            print('no sub-league selector')
        # go on to find the data table
        standing_table = wd.find_element(By.ID, 'tblStanding')
        table_rows = standing_table.find_elements(By.CSS_SELECTOR, '.puan_row')
        for r in table_rows:
            data_row = list()
            row_data = r.find_elements(By.XPATH, './td')
            for datum in row_data:
                data_row.append(datum.text.strip().lower())
            data_rows.append(data_row)
    except Exception as e:
        print(e)
    finally:
        wd.close()
        wd.switch_to.window(wh_main)
        return data_rows


DIR_PATH = os.path.dirname(os.path.abspath(__file__)).replace('/', '\\')
# see https://stackoverflow.com/a/55834112
options = Options()
options.binary_location = FIREFOX_BINARY_PATH
options.add_argument('--headless')
#options.add_argument('log-level=2')
wd: webdriver.Firefox = None
# these will be used to read the league table
IDX_TEAM_NAME = 1
IDX_TOTAL_MATCHES_PLAYED = 2
IDX_TOTAL_GOALS_SCORED = 6
IDX_TOTAL_GOALS_LOST = 7
IDX_HT_MATCHES_PLAYED = 11
IDX_HT_GOALS_SCORED = 15
IDX_HT_GOALS_LOST = 16
IDX_AT_MATCHES_PLAYED = 19
IDX_AT_GOALS_SCORED = 23
IDX_AT_GOALS_LOST = 24


file_list: list = list()

while True:
    try:
        file_list = list()
        wd = webdriver.Firefox(options=options)
        # see https://datarebellion.com/blog/using-firefox-extensions-with-selenium-in-python/
        wd.install_addon('{}\\ublock_origin-1.43.0.xpi'.format(DIR_PATH))
        wd.fullscreen_window()
        # record starting time
        dt_now = datetime.datetime.now() + datetime.timedelta(hours=TIME_OFFSET_HOURS)
        print('░▒▓ Iteration start {}'.format(dt_now.strftime('%Y.%m.%d %H:%M\'%S"')))

        # go to webpage
        wd.get('http://arsiv.mackolik.com/Iddaa-Programi')

        # change webpage settings
        Select(wd.find_element(By.ID, 'IddaaDateCmb')).select_by_value('-1')
        time.sleep(3)
        wd.find_element(By.XPATH, '//a[@href="#" and @type="2"]').click()
        time.sleep(3)
        wd.execute_script(JS_PAGE_CLEANER)

        # organise all the matches
        match_entries = wd.find_elements(By.XPATH, '//span[@id="iddaa-tab-body2"]/table/tbody/tr')

        # process relevant data
        skip_next = False
        list_date = dt_now
        for match_entry in match_entries:
            # check if we should skip this entry
            if skip_next:
                skip_next = False
                continue
            print('░▒▓ Entry:')
            # check if this entry is just a date header
            if match_entry.get_attribute('class') == 'iddaa-oyna-title2':
                datestr = match_entry.find_elements(By.TAG_NAME,'td')[0].text.strip()
                list_date = datetime.datetime.strptime(datestr, '%d.%m.%Y')
                skip_next = True
                print('is date header, skip next')
                continue
            # this entry is a match, figure out what time it starts
            entry_children = match_entry.find_elements(By.TAG_NAME, 'td')
            match_time = datetime.datetime.strptime(
                entry_children[0].get_attribute('innerHTML'), '%H:%M')
            match_dt = datetime.datetime(
                list_date.year,
                list_date.month,
                list_date.day,
                match_time.hour,
                match_time.minute
            )
            dt_difference = match_dt - dt_now
            print('time difference {}'.format(dt_difference))
            # check when the match takes place, relative to dt_now
            if dt_difference < MATCH_TIMEDELTA_MAX:
                # this match happens within the next hour, scrape data
                team_names = get_team_names(wd, entry_children[4])
                match_name = '{} {} - {}'.format(
                    match_dt.strftime('%Y.%m.%d %H.%M'), team_names[0], team_names[1])
                print('match {}'.format(match_name))
                league_table = get_league_table(wd, entry_children[1])
                time.sleep(1)
                if len(league_table) == 0:
                    # some leagues(?) such as ROK GKOK do not have standing tables
                    # just skip those I guess
                    print('no standing table for {} - {}'.format(team_names[0], team_names[1]))
                    continue
                ht_data = next(r for r in league_table if r[IDX_TEAM_NAME] == team_names[0])
                at_data = next(r for r in league_table if r[IDX_TEAM_NAME] == team_names[1])
                filename = '{}.xlsx'.format(match_name)
                probability_machine.write_spreadsheet(
                    filename=filename,

                    ht_total_matches_played=int(ht_data[IDX_TOTAL_MATCHES_PLAYED]),
                    ht_total_goals_scored=int(ht_data[IDX_TOTAL_GOALS_SCORED]),
                    ht_total_goals_lost=int(ht_data[IDX_TOTAL_GOALS_LOST]),
                    at_total_matches_played=int(at_data[IDX_TOTAL_MATCHES_PLAYED]),
                    at_total_goals_scored=int(at_data[IDX_TOTAL_GOALS_SCORED]),
                    at_total_goals_lost=int(at_data[IDX_TOTAL_GOALS_LOST]),
                    league_home_matches_played=sum(
                        [int(r[IDX_HT_MATCHES_PLAYED]) for r in league_table]),

                    ht_home_matches_played=int(ht_data[IDX_HT_MATCHES_PLAYED]),
                    ht_home_goals_scored=int(ht_data[IDX_HT_GOALS_SCORED]),
                    ht_home_goals_lost=int(ht_data[IDX_HT_GOALS_LOST]),
                    at_away_matches_played=int(at_data[IDX_AT_MATCHES_PLAYED]),
                    at_away_goals_scored=int(at_data[IDX_AT_GOALS_SCORED]),
                    at_away_goals_lost=int(at_data[IDX_AT_GOALS_LOST]),
                    league_total_matches_played=sum(
                        [int(r[IDX_TOTAL_MATCHES_PLAYED]) for r in league_table]),

                    league_home_goals=sum([int(r[IDX_HT_GOALS_SCORED]) for r in league_table]),
                    league_away_goals=sum([int(r[IDX_AT_GOALS_SCORED]) for r in league_table])
                )
                print('OK wrote file `{}`'.format(filename))
                file_list.append(filename)
            else:
                print('end entries')
                # this match is more than one hour later; break and go to sleep
                break

    except KeyboardInterrupt as e:
        os._exit(1)
    except Exception as e:
        print('Error in iteration: {}'.format(e))
    finally:
        # close window
        if wd is not None:
            wd.quit()
            wd = None
        # send emails and delete files
        if len(file_list) == 0:
            print('░▒▓ No matches this iteration')
        else:
            email_machine.send_update(file_list)
            print('░▒▓ Update sent, {} matches'.format(len(file_list)))
            for filename in file_list:
                if os.path.exists(filename):
                    os.remove(filename)
        file_list = list()
        # go to sleep
        print('░▒▓ Iteration end {}'.format(
            datetime.datetime.now().strftime('%Y.%m.%d %H:%M\'%S"')))
        time.sleep(LOOP_SLEEP_TIME_SECONDS)