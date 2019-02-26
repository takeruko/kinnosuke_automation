#!/usr/bin/env python
# -*- encoding: utf-8 -*-

import sys
import os.path
from argparse import ArgumentParser
import configparser
from datetime import datetime
import sqlite3
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


def get_argparser():
    parser = ArgumentParser(description=dedent("""勤之助で出退勤の打刻を行います。"""))
    parser.add_argument('record_type', metavar='IN|OUT', type=str,
                        help="""出勤か退勤を指定します。指定可能な値は'IN', 'OUT'のいずれかです [IN:出勤, OUT:退勤]""")
    parser.add_argument('--hide-browser', dest='hide_browser', action='store_true', default=False,
                        help="""Webブラウザを非表示にしてバックグラウンド実行ます。未指定時はWebブラウザを表示にして実行させます。""")

    tool_dir = os.path.dirname(__file__)    
    parser.add_argument('--config', metavar='PATH_TO_INI', type=str, default=os.path.join(tool_dir, 'TimeRecorder.ini'),
                        help="""勤之助のID,パスワードを記載した設定ファイルパスを指定します。未指定時はツールと同じフォルダのTimeRecorder.iniを使用します。""")
    parser.add_argument('--sqlite3', metavar='PATH_TO_DB', type=str, default=os.path.join(tool_dir, 'TimeRecord.sqlite3'),
                        help="""打刻履歴をローカルPC上に記録するためのDBファイルパスを指定します。未指定時はツールと同じフォルダのTimeRecord.sqlite3を使用します。""")

    return parser


def get_kinnosuke_config(config_file):
    config = configparser.ConfigParser()
    config.read(config_file)
    kinnosuke_conf = config['Kinnosuke']
    return (kinnosuke_conf['ID'], kinnosuke_conf['PASSWORD'], kinnosuke_conf['URL'])


def get_selenium_config(config_file):
    config = configparser.ConfigParser()
    config.read(config_file)
    selenium_conf = config['Selenium']
    return (selenium_conf['BROWSER'], selenium_conf['DRIVER_PATH'])


class TimeRecordDbManagaer:

    def __init__(self, sqlite3_file):
        self.__conn = sqlite3.connect(sqlite3_file)
        self.__conn.row_factory = sqlite3.Row
        self.__init_db()

    def __del__(self):
        self.__conn.close()

    def __get_date_key(self):
        return '{:%Y%m%d}'.format(datetime.now())

    def __init_db(self):

        DDL_TIME_RECORD = """
            create table if not exists time_record (
                id          integer     not null primary key autoincrement,
                date_key    text        not null,
                clock_type  text        not null,
                created_at  timestamp   not null default current_timestamp
            )
        """
        DDL_HOLIDAYS = """
            create table if not exists holidays (
                id          integer     not null primary key autoincrement,
                date_key    text        not null,
                yyyymm      text        not null,
                created_at  timestamp   not null default current_timestamp
            )
        """
        self.__conn.execute(DDL_TIME_RECORD)
        self.__conn.execute(DDL_HOLIDAYS)

    def has_initialized_thismonth_holidays(self):
        sql = "select count(1) as has_initialized from holidays where yyyymm = ?"
        yyyymm = str(self.__get_date_key()[0:6])
        rs = self.__conn.execute(sql, (yyyymm, ))
        return rs.fetchone()['has_initialized'] > 0

    def initialize_thismonth_holidays(self, holiday_date_keys):
        sql = "insert into holidays (date_key, yyyymm) values (?, ?)"
        for date_key in holiday_date_keys:
            yyyymm = date_key[0:6]
            self.__conn.execute(sql, (date_key, yyyymm))
        self.__conn.commit()  

    def __has_recorded(self, type):
        sql = "select count(1) as has_recorded from time_record where date_key = ? and clock_type = ?"
        rs = self.__conn.execute(sql, (self.__get_date_key(), type))
        return rs.fetchone()['has_recorded'] > 0

    def __record(self, type):
        sql = "insert into time_record (date_key, clock_type) values (?, ?)"
        self.__conn.execute(sql, (self.__get_date_key(), type))
        self.__conn.commit()

    def clock_in(self):
        self.__record('IN')

    def clock_out(self):
        self.__record('OUT')

    def has_clock_in(self):
        return self.__has_recorded('IN')

    def has_clock_out(self):
        return self.__has_recorded('OUT')

    def is_holiday(self, date_key=''):
        dkey = date_key if date_key != '' else self.__get_date_key()
        sql = 'select count(1) as is_holiday from holidays where date_key = ?'
        rs = self.__conn.execute(sql, (dkey,))
        return rs.fetchone()['is_holiday'] > 0


class KinnosukeAutomator:
    WAIT_SECONDS = 10

    def __init__(self, id, password, browser='Chrome', executable_path='', toppage_url='https://www.4628.jp/', hide_browser=False):
        
        if browser.upper() == 'FIREFOX':
            self.__driver = self.get_gecko_driver(executable_path, hide_browser)
        else:
            self.__driver = self.get_chrome_driver(executable_path, hide_browser)

        self.toppage_url = toppage_url
        self.timetable_url = toppage_url + '?module=timesheet&action=browse'

        self.__driver.get(self.toppage_url)
        # ログインボタンが表示されるまで待機
        WebDriverWait(self.__driver, self.WAIT_SECONDS).until(EC.presence_of_element_located((By.ID, 'id_passlogin')))

        # ID, パスワードを入力してログインボタンクリック
        id_box = self.__driver.find_element_by_id('y_logincd')
        id_box.send_keys(id)
        passwd_box = self.__driver.find_element_by_id('password')
        passwd_box.send_keys(password)
        login_button = self.__driver.find_element_by_id('id_passlogin')
        login_button.click()

    def get_chrome_driver(self, executable_path, hide_browser):
        options = webdriver.ChromeOptions()
        if hide_browser:
            options.add_argument('--headless')

        if executable_path == '':
            return webdriver.Chrome(options=options)
        else:
            return webdriver.Chrome(options=options, executable_path=executable_path)

    def get_gecko_driver(self, executable_path, hide_browser):
        options = webdriver.FirefoxOptions()
        if hide_browser:
            options.add_argument('-headless')

        if executable_path == '':
            return webdriver.Firefox(options=options)
        else:
            return webdriver.Firefox(options=options, executable_path=executable_path)

    def get_thismonth_holidays(self):
        self.__driver.get(self.timetable_url)
        # ページフッターが表示されるまで待機
        WebDriverWait(self.__driver, self.WAIT_SECONDS).until(EC.presence_of_element_located((By.ID, 'footer')))

        # タイムテーブルの表をparseして土日祝日のdate_keyをリストにして返却
        ID_PREFIX = 'fix_0_'
        yyyymm = '{:%Y%m}'.format(datetime.now())
        date_keys = []
        trs = self.__driver.find_elements_by_xpath('//tr[starts-with(@id, "{id_preifx}")]'.format(id_preifx=ID_PREFIX))
        for tr in trs:
            if tr.get_attribute('class') != 'bgcolor_white':  # bgcolor_white は平日
                datestr = tr.get_attribute('id').replace(ID_PREFIX, '').zfill(2)
                date_keys.append(yyyymm + datestr)
        return date_keys

    def clock_in(self):
        XPATH_FOR_CLOCK_IN = '//*[@id="timerecorder_txt" and (starts-with(text(), "出社") or starts-with(text(), "In"))]'
        self.__driver.get(self.toppage_url)
        # ページフッターが表示されるまで待機
        WebDriverWait(self.__driver, self.WAIT_SECONDS).until(EC.presence_of_element_located((By.ID, 'footer')))

        # 既に出社打刻済ならTrueを返して終了
        if self.__driver.find_elements_by_xpath(XPATH_FOR_CLOCK_IN):
            return True

        buttons = self.__driver.find_elements_by_name('_stampButton')
        for button in buttons:
            if button.text in ('出社', 'In'):
                button.click()
                # 出社時刻が表示されるまで待機
                WebDriverWait(self.__driver, self.WAIT_SECONDS).until(EC.presence_of_element_located((By.ID, 'footer')))
                WebDriverWait(self.__driver, self.WAIT_SECONDS).until(EC.presence_of_element_located((By.XPATH, XPATH_FOR_CLOCK_IN)))
                return True
        return False

    def clock_out(self):
        XPATH_FOR_CLOCK_OUT = '//*[@id="timerecorder_txt" and (starts-with(text(), "退社") or starts-with(text(), "Out"))]'
        self.__driver.get(self.toppage_url)
        # ページフッターが表示されるまで待機
        WebDriverWait(self.__driver, self.WAIT_SECONDS).until(EC.presence_of_element_located((By.ID, 'footer')))

        # 既に退社打刻済ならTrueを返して終了
        if self.__driver.find_elements_by_xpath(XPATH_FOR_CLOCK_OUT):
            return True

        buttons = self.__driver.find_elements_by_name('_stampButton')
        for button in buttons:
            if button.text in ('退社', 'Out'):
                button.click()
                # 退社時刻が表示されるまで待機
                WebDriverWait(self.__driver, self.WAIT_SECONDS).until(EC.presence_of_element_located((By.ID, 'footer')))
                WebDriverWait(self.__driver, self.WAIT_SECONDS).until(EC.presence_of_element_located((By.XPATH, XPATH_FOR_CLOCK_OUT)))
                return True
        return False

    def quit(self):
        self.__driver.quit()


def clock_in(mgr, id, password, browser, executable_path, url, hide_browser):
    if mgr.has_clock_in() or mgr.is_holiday():
        return
    automator = init_automator(mgr, id, password, browser, executable_path, url, hide_browser)
    if automator.clock_in():
        mgr.clock_in()
    automator.quit()


def clock_out(mgr, id, password, browser, executable_path, url, hide_browser):
    if mgr.has_clock_out() or mgr.is_holiday():
        return
    automator = init_automator(mgr, id, password, browser, executable_path, url, hide_browser)
    if automator.clock_out():
        mgr.clock_out()
    automator.quit()


def init_automator(mgr, id, password, browser, executable_path, url, hide_browser):
    automator = KinnosukeAutomator(id, password, browser, executable_path, url, hide_browser)

    if not mgr.has_initialized_thismonth_holidays():
        holidays = automator.get_thismonth_holidays()
        mgr.initialize_thismonth_holidays(holidays)

    return automator

if __name__ == '__main__':
    argparser = get_argparser()
    args = argparser.parse_args()

    mgr = TimeRecordDbManagaer(args.sqlite3)
    (id, password, url) = get_id_password(args.config)
    (browser, executable_path) = get_selenium_config(args.config)
    if args.record_type == 'IN':
        clock_in(mgr, id, password, browser, executable_path, url, args.hide_browser)
    elif args.record_type == 'OUT':
        clock_out(mgr, id, password, browser, executable_path, url, args.hide_browser)
    else:
        print("Error: '{record_type}' is invalid argument.".format(record_type=args.record_type), file=sys.stderr)
        argparser.print_usage()
