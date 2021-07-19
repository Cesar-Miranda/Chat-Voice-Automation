# Imports
import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import time
import wikipedia
import pyjokes
import os
import sys
import warnings
import pyautogui
import shutil
import webbrowser
import requests
import json
import wmi
import pandas as pd
import scrapy
import Scrapy_crawler
from scrapy.crawler import CrawlerProcess
from pandas_datareader import data as web
from GoogleNews import GoogleNews
from googletrans import Translator
# import matplotlib.pyplot as plt
# import win32com.client


# [BASIC FUNCTIONS]

warnings.filterwarnings('ignore')

# Listener
listener = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


# Talk
def talk(text):
    engine.say(text)
    engine.runAndWait()


# Wake word
def wake_word(text):
    keywords = ['hey', 'james']
    text = text.lower()
    for phrase in keywords:
        if phrase in text:
            return True
    return False


# take wake word
def take_wake_word():
    while True:
        try:
            with sr.Microphone() as source:
                listener.adjust_for_ambient_noise(source, duration=0.2)
                voice = listener.listen(source, phrase_time_limit=2)
                command = listener.recognize_google(voice)
                command = command.lower()
                return command
        except sr.UnknownValueError:
            pass
        except sr.RequestError:
            print("Something went wrong, sir.")
            talk("Something went wrong, sir.")


# take command
def take_command():
    while True:
        try:
            with sr.Microphone() as source:
                listener.adjust_for_ambient_noise(source, duration=0.2)
                voice = listener.listen(source, phrase_time_limit=5)
                command = listener.recognize_google(voice)
                command = command.lower()
                if 'james' in command:
                    command = command.replace('james', '')
                return command
        except sr.UnknownValueError:
            pass
        except sr.RequestError:
            print("Something went wrong, sir.")
            talk("Something went wrong, sir.")


# [SKILLS]

# weather today - scrapy crawls css selectors of weather.com
def weather(command):
    if "today" in command:
        class WeatherSpider(scrapy.Spider):
            name = 'weather'

            def start_requests(self):
                yield scrapy.Request('https://weather.com/weather/today/l/-12.98,-38.50?par=google')

            def parse(self, response):
                items = {}
                temperature_now = response.css('span.CurrentConditions--tempValue--1RYJJ::text').extract()
                temperature_now_in_celsius = temperature_now[0]

                temperature_now_in_celsius = (int(temperature_now_in_celsius[:2]) - 32) * 5 / 9
                weather_now = str(response.css('div.CurrentConditions--phraseValue--17s79::text').extract())
                str_temperature_now_in_celsius = str(temperature_now_in_celsius)
                str_weather_now = str(weather_now)
                items['temperature_now'] = temperature_now_in_celsius
                items['weather_now'] = weather_now
                yield items
                talk(
                    'In Salvador, the temperature is ' + str_temperature_now_in_celsius[:3]
                    + ' and the weather is: ' + str_weather_now)

    process = CrawlerProcess(settings = {
        'FEED_URI': 'weather.csv',
        'FEED_FORMAT': 'csv'
    })

    process.crawl(WeatherSpider)
    process.start()


# Shutdown the computer - os 
def shutdown(text):
    shutdown_words = ['turn off', 'good night', 'see you later']
    for phrase in shutdown_words:
        if phrase in text:
            print("Do you really want to turn this computer off, sir?")
            talk("Do you really want to turn this computer off, sir?")
            choice = take_command()
            if "yes" in choice:
                print("Shutting down in a few seconds...")
                talk("Shutting down in a few seconds...")
                os.system("shutdown /s /t 30")
            elif "no" in choice:
                print("See you later, sir.")
                talk("See you later, sir.")
                break
    # for fast shutdown
    if 'shutdown' in text:
        print("See you later, sir.")
        talk("See you later, sir.")
        os.system("shutdown /s /t 3")


# [Youtube]
# Opening youtube
def playing(self):
    video = self.replace('play', '')
    video = video.replace('charlie', '')
    talk('playing ' + video + ' sir')
    pywhatkit.playonyt(video)


# fullscreen
def full_screen():
    pyautogui.hotkey("f")


# Pausing videos.
def resume_pause():
    pyautogui.hotkey("playpause")
    talk("It's done, sir.")


# volume down/up
def volume_control(text):
    if "down" in text:
        for i in range(1, 20):
            pyautogui.hotkey("volumedown")
    elif "up" in text:
        for i in range(1, 20):
            pyautogui.hotkey("volumeup")


# brightness up
def light_control(text):
    brightness = 0
    if "up" in text:
        brightness = 100
    elif "down" in text:
        brightness = 30
    c = wmi.WMI(namespace='wmi')
    methods = c.WmiMonitorBrightnessMethods()[0]
    methods.WmiSetBrightness(brightness, 0)


# Searching on Wikipedia
def search_wiki(self):
    research = self.replace('search for', '')
    info = wikipedia.summary(research, 2)
    print("This is what I found about " + research + " sir:" + info)
    talk("This is what I found about " + research + " sir: " + info)


# Telling the Hour
def what_hour():
    hour = datetime.datetime.now().strftime('%I:%M %p')
    talk("It's is " + hour + " sir")


# Open game on steam
def brawlhalla():
    os.startfile("steam://rungameid/291550")
    talk("Openning Brawlhalla, Sir.")


# [Mail tasks - windows app]
def open_mail():
    pyautogui.hotkey("win")
    pyautogui.write("mail")
    time.sleep(2)
    pyautogui.press("enter")
    talk("Opening e-mail, sir.")


# windows and tabs management
def close_window():
    pyautogui.hotkey("alt", "f4")
    talk("window closed, sir.")


def kill_tab():
    pyautogui.hotkey("ctrl", "f4")
    talk("tab closed, sir.")


def switch_window():
    pyautogui.hotkey("alt", "tab")
    talk("window, switched, sir")


# take a screen shot - saving the screen in a chosen directory with date and time registered in its name.
def screenshot():
    hour = datetime.datetime.now().strftime('%b_%d_%Y__%H_%M_%S')
    pyautogui.screenshot(f"screenshot_{hour}.png")
    source = f"screenshot_{hour}.png"
    destination = "C:\\Users\Cesar\\Pictures\\Saved Pictures"
    shutil.move(source, destination)
    talk("screen printed, sir")


# work setups - setting up my daily worksites without the need of opening them one by one.
def kumon():
    first = True
    os.startfile("C:\Program Files\Google\Chrome\Application\chrome.exe")
    urls = (
        "url of choice...", # hidden for privacy
        "url of choice...",
        "url of choice...",
        "url of choice..."
    )
    for url in urls:
        if first:
            webbrowser.open(url)
            first = False
        else:
            webbrowser.open(url, new=2)
    time.sleep(3)
    pyautogui.hotkey("ctrl", "1")
    time.sleep(2)
    pyautogui.hotkey("ctrl", "f4")
    talk("Kumon setup is ready, sir.")


def desenvolver():
    first = True
    os.startfile("C:\Program Files\Google\Chrome\Application\chrome.exe")
    urls = (
        "url of choice...",
        "url of choice...",
        "url of choice...",
        "url of choice...",
        "url of choice..."
    )
    for url in urls:
        if first:
            webbrowser.open(url)
            first = False
        else:
            webbrowser.open(url, new=2)
    time.sleep(3)
    pyautogui.hotkey("ctrl", "1")
    time.sleep(2)
    pyautogui.hotkey("ctrl", "f4")
    talk("Desenvolver setup is ready, sir.")


# today's quotations - request to quotation api (Dol, Cad, Eur, Cny, Bit and Eth)
def get_quotation(text):
    quotation = requests.get("https://economia.awesomeapi.com.br/last/CAD-BRL,USD-BRL,EUR-BRL,CNY-BRL,BTC-BRL,ETH-BRL")
    quotation = quotation.json()
    quot_dol = quotation["USDBRL"]["bid"]
    quot_cad = quotation["CADBRL"]["bid"]
    quot_eur = quotation["EURBRL"]["bid"]
    quot_bth = quotation["BTCBRL"]["bid"]
    quot_cny = quotation["CNYBRL"]["bid"]
    quot_eth = quotation["ETHBRL"]["bid"]
    quotations = [quot_cad, quot_dol, quot_eur, quot_cny, quot_bth, quot_eth]

    if "canadian" in text:
        talk(f"One Canadian dolar, at this moment, is: {quot_cad} BRL.")
    elif "dolar" in text:
        talk(f"One Dolar, at this moment is: {quot_dol} BRL.")
    elif "euro" in text:
        talk(f"One Euro, at this moment, is: {quot_eur} BRL.")
    elif "chinese" in text:
        talk(f"One Chinese Yuan,at this moment, is: {quot_cny} BRL.")
    elif "bitcoin" in text:
        talk(f"One Bitcoin, at this moment, is: {quot_bth} BRL.")
    elif "ethereum" in text:
        talk(f"One Ethereum, at this moment, is: {quot_eth} BRL.")
    else:
        i = 0
        for key in quotation:
            talk(f" One {key[:3]}, at this moment, is {quotations[i]} BRL.")
            i += 1


# News reading - google_news lib + googgle_trans lib = news in english and portuguese
def news(text):
    text = text.replace('news', "")
    if "tell me" in text:
        text = text.replace("tell me", "")
    google_news = GoogleNews()
    google_news = GoogleNews(period='d')

    def news_reading():
        talk('The title is:')
        time.sleep(1)
        talk(x['title'])
        time.sleep(1)
        talk('Description:')
        time.sleep(1)
        talk(x['desc'])

    def leitura_noticias():
        talk('O título da notícia é:')
        time.sleep(1)
        talk(x['title'])
        time.sleep(1)
        talk('Descrição:')
        time.sleep(1)
        talk(x['desc'])

    if "in portuguese" in text:
        text = text.replace("in portuguese", "")
        engine.setProperty('voice', voices[2].id)
        google_news.setlang('pt')
        detector = Translator()
        lang_detector = detector.detect(text)
        # translating the requisition to portuguese
        if lang_detector.lang != 'pt':
            lang_detector = detector.translate(text, dest='pt')
            text = lang_detector.text
        google_news.search(text)
        result = google_news.result()
        talk(f'As notícias sobre {text}.')
        for x in result:
            leitura_noticias()
        engine.setProperty('voice', voices[0].id)
    elif "in english" in text:
        text = text.replace("in english", "")
        google_news.search(text)
        result = google_news.result()
        talk(f'The news about {text}, sir: ')
        for x in result:
            news_reading()
    else:
        talk('I could not find any news about' + text)


# tell me all the commands
def tell_me_all_commands():
    print("""
        weather today
        play
        time
        search for
        joke
        shutdown
        turn off
        volume up
        volume down
        close window
        switch window
        open mail
        read mail
        kill tab
        switch tab
        screenshot
        full mode
        pause
        resume
        game
        close
        set up kumon
        start desenvolver mais
        quotation
        news about something in english
        news about something in portuguese
        tell me all commands
    """)
    talk("""
        weather today
        play
        time
        search for
        joke
        shutdown
        turn off
        volume up
        volume down
        close window
        switch window
        open mail
        read mail
        kill tab
        switch tab
        screenshot
        full mode
        pause
        resume
        game
        close
        set up kumon
        start desenvolver mais
        quotation
        tell me all commands
        news about something in english
        news about something in portuguese
  """)


# Core of commands
def run_james(order):
    command = order
    if 'weather' in command:
        weather(command)
    elif 'play' in command:
        playing(command)
    elif ('time' or 'hour') in command:
        what_hour()
    elif 'search for' in command:
        search_wiki(command)
    elif 'joke' in command:
        talk(pyjokes.get_joke())
    elif shutdown(command):
        shutdown(command)
    elif "volume" in command:
        volume_control(command)
    elif "light" in command:
        light_control(command)
    elif ("screen" or "shot" or "screenshot") in command:
        screenshot()
    elif ("resume" or "pause") in command:
        resume_pause()
    elif "game" in command:
        brawlhalla()
    elif "open mail" in command:
        open_mail()
    elif "close" in command:
        close_window()
    elif ("kill" or "tab") in command:
        kill_tab()
    elif "switch window" in command:
        switch_window()
    elif ('set up' or 'kumon') in command:
        kumon()
    elif ('start' or 'desenvolver' or 'mais') in command:
        desenvolver()
    elif "quotation" in command:
        get_quotation(command)
    elif "news" in command:
        news(command)
    elif 'commands' in command:
        tell_me_all_commands()
    elif "full" in command:
        full_screen()
    else:
        talk("Sorry, I didn't understand what you said.")


# Running James
def running_james():
    print("I'm listening, sir...")
    talk("I'm listening, sir...")
    command = take_command()
    if "rest" in command:
        talk("Okay, sir.")
        pass
    elif "rest" not in command:
        run_james(command)
        talk("Anything else, sir?")
        repeat = take_command()
        while "yes" in repeat:
            print("I'm Listening...")
            talk("I'm Listening...")
            command = take_command()
            run_james(command)
            talk("Anything else, sir?")
            repeat = take_command()
        if "no" in repeat:
            print("Okay, sir.")
            talk("Okay, sir.")
            pass


# James Execution
def james_exec():
    talk("I'm on, sir.")
    while True:
        # taking the wake word
        call = take_wake_word()

        # fast shutdown
        if shutdown(call):
            shutdown(call)
        elif wake_word(call):
            running_james()


# Main
if __name__ == '__main__':
    while True:
        james_exec()
        os.execv(james_bot.py, sys.argv)
