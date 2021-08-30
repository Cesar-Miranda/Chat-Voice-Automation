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
from PIL import ImageGrab
import numpy as np
import cv2
from win32api import GetSystemMetrics

# BASIC FUNCTIONS

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
    keywords = ['hey james', 'james']
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
            talk("Something went wrong, sir.")
            time.sleep(10)


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
            talk("Something went wrong, sir.")
            time.sleep(10)


# [SKILLS]


# weather today
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
                hour_of_command = datetime.datetime.now().strftime('%I:%M %p')
                day_of_command = datetime.datetime.now().strftime('%b_%d_%Y')
                weather_now = str(response.css('div.CurrentConditions--phraseValue--17s79::text').extract())
                str_temperature_now_in_celsius = str(temperature_now_in_celsius)
                str_weather_now = str(weather_now)
                items['temperature_now'] = temperature_now_in_celsius
                items['weather_now'] = weather_now
                items['hour'] = hour_of_command
                items['day'] = day_of_command
                yield items
                talk(
                    'In Salvador, the temperature is ' + str_temperature_now_in_celsius[:3]
                    + ' celsius and the weather is: ' + str_weather_now)
        process = CrawlerProcess(settings={
            'FEED_URI': 'weather.csv',
            'FEED_FORMAT': 'csv'
        })

        process.crawl(WeatherSpider)
        process.start()

    elif "night" in command:
        class WeatherSpider(scrapy.Spider):
            name = 'weather'

            def start_requests(self):
                yield scrapy.Request('https://weather.com/weather/today/l/-12.98,-38.50?par=google')

            def parse(self, response):
                items = {}
                temperature_now = response.css('span.CurrentConditions--tempValue--1RYJJ::text').extract()
                temperature_now_in_celsius = temperature_now[0]
                temperature_now_in_celsius = (int(temperature_now_in_celsius[:2]) - 32) * 5 / 9
                hour_of_command = datetime.datetime.now().strftime('%I:%M %p')
                day_of_command = datetime.datetime.now().strftime('%b_%d_%Y')
                weather_now = str(response.css('div.CurrentConditions--phraseValue--17s79::text').extract())
                str_temperature_now_in_celsius = str(temperature_now_in_celsius)
                str_weather_now = str(weather_now)
                items['temperature_now'] = temperature_now_in_celsius
                items['weather_now'] = weather_now
                items['hour'] = hour_of_command
                items['day'] = day_of_command
                yield items
                talk(
                    'In Salvador, the temperature is ' + str_temperature_now_in_celsius[:3]
                    + ' celsius and the weather is: ' + str_weather_now)
        process = CrawlerProcess(settings={
            'FEED_URI': 'weather.csv',
            'FEED_FORMAT': 'csv'
        })

        process.crawl(WeatherSpider)
        process.start()


"""    elif "evening" in command:
    elif "tomorrow morning" in command:
    elif "tomorrow night" in command:
    elif "tomorrow afternoon" in command:
    elif "tomorrow" in command:
    elif "monday" in command:
    elif "tuesday" in command:
    elif "wednesday" in command:
    elif "thursday" in command:
    elif "friday" in command:
    elif "saturday" in command:
    elif "sunday" in command:
    """

# Shutdown the computer
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
    video = video.replace('James', '')
    talk('playing ' + video)
    pywhatkit.playonyt(video)


# fullscreen
def full_screen():
    pyautogui.hotkey("f")


# Pausing videos.
def resume_pause():
    pyautogui.hotkey("playpause")


# volume down
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


# speed up video
# next video
# download video from yt


# Searching on Wikipedia
def search_wiki(self):
    research = self.replace('search for', '')
    info = wikipedia.summary(research, 2)
    talk("This is what I found about " + research + " sir: " + info)


# Telling the Hour
def what_hour():
    hour = datetime.datetime.now().strftime('%I:%M %p')
    talk("It's is " + hour)


# Open game on steam
def brawlhalla():
    os.startfile("steam://rungameid/291550")
    talk("Openning Brawlhalla.")


# [Mail tasks]
def open_mail():
    pyautogui.hotkey("win")
    pyautogui.write("mail")
    time.sleep(2)
    pyautogui.press("enter")
    talk("E-mail is ready")


# windows and tabs management
def close_window():
    pyautogui.hotkey("alt", "f4")
    talk("window closed")


def kill_tab():
    pyautogui.hotkey("ctrl", "f4")
    talk("tab closed")


def switch_window():
    pyautogui.hotkey("alt", "tab")
    talk("window, switched")


# take a screen shot
def screenshot():
    hour = datetime.datetime.now().strftime('%b_%d_%Y__%H_%M_%S')
    pyautogui.screenshot(f"screenshot_{hour}.png")
    source = f"screenshot_{hour}.png"
    destination = "C:\\Users\Cesar\\Pictures\\Saved Pictures"
    shutil.move(source, destination)
    talk("screen printed.")


# work setups
def kumon():
    first = True
    os.startfile("C:\Program Files\Google\Chrome\Application\chrome.exe")
    urls = (
        "https://...",
        "http://...",
        "http://...",
        "http://..."
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
    talk("Kumon setup is ready.")


def desenvolver():
    first = True
    os.startfile("C:\Program Files\Google\Chrome\Application\chrome.exe")
    urls = (
        "https://...",
        "https://...",
        "http://...",
        "https://...,
        "https://..."
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
    talk("Desenvolver setup is ready.")

# conveniences
def gkeep():
    url = "https://keep.google.com/u/0/"
    webbrowser.open(url, new=1)
    talk("Google Keep is ready.")

def screen_recorder():
    talk('Pause with P key when needed. Five seconds to start.')
    talk('five')
    time.sleep(1)
    talk('four')
    time.sleep(1)
    talk('three')
    time.sleep(1)
    talk('two')
    time.sleep(1)
    talk('one')
    time.sleep(1)
    talk('now')

    width = GetSystemMetrics(0)
    height = GetSystemMetrics(1)
    day_hour = datetime.datetime.now().strftime('%b_%d_%Y__%H_%M_%S')
    fourcc = cv2.VideoWriter_fourcc('m', 'p', '4', 'v')
    captured_video = cv2.VideoWriter(f'Screen Record - {day_hour}.mp4', fourcc, 20.0, (width, height))

    while True:
        img = ImageGrab.grab(bbox=(0, 0, width, height,))
        img_np = np.array(img)
        img_final = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
        cv2.imshow('Secret Capture', img_final)
        captured_video.write(img_final)
        if cv2.waitKey(10) == ord('p'):
            break

    source = f"Screen Record - {day_hour}.mp4"
    destination = "C:\\Users\Cesar\\Videos\\Captures"
    shutil.move(source, destination)



# today's quotations
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


# News reading
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
    print(""" Those are the commands you can use:
        weather
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
    """)
    talk("""Those are the commands you can use:
        weather
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
  """)


# Core of commands
def run_james(order):
    command = order
    if 'weather' in command:
        try:
            weather(command)
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif 'play' in command:
        try:
            playing(command)
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif ('time' or 'hour') in command:
        try:
            what_hour()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif 'search for' in command:
        try:
            search_wiki(command)
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif 'joke' in command:
        try:
            talk(pyjokes.get_joke())
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif shutdown(command):
        try:
            shutdown(command)
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "volume" in command:
        try:
            volume_control(command)
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "light" in command:
        try:
            light_control(command)
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif ("screenshot") in command:
        try:
            screenshot()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif ("resume" or "pause") in command:
        try:
            resume_pause()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "game" in command:
        try:
            brawlhalla()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "open mail" in command:
        try:
            open_mail()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "close" in command:
        try:
            close_window()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "kill tab" in command:
        try:
            kill_tab()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "switch window" in command:
        try:
            switch_window()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif ('set up' or 'kumon') in command:
        try:
            kumon()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif ('start' or 'desenvolver' or 'mais') in command:
        try:
            desenvolver()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "quotation" in command:
        try:
            get_quotation(command)
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "news" in command:
        try:
            news(command)
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif 'commands' in command:
        try:
            tell_me_all_commands()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "full" in command:
        try:
            full_screen()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "open notes" in command:
        try:
            gkeep()
        except:
            talk('Sorry, this function is unavailable at the moment')
    elif "record the screen" in command:
        try:
            screen_recorder()
        except:
            talk('Sorry, this function is unavailable at the moment')
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
        elif "take a nap" in call:
            talk('See you later, sir.')
            return 'sleep'
        elif wake_word(call):
            running_james()


# Main
if __name__ == '__main__':
    looping = True
    while looping:
        try:
            james_do = james_exec()
            print(james_do)
            if james_do == 'sleep':
                looping = False
        except:
            if looping:
                talk('Reconnecting, sir.')
            else:
                pass