# Imports
import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
from time import sleep as slp
import wikipedia
import pyjokes
import os
import warnings
import pyautogui
import shutil
import webbrowser
import requests
import wmi
import scrapy
from scrapy.crawler import CrawlerProcess
from GoogleNews import GoogleNews
from googletrans import Translator
from PIL import ImageGrab
import numpy as np
import cv2
from win32api import GetSystemMetrics
from datetime import datetime as dt
import pyperclip
import imaplib
import email
from email.header import decode_header


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
                listener.adjust_for_ambient_noise(source, duration=0.1)
                voice = listener.listen(source, phrase_time_limit=3)
                command = listener.recognize_google(voice)
                command = command.lower()
                return command
        except sr.UnknownValueError:
            pass
        except sr.RequestError:
            talk("Something went wrong, sir.")
            slp(10)


# take command
def take_command():
    while True:
        try:
            with sr.Microphone() as source:
                listener.adjust_for_ambient_noise(source, duration=0.1)
                voice = listener.listen(source, phrase_time_limit=6)
                command = listener.recognize_google(voice)
                command = command.lower()
                if 'james' in command:
                    command = command.replace('james', '')
                return command
        except sr.UnknownValueError:
            pass
        except sr.RequestError:
            talk("Something went wrong, sir.")
            slp(30)


# [SKILLS]

"""
Ideas for James
- weather prediction for the coming days
- read newsletters
- make James bylingual pt eng - he should detect the language used to talk to him and response properly
- write an event on a calendar with alarms
- Taking writen notes on google keep through voice
- create a inbox notifications with my warnings with priority scale
- urgency e-mails immediately announced 
- Personalized Speech recognition (pytorch)
- Making it able to work on android or IoT devices
- warns when my favorite youtubers post new videos
- read pdf
- Make it play some games (with and against me)
"""


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
    slp(2)
    pyautogui.press("enter")
    slp(2)
    urls = (
        "...",
        "..."
    )
    for url in urls:
        webbrowser.open(url, new=1)

    talk("E-mail is ready")


"""
def read_mail():
    outlook = win32com.client.Dispatch("Outlook.Appilication").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    message = messages.GetLast()
    talk(message.subject)
"""


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
def freelancer():
    week_day = dt.today().strftime('%A')
    hour = dt.now().hour
    first = True

    urls_d = (
        "...",
        "...",
        "...",
        "...",
        "..."
    )

    urls_segunda = (
        "...",
        "...",
        "...",
        "...",
        "..."
    )

    urls_terca = (
        "...",
        "...",
        "...",
        "...",
        "...",
        "..."
    )

    urls_quinta = (
        "...",
        "...",
        "...",
        "...",
        "...",
        "...",
        "...",
        "..."
    )

    urls_sabado = (
        "...",
        "...",
        "...",
        "...",
        "..."
    )

    abrir_urls = True

    if hour < 18 and week_day in 'MondayTuesdayWednesdayThursdayFriday':
        urls = urls_d
    elif week_day == 'Monday' and hour >= 18:
        urls = urls_segunda
    elif week_day == "Tuesday" and hour >= 18:
        urls = urls_terca
    elif week_day == 'Thursday' and hour >= 18:
        urls = urls_quinta
    elif week_day == 'Saturday' and hour >= 10 and hour < 12:
        urls = urls_sabado
    else:
        abrir_urls = False

    while abrir_urls:
        os.startfile("...")
        for url in urls:
            if first:
                webbrowser.open(url)
                first = False
            else:
                webbrowser.open(url, new=2)
        if abrir_urls:
            slp(3)
            pyautogui.hotkey("ctrl", "1")
            slp(2)
            pyautogui.hotkey("ctrl", "f4")
        else:
            pass
        abrir_urls = False


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


# work setup
def set_up():
    week_day = dt.today().strftime('%A')
    hour = dt.now().hour
    first = True

    password = '...'
    username = '...'
    password_vpn = '...'
    pyperclip.copy(password_vpn)
    pos_password_vpn = (877, 701)
    pos_code_mail = (872, 767)

    urls = (
        "...",
        "...",
        "..."
    )

    open_urls = True

    if hour < 16 and week_day in 'MondayTuesdayWednesdayThursdayFriday':
        urls2 = urls
    else:
        open_urls = False

    while open_urls:
        os.startfile(r'...')
        os.startfile(r'...')
        for url in urls2:
            if first:
                webbrowser.open(url)
                first = False
            else:
                webbrowser.open(url, new=2)
        if open_urls:
            slp(3)
            pyautogui.hotkey("ctrl", "1")
            slp(2)
            pyautogui.hotkey("ctrl", "f4")
        else:
            pass

        slp(5)
        pyautogui.hotkey("alt", "tab", "tab", "tab")
        slp(2)
        pyautogui.click(pos_password_vpn, clicks=1)
        password_vpn = '...'
        pyperclip.copy(password_vpn)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.hotkey("enter")
        slp(15)
        imap = imaplib.IMAP4_SSL("....")
        imap.login(username, password)
        status, messages = imap.select("...")
        N = 1
        messages = int(messages[0])
        for i in range(messages, messages - N, -1):
            res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            code = str(body)[-7:].replace('.', '')
                        elif "attachment" in content_disposition:
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    os.mkdir(folder_name)
                                filepath = os.path.join(folder_name, filename)
                                open(filepath, "wb").write(part.get_payload(decode=True))
                else:
                    content_type = msg.get_content_type()
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        print(body)
                        code = str(body)[-11:].strip().replace('.', '')
        imap.close()
        imap.logout()
        pyautogui.click(pos_code_mail, clicks=1)
        pyperclip.copy(code)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.hotkey("enter")
        slp(5)
        os.startfile(r'...')
        os.startfile(r'...')

        open_urls = False


# Scheduled_tasks
def scheduled():
    week_day = dt.today().strftime('%A')
    hour =  f'{dt.now().hour}:{dt.now().minute}'
    if week_day in 'MondayTuesdayWednesdayThursdayFriday' and hour == '15:0':
        freelancer()
        slp(60)
    elif week_day == 'Saturday' and hour == '10:0':
        freelancer()
        slp(60)
    elif week_day in 'MondayTuesdayWednesdayThursdayFriday' and hour == '9:0':
        work_set_up()
        slp(60)
    elif week_day in 'MondayTuesdayWednesdayThursdayFriday' and hour == '10:0':
        talk("It's almost in time for the daily meeting, sir.")
    elif week_day in 'MondayTuesdayWednesdayThursdayFriday' and hour == '10:10':
        talk("It's almost in time for the daily meeting, sir.")
    elif week_day in 'MondayTuesdayWednesdayThursdayFriday' and hour == '10:15':
        talk("It's time for the daily meeting, sir.")
    else:
        pass


# conveniences
def gkeep():
    url = "https://keep.google.com"
    webbrowser.open(url, new=1)
    talk("Google Keep is ready.")


def screen_recorder():
    talk('Pause with P key when needed. Five seconds to start.')
    talk('five')
    slp(1)
    talk('four')
    slp(1)
    talk('three')
    slp(1)
    talk('two')
    slp(1)
    talk('one')
    slp(1)
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


"""
# Historical quotation
def historical_quotation(text):
    cad_quotation = web.DataReader('CADBRL=X', data_source='yahoo', start="12-01-2020", end="07-12-2021")
    dol_quotation = web.DataReader('BRL=X', data_source='yahoo', start="12-01-2020", end="07-12-2021")
    eur_quotation = web.DataReader('EURBRL=X', data_source='yahoo', start="12-01-2020", end="07-12-2021")
    cny_quotation = web.DataReader('CNYBRL=X', data_source='yahoo', start="12-01-2020", end="07-12-2021")
    btc_quotation = web.DataReader('BTCBRL=X', data_source='yahoo', start="12-01-2020", end="07-12-2021")
    eth_quotation = web.DataReader('ETHBRL=X', data_source='yahoo', start="12-01-2020", end="07-12-2021")
"""


# News reading
def news(text):
    text = text.replace('news', "")
    if "tell me" in text:
        text = text.replace("tell me", "")
    google_news = GoogleNews()
    google_news = GoogleNews(period='d')

    def news_reading():
        talk('The title is:')
        slp(1)
        talk(x['title'])
        slp(1)
        talk('Description:')
        slp(1)
        talk(x['desc'])

    def leitura_noticias():
        talk('O título da notícia é:')
        slp(1)
        talk(x['title'])
        slp(1)
        talk('Descrição:')
        slp(1)
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

# error case
def error_case():
    talk('Sorry, this function is unavailable at the moment')

# AI chat
"""def bot_chat():
    talk('Okay, what would you like to talk about?')
    in_conversation = True
    while in_conversation:
        conv1_start = take_command()
        conv1 = Conversation(conv1_start)
        bot_speaks = str(conversational_pipeline([conv1]))
        speak_starts = int(bot_speaks.find('bot >>')) + 7
        talk(f"{bot_speaks[speak_starts:]}")
        if conv1_start == 'goodbye':
            in_conversation = False
            break
        conv1_next = take_command()
        conv1.add_user_input(conv1_next)
        bot_speaks = str(conversational_pipeline([conv1]))
        speak_starts = int(bot_speaks.find('bot >>')) + 7
        bot_speaks = bot_speaks[speak_starts:]
        speak_starts = int(bot_speaks.find('bot >>')) + 7
        talk(f"{bot_speaks[speak_starts:]}")
        if conv1_next == 'goodbye':
            in_conversation = False
            break
"""

# tell me all the commands
def tell_me_all_commands():
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
        freelancer setup
        quotation
        tell me all commands
        record the screen
  """)


# Core of commands
def run_james(order):
    command = order
    if 'weather' in command:
        try:
            weather(command)
        except:
            error_case()
    elif 'play' in command:
        try:
            playing(command)
        except:
            error_case()
    elif ('time' or 'hour') in command:
        try:
            what_hour()
        except:
            error_case()
    elif 'search for' in command:
        try:
            search_wiki(command)
        except:
            error_case()
    elif 'joke' in command:
        try:
            talk(pyjokes.get_joke())
        except:
            error_case()
    elif shutdown(command):
        try:
            shutdown(command)
        except:
            error_case()
    elif "volume" in command:
        try:
            volume_control(command)
        except:
            error_case()
    elif "light" in command:
        try:
            light_control(command)
        except:
            error_case()
    elif ("screenshot") in command:
        try:
            screenshot()
        except:
            error_case()
    elif ("resume" or "pause") in command:
        try:
            resume_pause()
        except:
            error_case()
    elif "game" in command:
        try:
            brawlhalla()
        except:
            error_case()
    elif "open mail" in command:
        try:
            open_mail()
        except:
            error_case()
    elif "close" in command:
        try:
            close_window()
        except:
            error_case()
    elif "kill tab" in command:
        try:
            kill_tab()
        except:
            error_case()
    elif "switch window" in command:
        try:
            switch_window()
        except:
            error_case()
    elif ('set up' or 'freelancer') in command:
        try:
            freelancer()
        except:
            error_case()
    elif 'work' or 'begin' in command:
        try:
            work()
        except:
            error_case()
    elif "quotation" in command:
        try:
            get_quotation(command)
        except:
            error_case()
    elif "news" in command:
        try:
            news(command)
        except:
            error_case()
    elif 'commands' in command:
        try:
            tell_me_all_commands()
        except:
            error_case()
    elif "full" in command:
        try:
            full_screen()
        except:
            error_case()
    elif "open notes" in command:
        try:
            gkeep()
        except:
            error_case()
    elif "record the screen" in command:
        try:
            screen_recorder()
        except:
            error_case()
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
        # checking schedule
        scheduled()

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