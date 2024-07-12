import os
import time
from turtle import bgcolor
import pyttsx3
import speech_recognition as sr

import nltk
import replicate                                    #image creating with prompt
import sklearn
import tensorflow
import torch
import spacy

import datetime
import platform
import wikipedia                                   # pip install wikipedia
import webbrowser
import subprocess
import sys
import smtplib
from email.message import EmailMessage
import pywhatkit                                   # pip install pywhatkit. Whatsapp and Youtube automation and more
#import MyAlarm      
#import ecapture as ec                
import pyjokes                                     # pip install pyjokes
from speedtest import Speedtest                    # pip install speedtest-cli
from pywikihow import search_wikihow               # pip install pywikihow
import pyautogui                                   # pip install pyAutoGUI
import poetpy                                      # pip install poetpy
import random
from forex_python.converter import CurrencyRates   # pip install forex-python
import requests                                    # pip install requests
import bs4                                         # pip install beautifulsoup4
import time
import wolframalpha                                # pip install wolframalpha
from quote import quote                            # pip install quote
import winshell as winshell                        # pip install winshell
from geopy.geocoders import Nominatim              # pip install geopy  and pip install geocoder
from geopy import distance
import pygetwindow as gw 

import requests                                    #we can use the requests library to download the contents of the page of interest. Used here to extract news from bbc news website.
from bs4 import BeautifulSoup                      #pip or pip3 install beautifulsoup4. Import BeautifulSoup and created an object called soup that parses HTML pages (for news query).
from PIL import Image                              #PIL for name 'Image' in "images" query
import win32com.client
import ctypes
import comtypes.client
import getpass
import keyboard

import PyPDF2                                      #to read text from PDFs. Accompanied with pyttsx3, to convert text to speech. Could also use PyPDF and Pyttsx3 to convert a webpage into an audio file
import fitz  # PyMuPDF
from collections import Counter

import replicate

from selenium import webdriver
import psutil



if platform.system() == "Windows":
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)
    userSaid = "hello world"
else:
    engine = pyttsx3.init()
    engine.setProperty('voice', 'english-us')
    engine.setProperty('rate', 160)
    userSaid = "hello world"


# Wishing Function
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

# Speak Function


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

# Listening Function


def takecommand(wtr=0):
    r = sr.Recognizer()
    r.pause_threshold = 2
    r.operation_timeout = 5
    with sr.Microphone() as source:
        speak("Listening...")
        print("Listening...")
        try:
            audio = r.listen(source, phrase_time_limit=8)
        except sr.WaitTimeoutError:
            speak("Listening stopped due to silence.")
            print("Listening stopped due to silence.")
            return 0

        try:
            print("Recognizing...\n")
            heard = r.recognize_google(audio)
            print(f"You Said: \"{heard}\"")
            return heard.lower()
        
        except sr.UnknownValueError:
            speak("I didn't understand what you said.")
            print("You said something that is beyond my understanding or maybe you didn't say anything.\n")
            return 0
'''
def process_command(sentence):
    # Load the English NLP model
    nlp = spacy.load("en_core_web_sm")

    # Process the command using spaCy
    doc = nlp(sentence)

    # Extract lemmatized and stemmed keywords into a set
    query = set()
    for token in doc:
        if token.is_alpha and not token.is_stop:
            query.add(token.lemma_.lower())

    return query #query is a set where all the keywords from user's inputed sentence are present.
'''

def contains_keywords(query, keywords):
    # Check if the command contains any of the specified keywords
    return any(re.search(fr'\b{keyword}\b', query) for keyword in keywords)

def run_as_admin():
    # Check if the script is already running as administrator
    if ctypes.windll.shell32.IsUserAnAdmin():
        return True

    # If not, relaunch the script with administrator privileges
    script = sys.argv[0]
    params = " ".join(sys.argv[1:])
    
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, script, params, 1)



###########################################################################################OPEN CLOSE
def open_url(url):
    command = f"start {url}"
    os.system(command)

def open_application(application_path):
    try:
        os.startfile(application_path)
    except Exception as e:
        print(f"Error opening application: {e}")

def open_application_from_shell(application_name):
    try:
        subprocess.Popen(application_name, shell=True)
    except Exception as e:
        print(f"Error opening {application_name}: {e}")

def extract_open_tabs_urls():
    urls = []
    # Set up Chrome WebDriver
    driver = webdriver.Chrome()
    driver.get("chrome://version/")  # This opens Chrome's version page which lists all open tabs
    elements = driver.find_elements_by_css_selector(".content td")  # Extract elements containing URLs
    for element in elements:
        url = element.text
        if url.startswith("http"):
            urls.append(url)
    driver.quit()  # Close the WebDriver
    return urls


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

def close_tabs(domain):
    driver = webdriver.Chrome()
    try:
        driver.get("chrome://version/")  # Open Chrome's version page which lists all open tabs
        tabs = driver.find_element(".content td")  # Find elements containing URLs
        for tab in tabs:
            url = tab.text
            if url.startswith("http") and domain in url:
                # Close the tab
                driver.execute_script("window.open('');")  # Open a new tab to switch to the YouTube tab
                driver.switch_to.window(driver.window_handles[-1])  # Switch to the newly opened tab
                driver.get(url)  # Navigate to the YouTube URL
                driver.close()  # Close the tab
                driver.switch_to.window(driver.window_handles[0])  # Switch back to the first tab
    except Exception as e:
        print(f"Error occurred while processing Chrome: {e}")
    finally:
        driver.quit()


#############################################################################################SUMMARY related
def preprocess_text(text):
    nlp = spacy.load("en_core_web_sm", disable=["parser", "ner"])
    doc = nlp(text)
    tokens = [token.text.lower() for token in doc if not token.is_stop and token.is_alpha]
    return tokens

def summarize_web_page(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        paragraphs = soup.find_all('p')
        #text = '\n'.join(paragraph.get_text() for paragraph in paragraphs)
        text = ' '.join(map(lambda p: p.text, soup.find_all('p')))
        #print("Extracted text:", text)  # Add this line for debugging
        return summarize_text(text)
    except Exception as e:
        print(f"Error extracting or summarizing web page: {e}")
        return "Error extracting or summarizing web page."


def summarize_pdf(file_path):
    text = ""
    with fitz.open(file_path) as pdf:
        for page in pdf:
            text += page.get_text()
    return summarize_text(text)


def summarize_text(text):
    if not text.strip():
        return "No text to summarize."

    tokens = preprocess_text(text)
    word_freq = Counter(tokens)

    if not word_freq:
        return "No meaningful words found in the text."

    max_freq = max(word_freq.values())
    for word in word_freq.keys():
        word_freq[word] /= max_freq

    sent_scores = {}
    sentences = text.split(".")
    for sent in sentences:
        for word in sent.split():
            if word.lower() in word_freq.keys():
                if len(sent.split()) < 30:
                    if sent not in sent_scores.keys():
                        sent_scores[sent] = word_freq[word.lower()]
                    else:
                        sent_scores[sent] += word_freq[word.lower()]

    summarized_sentences = sorted(sent_scores, key=sent_scores.get, reverse=True)[:5]

    summary = ". ".join(summarized_sentences)
    return summary

def set_alarm(alarm_time, message):
    while True:
        current_time = time.strftime("%H:%M:%S")
        if current_time == alarm_time:
            print(message)
            frequency = 2500  # Set Frequency To 2500 Hertz
            duration = 1000  # Set Duration To 1000 ms == 1 second
            winsound.Beep(frequency, duration)
            break
        time.sleep(1)

def set_timer(duration, message):
    print(f"Timer set for {duration} seconds.")
    time.sleep(duration)
    print(message)
    frequency = 2500  # Set Frequency To 2500 Hertz
    duration = 1000  # Set Duration To 1000 ms == 1 second
    winsound.Beep(frequency, duration)

def stopwatch():
    input("Press Enter to start the stopwatch.")
    start_time = time.time()
    input("Press Enter to stop the stopwatch.")
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Elapsed time: {elapsed_time:.2f} seconds")

def introduction():
    print("Getting started...")
    speak("Welcome, Getting started.")

# ScilenceChecker Function Takes input and removes scilence
# def scilenceChecker():
# 	userSaid = takecommand().lower()
# 	if userSaid == "":
# 		userSaid = "nothing"
# 	elif userSaid == " ":
# 		userSaid = "nothing"
# 	else:
# 		userSaid = takecommand().lower()


def clearLog():
    if platform.system() == "Windows":
        os.system("cls")
    else:
        os.system("clear")


creditsScreen = r'''
 _____ _                 _          _____            _   _     _
|_   _| |__   __ _ _ __ | | _____  |  ___|__  _ __  | | | |___(_)_ __   __ _
  | | |  _ \ / _` |  _ \| |/ / __| | |_ / _ \|  __| | | | / __| |  _ \ / _` |
  | | | | | | (_| | | | |   <\__ \ |  _| (_) | |    | |_| \__ \ | | | | (_| |
  |_| |_| |_|\__,_|_| |_|_|\_\___/ |_|  \___/|_|     \___/|___/_|_| |_|\__, |
                                                                       |___/
'''

clearLog()

welcomeSplashScreen = r'''

|----------------------------------------------------------------------------------------|
|                                                                                        |
|                  \\    	| W O R K               WITH            C O M M A N D    |
|	============\\		| ------------------------------------------------------ |
|	============//          | P Y T H O N           ||||                AI           |
|                  //                                                                    |
|----------------------------------------------------------------------------------------|        
	
'''

Title = r'''
________________________________________________
_________________________________________________\
__________________________________________________\
___        ___   ___        ___     ______  \      \
\  \  __  /  /   \  \  __  /  /    /  ____\  \      \
 \  \/  \/  /     \  \/  \/  /    /  /        \      \
  \   /\   /       \   /\   /     \  \_____   /      /
   \_/  \_/         \_/  \_/       \______/  /      /
____________________________________________/      /
                                                  /
   P Y T H O N        |||            A I         /
________________________________________________/
'''



############################################################################################ RANDOM def
def disable_startup_programs():
    # Add code to disable startup programs
    startup_folder = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    for filename in os.listdir(startup_folder):
        filepath = os.path.join(startup_folder, filename)
        os.remove(filepath)

def clean_temp_files():
    subprocess.run(["del", "/Q", "/S", os.path.join(os.getenv("TEMP"), "*")], shell=True)

def optimize_power_settings():
    subprocess.run(["powercfg", "/S", "SCHEME_MIN"])

def manage_services():
    subprocess.run(["sc", "config", "SERVICE_NAME", "start=auto"], shell=True)
    subprocess.run(["sc", "start", "SERVICE_NAME"], shell=True)

def schedule_disk_defragmentation():
    subprocess.run(["defrag", "C:", "/C"], shell=True)

def monitor_system_resources():
    subprocess.run(["tasklist"])

def automate_updates():
    subprocess.run(["sconfig", "1"])

def manage_virtual_memory():
    subprocess.run(["wmic", "pagefileset", "where", "name='C:\\pagefile.sys'", "set", "InitialSize=4096,MaximumSize=8192"])

def optimize_visual_effects():
    subprocess.run(["SystemPropertiesPerformance"])

def clean_registry():
    subprocess.run(["reg", "delete", "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", "/f"], shell=True)

def open_file_explorer_with_search(query):
    explorer_cmd = f'explorer "search-ms:displayname=Search%20Results%20in%20Desktop&crumb=System.Generic.String%3A{search_query}"'
    subprocess.Popen(explorer_cmd, shell=True)


class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'

clearLog()
print(f"{bcolors.OKCYAN + welcomeSplashScreen + bcolors.ENDC}")

if __name__ == '__main__':

    #checking for admin priviledges
    if run_as_admin():
        print("Running with administrator privileges.")
    else:
        print("Script relaunched with administrator privileges.")

    introduction()
    wishMe()
    looper = 5
    system_counter = 0
    lib_counter = 0
    
    while looper != 50:

        #sentence = str(takecommand())
        sentence = str(input("Enter: "))

        #process_command(sentence) #NLP
        
        # Load the English NLP model
        nlp = spacy.load("en_core_web_sm")
        custom_stop_words = {'what','who','when','where','how','to','you','yourself','on','off','take','next','show','make','create'}
        for word in custom_stop_words:
            nlp.vocab[word].is_stop = False

        # Process the command using spaCy
        doc = nlp(sentence)

        # Extract lemmatized and stemmed keywords into a set
        query = set()
        for token in doc:
            if token.is_alpha and not token.is_stop:
                query.add(token.lemma_.lower())

        # Extracting Nouns and Proper Nouns
        for word in custom_stop_words: #disabling stop words
            nlp.vocab[word].is_stop = False
        app_name = ""
        for token in doc:
            if token.pos_ in ['NOUN', 'PROPN'] and not token.is_stop:
                app_name += token.text + " "        
        
        # Probablity of Commands
        if ('who' in query and 'you' in query) or ('tell' in query and ('you' in query or 'yourself' in query)):
            speak("Python here! Simply command me to do something"
                        "I'll try to prepare it for you.")
            print(">> Add elements to the html file\n>> Open and Close an Application\n>> Ask for news headlines or Browse online\n>> Read a PDF")

        elif "clear" in query and "log" in query:
            clearLog()
        
        elif "pause" in query or "hold" in query or "wait" in query:
            a = input("press any key to continue...")


##################################################################################################################          HTML
        
        
        elif "add" in query and "address" in query:
            speak("Adding your address, Tell me your address")
            address = input("Please type your address: ")
            finalTag = (f"<address>{address}</address>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("Address added!\n")

        elif "add" in query and "anchor" in query:
            speak("adding your Anchor")
            print("Tell me the your URL")
            URL = input("Enter/Paste your URL: ")
            URLName = input("Enter tittle of your URL: ")
            finalTag = (f'<a href="{URL}">{URLName}</a>\n')
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("Attribute/URL added!\n")

        elif "add" in query and "audio" in query:
            speak("adding your Audio")
            print("Tell me path of your audio.")
            audioPath = input("Enter/Paste your path: ")
            finalTag = (
                f'<audio controls autoplay><source src="{audioPath}" type="audio/mpeg"></audio>\n')
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("Audio added!\n")

        elif "add" in query and "blockquote" in query:
            speak("adding your blockquote, Tell me your blockquote content")
            print("Say your blockquote content: ")
            internalquery = str(takecommand())
            blockquote = internalquery
            finalTag = (f"<blockquote>{blockquote}</blockquote>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("Comment added!\n")

        elif "add" in query and "br" in query:
            speak("adding your br tag")
            finalTag = (f"<br>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("br added!\n")

        elif "add" in query and "button" in query:
            speak("adding your Button")
            buttonName = input("Enter your Button name: ")
            finalTag = (f'<button type="button">{buttonName}</button>\n')
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("Attribute/URL added!\n")

        elif "add" in query and "comment" in query:
            speak("adding your comment, Tell me your comment")
            print("Say your comment: ")
            internalquery = str(takecommand())
            comment = internalquery
            finalTag = (f"<!--{comment}-->\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("Comment added!\n")

        elif "add" in query and "heading" in query:
            speak("adding your heading, Tell me your heading Size")
            print("Say your heading size: ")
            internalquery = str(takecommand())
            sizeOfHeadingTag = internalquery
            # sizeOfHeadingTag = input("Enter size of Heading Tag(1, 2 ,3 ,4...): ")
            # contentOfHeading = input("Enter content of Heading: ")
            speak("Adding your heading, Tell me your heading content")
            print("Say your heading content: ")
            internalquery = str(takecommand())
            contentOfHeading = internalquery
            finalTag = (
                f"<h{sizeOfHeadingTag}>{contentOfHeading}</h{sizeOfHeadingTag}>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("Heading added!\n")

        elif "add" in query and "hr" in query:
            speak("adding your hr tag")
            finalTag = (f"<hr>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("hr added!\n")

        elif "add" in query and "iframe" in query:
            speak("adding your iframe tag")
            width = input("Enter Width of your ifame in pixels: ")
            height = input("Enter height of your ifame in pixels: ")
            url = input("paste/enter url for iframe: ")
            finalTag = (
                f"<iframe src=\"{url}\" width=\"{width}px\" height=\"{height}px\"></iframe>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("iframe added!\n")

        elif "add" in query and "image" in query:
            speak("adding your image tag")
            image_name = input("Enter name of your image in pixels: ")
            width = input("Enter Width of your image in pixels: ")
            height = input("Enter height of your image in pixels: ")
            url = input("paste/enter url for image with extension: ")
            finalTag = (
                f"<img src=\"{url}\" width=\"{width}px\" height=\"{height}px\" alt=\"{image_name}\">\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("image added!\n")

        # elif "add" in query and "list" in query:
        # 	speak("adding your list tag (Unordered list)")
        # 	noofitems = input("Enter number of items in your list: ")
        # 	finalTag = (f"<ul>\n")
        # 	f = open("index.html", "a")
        # 	f.write(finalTag)
        # 	for i in noofitems:
        # 		user_item_input = input("Enter your item: ")
        # 		finalTag = (f"<li>{list_items}</li>\n")
        # 		f = open("index.html", "a")
        # 		f.write(finalTag)
        # 		f.close()
        # 	finalTag = (f"/ul>\n")
        # 		f = open("index.html", "a")
        # 		f.write(finalTag)
        # 	clearLog()
        # 	print("list added!\n")

        elif "add" in query and "paragraph" in query:
            speak("adding your paragraph")
            print("\nNo Speech input for paragraph because paragraphs can be too long to speak\nDo not press enter for new line type \"<br>\" to change line")
            content = input("Start typing/paste your paragraph here: \n")
            finalTag = (f"<p>{content}</p>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("paragraph added!\n")

        elif "add" in query and "video" in query:
            speak("adding your video tag")
            width = input("Enter Width of your video in pixels: ")
            height = input("Enter height of your video in pixels: ")
            url = input(
                "paste/enter url for video without extension (mp4 only): ")
            finalTag = (
                f"<video width=\"{width}\" height=\"{height}\" controls><source src=\"{url}.mp4\" type=\"video/mp4\"></video>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            print("video added!\n")

        elif "complete" in query and "website" in query:
            speak("completing your website")
            finalTag = (f"</body>\n</html>\n")
            f = open("index.html", "a")
            f.write(finalTag)
            f.close()
            clearLog()
            speak("Website Completed!, Thanks for using VOCALGRAMMER\n")
            print("Website Completed!\n")


##################################################################################################################      Open

        elif 'time' in query:  #done1
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")

        elif 'date' in query:  #done1
            strDate = datetime.datetime.today().strftime('%Y-%m-%d')
            print(strDate)
            speak(f"The date is {strDate}")

        elif 'localhost' in query:  #done1
            webbrowser.open("localhost")

        elif ("take" in query and "picture" in query) or "capture" in query: #done3
            try:
                subprocess.run('start microsoft.windows.camera:', shell=True)
            except Exception as e:
                print(f"Error opening Camera: {e}")

        elif ("open" in query or "launch" in query or "start" in query) and ('how' not in query or 'who' not in query or 'what' not in query or 'where' not in query or 'when' not in query):
            system_counter +=1

            if "file" in query or "Explorer" in query or 'folder' in query:
                open_application_from_shell("explorer")
                
            elif ("recycle" in query and "bin" in query) or "bin" in query:
                open_application_from_shell("explorer shell:RecycleBinFolder")
                
            elif "calculator" in query:
                open_application_from_shell("calc")
                
            elif "calendar" in query:
                open_application_from_shell("start outlookcal:")
                
            elif "clock" in query:
                open_application_from_shell("start ms-clock:")
                
            elif "store" in query:
                open_application_from_shell("start ms-windows-store:")
                
            elif "chrome" in query:
                open_application_from_shell("start chrome")
                
            elif "Maps" in query:
                open_application_from_shell("start bingmaps:")
                
            elif ("Microsoft" in query and "Edge" in query) or "edge" in query:
                open_application_from_shell("msedge")
                
            elif ("Microsoft" in query and "Teams" in query) or "teams" in query:
                open_application_from_shell("start msteams")
                
            elif "Outlook" in query:
                open_application_from_shell("start outlook")
                
            elif "Paint" in query:
                open_application_from_shell("mspaint")
                
            elif "Photos" in query:
                open_application_from_shell("start ms-photos:")
                
            elif "skype" in query:
                open_application_from_shell("skype")
                
            elif "Terminal" in query:
                open_application_from_shell("wt")
                
            elif ("Windows" in query and "Security" in query) or "Security" in query:
                open_application_from_shell("start ms-settings:windowsdefender")
                
            elif "Word" in query:
                open_application_from_shell("start winword")
                
            elif "Access" in query:
                open_application_from_shell("start msaccess")

            elif ("Power" in query and "Point" in query) or "ppt" in query:
                open_application_from_shell("start powerpnt")
                
            elif "Excel" in query:
                open_application_from_shell("start excel")

            elif ("sound" in query and "recorder" in query):
                os.startfile("sound recorder")
                
            elif "sticky" in query and "notes" in query:
                open_application_from_shell("sticky notes")
            
            elif "task" in query and "manager" in query: #done2
                try:
                    command = "taskmgr"
                    subprocess.Popen(command, shell=True)
                except Exception as e:
                    print(f"Error opening Task Manager: {e}")

            elif "notepad" in query: #done2
                try:
                    command = "C:\\WINDOWS\\system32\\notepad.exe"
                    os.system(command)
                except Exception as e:
                    print(f"Error opening Notepad: {e}")
            elif "camera" in query: #done2
                try:
                    subprocess.run('start microsoft.windows.camera:', shell=True)
                except Exception as e:
                    print(f"Error opening Camera: {e}")
            elif "browser" in query: #done3
                try:
                    command = "start https://www.google.com"
                    os.system(command)
                except Exception as e:
                    print(f"Error opening Browser: {e}")
            elif "command" in query and "prompt" in query: #done4
                try:
                    command = "cmd"
                    subprocess.Popen(command, shell=True)
                except Exception as e:
                    print(f"Error opening Command Prompt: {e}")
            elif "youtube" in query: #done4
                try:
                    open_url("www.youtube.com")
                except Exception as e:
                    print(f"Error opening YouTube: {e}")
            elif "email" in query or "mail" in query or "gmail" in query: #done4
                try:
                    open_url("https://mail.google.com/")
                except Exception as e:
                    print(f"Error opening Email: {e}")
            elif "google" in query: #done4
                try:
                    open_url("www.google.com")
                except Exception as e:
                    print(f"Error opening Google: {e}")
            elif "stack" in query and "overflow" in query: #done4
                try:
                    open_url("www.stackoverflow.com")
                except Exception as e:
                    print(f"Error opening Stack Overflow: {e}")
            elif ("visual" in query and "studio" in query) or ("VS" in query or "code" in query): #done5
                try:
                    open_application("C:\\Users\\User1" + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\"
                                     "Programs\\Visual Studio Code\\Visual Studio Code")
                except Exception as e:
                    print(f"Error opening Visual Studio Code: {e}")
            elif "mozilla" in query or "firefox" in query:
                try:
                    open_application("C:\\Program Files\\Mozilla Firefox\\firefox.exe")
                except FileNotFoundError:
                    search_query = "download Mozilla firefox"
                    search_url = f"https://www.google.com/search?q={search_query}"
                    webbrowser.open(search_url)
                    speak(f"Application '{application_name}' not found.")
                except Exception as e:
                    print(f"Error opening Mozilla Firefox: {e}")
            elif "whatsapp" in query:
                try:
                    open_application("C:\\Users\\User1" + "\\AppData\\Local\\WhatsApp\\WhatsApp.exe")
                except Exception as e:
                    print(f"Error opening WhatsApp: {e}")
            elif "vlc" in query:
                try:
                    open_application("C:\\Program Files\\VideoLAN\\VLC\\vlc.exe")
                except Exception as e:
                    print(f"Error opening VLC Media Player: {e}")

            else:
                # Check if the application exists in C:\WINDOWS\system32\
                app_path = os.path.join("C:\\WINDOWS\\system32\\", app_name + ".exe")
                if os.path.exists(app_path):
                    try:
                        os.startfile(app_path)                    
                    except Exception as e:
                        speak_error(f"Sorry, I couldn't open {app_name}.")
                else:
                    search_query = f"download {app_name}"
                    search_url = f"https://www.google.com/search?q={search_query}"
                    webbrowser.open(search_url)
                    speak(f"Application '{app_name}' not found in your system.")


##################################################################################################################      Close


        elif ('close' in query or "eliminate" in query or "terminate" in query) and ('how' not in query or 'who' not in query or 'what' not in query or 'where' not in query or 'when' not in query):
            
            if ("recycle" in query and "bin" in query) or "bin" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "explorer.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing File Explorer: {e}")
            elif "calculator" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "calc.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Calculator: {e}")
            elif "calendar" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "outlookcal:"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Calendar: {e}")
            elif "clock" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "ms-clock.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Clock: {e}")
            elif "store" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "ms-windows-store.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Microsoft Store: {e}")
            elif "chrome" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "chrome.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Google Chrome: {e}")
            elif "Maps" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "bingmaps:"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Maps: {e}")
            elif ("Microsoft" in query and "Edge" in query) or "edge" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "msedge.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Microsoft Edge: {e}")
            elif ("Microsoft" in query and "Teams" in query) or "teams" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "msteams.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Microsoft Teams: {e}")
            elif "Outlook" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "outlook.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Outlook: {e}")
            elif "Paint" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "mspaint.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Paint: {e}")
            elif "Photos" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "Microsoft.Photos.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Photos: {e}")
            elif "skype" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "skype.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Skype: {e}")
            elif "Terminal" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "wt.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Terminal: {e}")
            elif ("Windows" in query and "Security" in query) or "Security" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "WindowsDefenderSecurityCenter.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Windows Security: {e}")
            elif "Word" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "winword.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Word: {e}")
            elif "Access" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "msaccess.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Access: {e}")
            elif ("Power" in query and "Point" in query) or "ppt" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "powerpnt.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Power Point: {e}")
            elif "Excel" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "excel.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Excel: {e}")

            elif ("sound" in query and "recorder" in query):
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "VoiceRecorder.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Sound Recorder: {e}")
                    
            elif "sticky" in query and "note" in query:
                try:
                    subprocess.run(["TASKKILL", "/F", "/IM", "Microsoft.Notes.exe"], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Error closing Sticky Notes: {e}")
            
            elif 'command' in query and 'prompt' in query:
                try:
                    os.system("TASKKILL /F /IM cmd.exe")
                except Exception as e:
                    print(f"Error closing Command Prompt: {e}")
            elif 'task' in query and 'manager' in query: #done2
                try:
                    os.system("TASKKILL /F /IM Taskmgr.exe")
                except Exception as e:
                    print(f"Error closing task manager: {e}")
            elif 'camera' in query: #done2 #done3
                try:
                    subprocess.run('Taskkill /IM WindowsCamera.exe /F', shell=True)
                except Exception as e:
                    print(f"Error closing Camera: {e}")
            elif 'youtube' in query:
                close_tabs("youtube")
            elif 'firefox' in query:
                try:
                    os.system("TASKKILL /F /IM firefox.exe")
                except Exception as e:
                    print(f"Error closing Firefox: {e}")
            elif 'visual' in query and 'studio' in query and 'code' in query:
                try:
                    os.system("TASKKILL /F /IM Code.exe")
                except Exception as e:
                    print(f"Error closing Visual Studio Code: {e}")
            elif 'notepad' in query: #done2
                try:
                    os.system("TASKKILL /F /IM notepad.exe")
                except Exception as e:
                    print(f"Error closing Notepad: {e}")
            elif 'chrome' in query:
                try:
                    os.system("TASKKILL /F /IM chrome.exe")
                except Exception as e:
                    print(f"Error closing Chrome: {e}")
            elif 'whatsapp' in query:
                try:
                    os.system("TASKKILL /F /IM WhatsApp.exe")
                except Exception as e:
                    print(f"Error closing WhatsApp: {e}")
            elif 'vlc' in query:
                try:
                    os.system("TASKKILL /F /IM vlc.exe")
                except Exception as e:
                    print(f"Error closing VLC: {e}")
            elif 'spotify' in query:
                try:
                    os.system("TASKKILL /F /IM Spotify.exe")
                except Exception as e:
                    print(f"Error closing Spotify: {e}")
            else:
                print("Application not recognized.")


##################################################################################################################    Search Online


        #elif "camera" or "take a photo" in query:
         #   ec.capture(0,"robo camera","img.jpg")

        elif 'find' in query and 'image' in query:
            speak("Please provide the desired keywords to search related Images")
            txt=takecommand()
            #txt=input("image of: ")
            speak("Providing Images of "+txt)
            response = requests.get("https://source.unsplash.com/random?{0}".format(txt))
            try:
                with open('sample_image.jpg', 'wb') as file:
                    file.write(response.content)
                img = Image.open(r"sample_image.jpg")
                img.show()
            except Exception as e:
                print(f"Error: {e}")
                speak("Sorry, I couldn't find any images for your query.")

        elif 'send' in query and ('email' in query or 'mail' in query):

            def send_mail(receiver, subject, message):
                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    _email = input('Enter your email address: ')
                    _password = getpass.getpass('Enter your password (it will not be displayed): ')
                    server.login(_email, _password)

                    email = EmailMessage()
                    email['From'] = _email
                    email['To'] = receiver
                    email['Subject'] = subject
                    email.set_content(message)

                    server.send_message(email)
                    print("Email sent successfully!")
                except Exception as e:
                    print(f'Error sending email: {e}')
                finally:
                    server.quit()

            try:
                receiver = input('Enter receiver\'s email address: ')
                subject = input('Enter the subject of your email: ')
                message = input('Enter the content of your email: ')

                send_mail(receiver, subject, message)
            except Exception as e:
                print(f'Error: {e}')

        elif 'play' in query or 'youtube' in query: #used to play specific video on youtube
    
            if 'play' in query and 'youtube' in query:
                cmd_info = sentence.replace('play', '')
                cmd_info = cmd_info.replace('youtube', '')
            elif 'play' in query:
                cmd_info = sentence.replace('play', '')
            elif 'youtube' in query:
                cmd_info = sentence.replace('youtube', '')
            elif 'in youtube' in query:
                cmd_info = sentence.replace('youtube', '')
            speak(f'Playing {cmd_info} ')
            print(cmd_info)
            pywhatkit.playonyt(cmd_info)

        elif 'search' in query and ('file' in query or 'folder' in query):
            speak('Input the file or folder name to search for')
            search_query = input('Name the file or folder to search for: ')  # Change this to your search query
            open_file_explorer_with_search(search_query)

        elif 'search' in query:
            cmd_info = sentence.replace('search', '')
            pywhatkit.search(cmd_info)


        elif ('price' in query and 'of' in query) or 'price' in query:
            cmd_info = sentence.replace('price of', '')  # Replace 'price of' instead of just 'price'
            cmd_info = ' '.join([word for word in cmd_info.split() if word.lower() not in custom_stop_words])  # Remove custom stop words
            doc = nlp(cmd_info)  # Process the modified command info with spaCy
            cmd_info = ' '.join(token.text for token in doc if not token.is_stop)  # Remove stop words
            to_search = "https://www.amazon.in/s?k=" + cmd_info  # Construct the search URL
            webbrowser.open(to_search)
            try:
                response = requests.get(to_search)
                soup = BeautifulSoup(response.text, 'html.parser')
                price_tag = soup.find('span', attrs={'class': 'a-price-whole'})
                price_symbol = soup.find('span', attrs={'class': 'a-price-symbol'})
                if price_tag and price_symbol:
                    price = price_tag.text
                    symbol = price_symbol.text
                    print(f'price {price} {symbol}')
                    speak(f'The price of {cmd_info} is {price} {symbol}')
                else:
                    print('Price not found')
                    speak('Price not found')
            except Exception as e:
                print(f'Error scraping Amazon: {e}')

        elif 'poem' in query:
            print('Poem of which Poet you want to listen?')
            #auth = takecommand()
            auth = input()
            poem = poetpy.get_poetry('author', auth, 'title,linecount') 
            poems = poetpy.get_poetry('author', auth, 'lines')  

            poem_len = len(poem)
            # print(poem_len)
            poem_no = random.randint(0, poem_len - 1)
            print("Title- ", poem[poem_no]['title'])
            speak(f"Title- {poem[poem_no]['title']}")
            print("No. of lines-", poem[poem_no]['linecount'])
            speak(f"No. of lines- {poem[poem_no]['linecount']}")
            poem_str = '\n'
            print("Poem-\n", poem_str.join(poems[poem_no]['lines']))
            #speak(f"Poem-\n {poem_str.join(poems[poem_no]['lines'])}")

        elif 'summarize' in query or 'summary' in query:
            if ('web' in query and 'page' in query) or 'website' in query:
                url = input("Enter the URL of the web page: ")
                summary = summarize_web_page(url)
            elif 'pdf' in query:
                file_path = input("Enter the path to the local PDF file: ")
                summary = summarize_pdf(file_path)
            else:
                text = input("Enter the text to summarize: ")
                summary = summarize_text(text)

            print("Summary:")
            print(summary)
            time.sleep(5)
            
        elif ('resume' in query or 'pause' in query or 'unpause' in query or 'start' in query) and ('music' in query or 'song' in query):
            try:
                pyautogui.hotkey('fn', 'F4')
            except Exception as e:
                print(f"Error: {e}")

        elif 'previous' in query and ('track' in query or 'song' in query or 'music' in query):
            pyautogui.press("prevtrack")

        elif 'next' in query and ('track' in query or 'song' in query or 'music' in query):
            pyautogui.press("nexttrack")

        elif 'convert' in query and 'currency' in query:
            try:
                curr_list = {
                    'taka': 'BDT', 'USD': 'USD', 'American dollar': 'USD', 'dinar': 'BHD',
                    'rupee': 'INR', 'afghani': 'AFN', 'real': 'BRL',
                    'yen': 'JPY', 'peso': 'ARS', 'pound': 'EGP', 'rial': 'OMR',
                    'lek': 'ALL', 'kwanza': 'AOA', 'manat': 'AZN', 'franc': 'CHF',
                    'rupees': 'INR', 'dinars': 'BHD', 'euro': 'EUR', 'pounds': 'GBP', 'dirham': 'AED',
                    'ringgit': 'MYR', 'won': 'KRW', 'yuan': 'CNY', 'koruna': 'CZK', 'zloty': 'PLN',
                    'dollar': 'USD', 'rubles': 'RUB', 'lira': 'TRY', 'kwacha': 'MWK', 'riyal': 'SAR',
                    'krona': 'SEK', 'shekel': 'ILS', 'rand': 'ZAR', 'baht': 'THB', 'dram': 'AMD',
                    'forint': 'HUF', 'krone': 'NOK', 'leu': 'RON', 'colon': 'CRC', 'ruble': 'RUB',
                    'kuna': 'HRK', 'lempira': 'HNL', 'cordoba': 'NIO', 'quetzal': 'GTQ', 'sol': 'PEN',
                    'soum': 'UZS', 'tenge': 'KZT', 'togrog': 'MNT', 'uf': 'CLF', 'vatu': 'VUV',
                    'won': 'KPW', 'yen': 'JPY', 'yuan': 'CNY', 'zloty': 'PLN',
                    'BD': 'BDT', 'US Dollar': 'USD', 'Euro': 'EUR', 'Pound Sterling': 'GBP',
                    'JPY': 'JPY', 'CNY': 'CNY', 'INR': 'INR', 'RUB': 'RUB', 'KRW': 'KRW'
                }

                cur = CurrencyRates()
                # print(cur.get_rate('USD', 'INR'))
                speak('From which currency you want to convert?')
                from_cur = takecommand()
                src_cur = curr_list[from_cur.lower()]
                speak('To which currency you want to convert?')
                to_cur = takecommand()
                dest_cur = curr_list[to_cur.lower()]
                speak('Tell me the value of currency you want to convert.')
                val_cur = float(input('amount: '))
                # print(val_cur)
                print(cur.convert(src_cur, dest_cur, val_cur))
                        
            except Exception as e:
                print("Couldn't get what you have said, Can you say it again??")

        elif 'covid-19' in query or 'corona' in query:
            speak('For which region u want to see the Covid-19 cases. '
                        'Overall cases in the world or any specific country?')
            c_query = takecommand()
            if 'overall' in c_query or 'world' in c_query or 'total' in c_query or 'worldwide' in c_query:
                def world_cases():
                    try:
                        url = 'https://www.worldometers.info/coronavirus/'
                        info_html = requests.get(url)
                        info = bs4.BeautifulSoup(info_html.text, 'lxml')
                        info2 = info.find('div', class_='content-inner')
                        new_info = info2.findAll('div', id='maincounter-wrap')
                        # print(new_info)
                        print('Worldwide Covid-19 information--')
                        speak('Worldwide Covid-19 information--')

                        for i in new_info:
                            head = i.find('h1', class_=None).get_text()
                            counting = i.find('span', class_=None).get_text()
                            print(head, "", counting)
                            speak(f'{head}: {counting}')

                    except Exception as e:
                        pass


                world_cases()

            else:
                def country_cases():
                    try:
                        speak('Tell me the country name.')
                        c_name = takecommand()
                        c_url = f'https://www.worldometers.info/coronavirus/country/{c_name}/'
                        data_html = requests.get(c_url)
                        c_data = bs4.BeautifulSoup(data_html.text, 'lxml')
                        new_data = c_data.find('div', class_='content-inner').findAll('div', id='maincounter-wrap')
                        # print(new_data)
                        print(f'Covid-19 information for {c_name}--')
                        speak(f'Covid-19 information for {c_name}')

                        for j in new_data:
                            c_head = j.find('h1', class_=None).get_text()
                            c_counting = j.find('span', class_=None).get_text()
                            print(c_head, "", c_counting)
                            speak(f'{c_head}: {c_counting}')

                    except Exception as e:
                        pass


                country_cases()

        elif 'weather' in query or 'temperature' in query:
            try:
                speak("Tell me the city name.")
                city = takecommand()
                api = "http://api.openweathermap.org/data/2.5/weather?q=" + city + "&appid=eea37893e6d01d234eca31616e48c631"
                w_data = requests.get(api).json()
                weather = w_data['weather'][0]['main']
                temp = int(w_data['main']['temp'] - 273.15)
                temp_min = int(w_data['main']['temp_min'] - 273.15)
                temp_max = int(w_data['main']['temp_max'] - 273.15)
                pressure = w_data['main']['pressure']
                humidity = w_data['main']['humidity']
                visibility = w_data['visibility']
                wind = w_data['wind']['speed']
                sunrise = time.strftime("%H:%M:%S", time.gmtime(w_data['sys']['sunrise'] + 19800))
                sunset = time.strftime("%H:%M:%S", time.gmtime(w_data['sys']['sunset'] + 19800))

                all_data1 = f"Condition: {weather} \nTemperature: {str(temp)}°C\n"
                all_data2 = f"Minimum Temperature: {str(temp_min)}°C \nMaximum Temperature: {str(temp_max)}°C \n" \
                            f"Pressure: {str(pressure)} millibar \nHumidity: {str(humidity)}% \n\n" \
                            f"Visibility: {str(visibility)} metres \nWind: {str(wind)} km/hr \nSunrise: {sunrise}  " \
                            f"\nSunset: {sunset}"
                speak(f"Gathering the weather information of {city}...")
                print(f"Gathering the weather information of {city}...")
                print(all_data1)
                speak(all_data1)
                print(all_data2)
                speak(all_data2)

            except Exception as e:
                pass

        elif 'month' in query:
            current_date = datetime.datetime.now()
            if 'next' in query:
                next_month = current_date.replace(day=1, month=current_date.month + 1)
                month = next_month.strftime("%B")
            elif 'previous' in query:
                previous_month = current_date.replace(day=1, month=current_date.month - 1)
                month = previous_month.strftime("%B")
            else:
                month = datetime.datetime.now().strftime("%B")
            speak(month)

        elif 'day' in query:
            current_day = datetime.datetime.now().strftime("%A")
            days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            current_index = days_of_week.index(current_day)
            
            if 'next' in query or 'tomorrow' in query:
                next_index = (current_index + 1) % 7
                day = days_of_week[next_index]
            elif 'previous' in query or 'yesterday' in query:
                previous_index = (current_index - 1) % 7
                day = days_of_week[previous_index]
            else:
                day = current_day
            speak(day)
            
        elif 'calculate' in query:
            def calculate(query):
                try:
                    app_id = "JUGV8R-RXJ4RP7HAG"
                    client = wolframalpha.Client(app_id)
                    indx = query.lower().split().index('calculate')
                    query = query.split()[indx + 1:]
                    res = client.query(' '.join(query))
                    answer = next(res.results).text
                    print("The answer is " + answer)
                    speak("The answer is " + answer)
                except Exception as e:
                    print("Couldn't get what you have said. Can you say it again?")

            calculate(sentence)

        elif 'quote' in query or 'motivate' in query:
            speak("Tell me the author or person name.")
            q_author = takecommand()
            quotes = quote(q_author)
            quote_no = random.randint(1, len(quotes))
            # print(len(quotes))
            # print(quotes)
            print("Author: ", quotes[quote_no]['author'])
            print("-->", quotes[quote_no]['quote'])
            speak(f"Author: {quotes[quote_no]['author']}")
            speak(f"He said {quotes[quote_no]['quote']}")

        elif 'what' in query or 'who' in query or 'where' in query or 'when' in query:             
            client = wolframalpha.Client("JUGV8R-RXJ4RP7HAG")
            res = client.query(sentence)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)

            except StopIteration:
                search_url = f"https://www.google.com/search?q={sentence}"
                webbrowser.open(search_url)
                speak("Don't have an answer for it. This is what I found in Google")

        elif ('write' in query or 'make' in query or 'take' in query) and 'note' in query:
            speak("What should I write??")
            note = ''
            note = takecommand()
            file = open('Notes.txt', 'a')
            speak("Should I include the date and time??")
            n_conf = takecommand()
            if 'yes' in n_conf or 'sure' in n_conf:
                str_time = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(str_time)
                file.write(" --> ")
                file.write(note)
                speak("Point noted successfully.")
            else:
                file.write("\n")
                file.write(note)
                speak("Point noted successfully.")

        elif ('show' in query or 'read' in query) and 'note' in query:
            speak("Reading Notes")
            file = open("Notes.txt", "r")
            data_note = file.readlines()
            # for points in data_note:
            print(data_note)
            speak(data_note)

        elif 'distance' in query:
            geocoder = Nominatim(user_agent="main")
            speak("Tell me the first city name??")
            location1 = takecommand()
            speak("Tell me the second city name??")
            location2 = takecommand()

            coordinates1 = geocoder.geocode(location1)
            coordinates2 = geocoder.geocode(location2)

            lat1, long1 = coordinates1.latitude, coordinates1.longitude
            lat2, long2 = coordinates2.latitude, coordinates2.longitude

            place1 = (lat1, long1)
            place2 = (lat2, long2)

            distance_places = distance.distance(place1, place2)

            print(f"The distance between {location1} and {location2} is {distance_places}.")
            speak(f"The distance between {location1} and {location2} is {distance_places}")


##################################################################################################################     System

        elif 'set' in query and 'alarm' in query:
            speak("Enter alarm time.")
            alarm_time = input("Enter the alarm time in HH:MM:SS format: ")
            message = input("Enter your message: ")
            set_alarm(alarm_time, message)

        elif 'set' in query and 'timer' in query:
            speak("Enter duration of timer.")
            duration = int(input("Enter the duration of the timer in seconds: "))
            message = input("Enter your message: ")
            set_timer(duration, message)

        elif 'stopwatch' in query:
            stopwatch()

        elif ('empty' in query or 'clear' in query) and ('recycle' in query and 'bin' in query):
            try:
                winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                print("Recycle Bin is cleaned successfully.")
                speak("Recycle Bin is cleaned successfully.")

            except Exception as e:
                print("Recycle bin is already Empty.")
                speak("Recycle bin is already Empty.")

        elif 'shut' in query and 'down' in query:
            print("Do you want to shutdown you system?")
            speak("Do you want to shutdown you system?")
            cmd = takecommand()
            if 'no' in cmd:
                continue
            else:
                os.system("shutdown /s /t 1")

        elif 'restart' in query:
            print("Do you want to restart your system?")
            speak("Do you want to restart your system?")
            cmd = takecommand()
            if 'no' in cmd:
                continue
            else:
                os.system("shutdown /r /t 1")

        elif 'log' in query and 'out' in query:
            print("Do you want to logout from your system?")
            speak("Do you want to logout from your system?")
            cmd = takecommand()
            if 'no' in cmd:
                continue
            else:
                os.system("shutdown -l")

        elif 'check' in query and 'update' in query: #done
            try:
                update_session = comtypes.client.CreateObject("Microsoft.Update.Session")
                update_searcher = update_session.CreateUpdateSearcher()
                search_result = update_searcher.Search("IsInstalled=0 and Type='Software'")
                for update in search_result.Updates:
                    print(f"Title: {update.Title}")
                    print(f"Description: {update.Description}")
                    try:
                        print(f"Download size: {update.DownloadContentsLength / 1024 / 1024:.2f} MB")
                    except AttributeError:
                        print("Download size: Not available")
                    print("="*10)

            except Exception as e:
                speak("Error")
                print(f"Error checking for updates: {e}")

        elif 'defrag' in query or 'defragmentation' in query: #done
            speak("Provide the Drive you want to defrag")
            drive_letter= input("enter the letter of the Drive you want to defrag: ")+":"

            try:
                partitions = psutil.disk_partitions()
                for partition in partitions:
                    if partition.device.startswith(drive_letter):
                        drive_info = psutil.disk_usage(partition.mountpoint)
                        drive_type = "SSD" if drive_info.percent < 10 else "HDD"
                        print(drive_info , drive_type)

            except Exception as e:
                print(f"Error: {e}")

            if drive_type == "HDD":
                try:
                    # Run the defrag command for the specified drive
                    subprocess.run(["defrag", drive_letter + ":"], check=True)
                    speak(f"Defragmentation completed for drive {drive_letter}.")
                except subprocess.CalledProcessError as e:
                    speak("Failed")
                    print(f"Failed to defragment drive {drive_letter}: {e}")
            elif drive_type == "SSD":
                subprocess.run(["optimize", "/O"], check=True) #usually optimisation is automatic in windows
            else:
                speak("error")
                write("Drive might not exist")

        elif 'troubleshooting' in query or 'troubleshoot' in query: #done
            try:
                os.system("start ms-settings:System/Troubleshoot/Other troubleshooters")
                print("troubleshooter opened successfully.")
            except subprocess.CalledProcessError as e:
                print(f"Failed to open Windows troubleshooters: {e}")

##################################################################################################################      Lib based

        elif ('make' in query or 'create' in query) and ('image' in query or 'picture' in query):
            speak("tell me how do you want the image to be") #done
            #prompt = takecommand()
            prompt = input('prompt: ')
            os.environ["REPLICATE_API_TOKEN"] = "r8_LDEu1Ff3lAH56Pp1Xl2SS6UXqIxnVgV4VVPlN"

            input_data = {
                "image": "https://pngimg.com/uploads/square/square_PNG19.png",
                "seed": 20,
                "prompt": prompt,
                "structure": "scribble",
                "image_resolution": 512
            }

            try:
                output = replicate.run(
                    "rossjillian/controlnet:795433b19458d0f4fa172a7ccf93178d2adb1cb8ab2ad6c8fdc33fdbcd49f477",
                    input=input_data
                )

                print(output)
                image_url = output[0]
                image = Image.open(image_url)
                
                image.show()
            except replicate.exceptions.ReplicateError as e:
                print(f"ReplicateError: {e}")
            
        
        elif 'screenshot' in query: #done
            if not os.path.exists('Screenshots'):
                os.makedirs('Screenshots')
            now = datetime.datetime.now()
            date_time = now.strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f'Screenshots/{date_time}.png'
            pyautogui.screenshot(file_name)
            print("Screenshot taken successfully.")

        elif 'volume' in query and('raise' in query or 'up' in query or 'high' in query):
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")

        elif 'volume' in query and('lower' in query or 'down' in query):
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")

        elif ('unmute' in query or 'mute' in query) and 'volume' in query:
            pyautogui.press("volumemute")

        elif 'joke' in query:
            joke = pyjokes.get_joke()
            print(joke)
            speak(joke)

        elif 'internet' in query and 'speed' in query:
            st = Speedtest()
            print("Wait!! I am checking your Internet Speed...")
            speak("Wait!! I am checking your Internet Speed...")
            dw_speed = st.download()
            up_speed = st.upload()
            dw_speed = dw_speed / 1000000
            up_speed = up_speed / 1000000
            print('Your download speed is', round(dw_speed, 3), 'Mbps')
            print('Your upload speed is', round(up_speed, 3), 'Mbps')
            speak(f'Your download speed is {round(dw_speed, 3)} Mbps')
            speak(f'Your upload speed is {round(up_speed, 3)} Mbps')

        elif ('message' in query or 'send' in query) and 'whatsapp' in query:
            phno_list = {
                'Parth': '+918140749277',
                'anyone': '+918140749277',
                'test': '+918140749277'
            }

            def send_whtmsg():
                speak('To whom do you want to send a message on WhatsApp?')
                #recipient = takecommand()
                recipient = input()
                
                if recipient in phno_list:
                    recipient_number = phno_list[recipient]
                else:
                    speak('Recipient not found. Please choose from the available contacts.')
                    return

                speak('What message would you like to send?')
                #msg = takecommand()
                msg = input()

                speak('Do you want to send it immediately?')
                #act_msg = takecommand()
                act_msg = input()

                if 'yes' in act_msg:
                    send_now(recipient_number, msg)
                else:
                    speak('At what time would you like to send this message? For example, 11:21 PM')
                    #msg_time = takecommand()
                    msg_time = input("Time: ")

                    try:
                        send_later(recipient_number, msg, msg_time)
                    except ValueError:
                        speak('Invalid time format. Please try again.')

                speak('Do you want to send more WhatsApp messages?')
                more_msg = takecommand()
                if 'yes' in more_msg:
                    send_whtmsg()

            def send_now(recipient_number, msg):
                current_time = datetime.datetime.now()
                pywhatkit.sendwhatmsg(recipient_number, msg, current_time.hour, current_time.minute + 2, 30)
                k.press_and_release('enter')
                print('Your message has been sent successfully.')

            def send_later(recipient_number, msg, msg_time):
                try:
                    send_time = datetime.datetime.strptime(msg_time, '%I:%M %p')
                    pywhatkit.sendwhatmsg(recipient_number, msg, send_time.hour, send_time.minute, 30)
                    print('Your message has been scheduled to be sent at the specified time.')
                except ValueError:
                    raise ValueError('Invalid time format')

            send_whtmsg()

        elif 'how' in query and 'to' in query:
            try:
                # query = query.replace('how to', '')
                max_results = 1
                data = search_wikihow(sentence, max_results)
                # assert len(data) == 1
                data[0].print()
                speak(data[0].summary)
            except Exception as e:
                print(f"Error: {e}")
                speak('Sorry, I am unable to find the answer for your query.')
                        
        #elif 'news' in query or 'news headlines' in query:
        #    url = "https://news.google.com/news/rss"
        #    client = webbrowser(url)
        #    xml_page = client.read()
        #    client.close()
        #    page = bs4.BeautifulSoup(xml_page, 'xml')
        #    news_list = page.findAll("item")
        #    speak("Today's top headlines are--")
        #    try:
        #        for news in news_list:
        #            print(news.title.text)
        #            # print(news.pubDate.text)
        #            speak(f"{news.title.text}")
        #            # speak(f"{news.pubDate.text}")
        #            print()

        #    except Exception as e:
        #        pass

        elif 'screen' in query and 'recording' in query:
            try:
                print('Press Q to stop and save recording')
                record.screen_record()
                print("Screen recording started...")

                while True:
                    if keyboard.is_pressed('q'):
                        record.stop_screen_record()
                        print("Screen recording stopped and saved.")
                        break
            except Exception as e:
                print(f"Error during screen recording: {e}")

        elif "news" in query: #done
            try:
                response = requests.get('https://www.bbc.com/news')
                if response.status_code == 200:
                    print("Successfully downloaded the page's contents.")
                    soup = BeautifulSoup(response.content, 'html.parser')
                    news_headings = soup.find_all('h2')
                    for heading in news_headings:
                        print(heading.text)
                else:
                    print(f"Failed to download page contents. Status code: {response.status_code}")
            except Exception as e:
                print(f"Error fetching news: {e}")

        elif "exit" in query:
            speak("ending program, thanks for using")
            os.system("cls")
            print(f"{bcolors.OKCYAN + creditsScreen + bcolors.ENDC}")
            print("WORK WITH COMMAND ended sucessfully")
            looper = 50

        else:                                                                                                           
            print(f"\n\n{bcolors.FAIL}REQUEST ERROR\n\n{bcolors.ENDC}")
