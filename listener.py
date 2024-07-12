import os
import subprocess
import speech_recognition as sr

def take_command():
    r = sr.Recognizer()
    r.pause_threshold = 2
    r.operation_timeout = 5
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source, phrase_time_limit=8)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        return query.lower()
    except Exception as e:
        print(e)
        print("Unable to recognize your voice.")
        return ""

def run_voice_assistant():
    print("Trigger phrase detected. Running voice assistant...")
    # Change directory to your desired location
    os.chdir(r"C:\Users\User1\Documents\G70\Command It")
    # Run the main.py script
    subprocess.run(["python", "main.py"])

def main():
    while True:
        query = take_command()
        if 'assist me' in query:
            run_voice_assistant()

if __name__ == "__main__":
    main()
