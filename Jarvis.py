from win32com.client import Dispatch
import speech_recognition as sr
import os
import webbrowser
from passwordgenerator import pwgenerator
import pyperclip
import PyDictionary
from win10toast import ToastNotifier
notification = ToastNotifier()


def speak(str):
    speak = Dispatch(('sapi.SpVoice'))
    print(str)
    speak.Speak(str)


r = sr.Recognizer()
with sr.Microphone() as source:
    print("speak anything")
    audio = r.listen(source)
    try:
        text = r.recognize_google(audio)
        print("you said :" + format(text))
        query = format(text)

    except:
        speak("Sorry I can't recognize it")
    if "hello" in query:

        speak("hi")
    elif "what can you do" in query:
        speak("check  out the list with the command")
        print("open things like -- excel,word,powerpoint,onenote,access etc\n"
              "command is  open excel for example\n"
              "it can open  like youtube and google are inbuilt \n"
              "type others if want to open \n"
              "some other feautres are genrating password,play games tell weather dictionary etc")

    elif "open Excel" in query:
        speak("opening excel")
        os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010')

    elif "open 1 note" in query:
        speak("opening Onenote")
        os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft OneNote 2010')

    elif "open PowerPoint" in query:
        speak("opening Powerpoint")
        os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft PowerPoint 2010')

    elif "open Word" in query:
        speak("opening Word")
        os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010')

    elif "open YouTube" in query:
        speak("opening Youtube")
        webbrowser.open('www.youtube.com')

    elif 'play music' in query or "play song" in query:
        speak("Here you go with music")
        music_dir = "C:\\Users\\GAURAV\\Music"
        songs = os.listdir(music_dir)
        print(songs)
        random = os.startfile(os.path.join(music_dir, songs[1]))

    elif "open Google" in query:
        speak("opening Google")
        webbrowser.open('www.Google.com')

    elif "open Instagram" in query:
        speak("opening Instagram")
        webbrowser.open('www.Instagram.com')

    elif "open YouTube" in query:
        speak("opening Youtube")
        webbrowser.open('')

    elif "open SST class" in query:
        speak("opening Sst class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTAwMDY1NzAwNDRa')

    elif "open science class" in query:
        speak("opening science class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTAwMTM4OTQ3NDBa')

    elif "open Maths class" in query:
        speak("opening Maths class")
        webbrowser.open('https://classroom.google.com/u/1/c/MTE2MTkyMDcyOTcz')

    elif "open English class" in query:
        speak("opening English class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTQ2MTAyMDgyMzla')

    elif "open Hindi class" in query:
        speak("opening Hindi class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTA3NzA5NzUxMTJa')

    elif "open French class" in query:
        speak("opening French class")
        webbrowser.open('https://classroom.google.com/u/1/c/ODA4OTE5MzE1MTFa')

    elif "open Coumputer class" in query:
        speak("opening Coumputer class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTMyNTY5MjM0Njla')

    elif "open school mail" in query:
        speak("opening gmail")
        webbrowser.open('https://mail.google.com/mail/u/1/?pli=1')

    elif "generate password" in query:
       try:
            speak("What should I name the app?")
            r = sr.Recognizer()
            with sr.Microphone() as source:
                print("speak anything")
                audio = r.listen(source)
                try:
                    text = r.recognize_google(audio)
                    print("you said :" + format(text))
                    password_name = format(text)

                except:
                    speak("Sorry I can't recognize it")
            content = password_name
            password = pwgenerator.generate()
            Myfile = open('Main.txt', 'a')
            Myfile.write(f"{content} -- {password} \n")
            Myfile.close()
            speak("password is genrated")
            pyperclip.copy(password)
       except Exception as e:
            print(e)
            speak("password wasn't genrated")

    elif "antonym" in query:
        r = sr.Recognizer()
        with sr.Microphone() as source:
            speak("please say the antoym sir")
            audio = r.listen(source)
            try:
                text = r.recognize_google(audio)
                print("you said:" + format(text))
                antonym_word = format(text)

            except:
                speak("Sorry I can't recognize it")

        antonym = PyDictionary.PyDictionary.antonym(antonym_word)

        speak(f"the synonym of the word is {antonym}")

    elif "the meaning" in query:
        r = sr.Recognizer()
        with sr.Microphone() as source:
            speak("please say the word sir")
            audio = r.listen(source)
            try:
                text = r.recognize_google(audio)
                print("you said:" + format(text))
                theword = format(text)
                word = PyDictionary.PyDictionary.meaning(theword)
                speak("The meaning of the word is" + str(word))

            except:
                speak("Sorry I can't recognize it")
    else:
        speak("sorry I can't do that")
