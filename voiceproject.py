import speech_recognition as sr
import pyttsx3
from tkinter import *
import pyaudio
import pywhatkit
import wikipedia
import datetime
import cv2 as cv
import numpy as np
from PIL import ImageTk,Image
listener = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty('voices')
def talk(text):
    engine.say(text)
    engine.runAndWait()
def take_command():
    try:
        with sr.Microphone() as source:
                print("listening....")
                voice  = listener.listen(source)
                command = listener.recognize_google(voice)
                command=command.lower()
                if 'buddy' or 'badi' in command:
                    command=command.replace('badi','')
    except:
        pass
    return command
def run_buddy(self):
    talk('listening')
    comand=take_command()
    if 'play' in comand:
        song=comand.replace('play','')
        talk('playing'+song)
        pywhatkit.playonyt(song)
    elif 'time' in comand:
        time=datetime.datetime.now().strftime('%I:%M %p')
        talk('Current time is '+time)
    elif 'i love you' in comand:
        talk('sorry i am in a relationship')
    elif 'what is my name' in comand:
        talk('your name is syed jhonny')
    elif 'whatsapp' in comand:
        try:
            pywhatkit.sendwhatmsg("+916281670393","hello from buddy",21,58)
        except:
            talk('unexpected error occured')
    elif 'search' in comand:
        search1=comand
        search2=pywhatkit.search(search1)
    elif 'image' in comand:
        img=cv.imread('lion.jpg')
        imghor=np.hstack((img,img))
        imgver=np.vstack((img,img))
        cv.imshow('lion',img)
        cv.waitKey(0)
    elif 'video' in comand:
        talk('playing video')
        cap=cv.VideoCapture('1280.mp4')
        while True:
            success, img = cap.read()
            cv.imshow('video',img)
            if cv.waitKey(1) & 0xFF == ord('q'):
                break
    elif 'who is johnny' in comand:
        talk('he is my creator')
    elif 'webcam' in comand:
        talk('accesing your webcam')
        cap=cv.VideoCapture(0)
        cap.set(3,1280)
        cap.set(4,760)
        cap.set(10,100)
        while True:
            success, img = cap.read()
            cv.imshow('video',img)
            if cv.waitKey(1) & 0xFF == ord('q'):
                break
    elif 'wikipedia' in comand:
        person=comand
        info=wikipedia.summary(person,2)
        print(info)
        talk(info)
    else:
        talk('speak clearly')
editor=Tk()
editor.geometry('1980x1020')
my_img=ImageTk.PhotoImage(Image.open('mike.jpg'))
label=Label(image=my_img,height=200,width=150)
label.place(x=560,y=100)
b1=Button(text='Click Here To Speak',height=4,width=20,font=('Times 15',10,'bold'),fg='red',bg='white')
b1.place(x=550,y=400)
b1.bind('<Button-1>',func=run_buddy)
editor.mainloop()
