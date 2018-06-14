import speech_recognition as sr
import webbrowser #opening web
import wolframalpha #mathematical operations
import time
import os #file handling
import pyperclip #clipping
import wikipedia
#import pyttsx3 as tts #text to speech_recognition
import win32com.client as wincl
from datetime import datetime

v=wincl.Dispatch("SAPI.SpVoice")
cl=wolframalpha.Client('YVH8AY-R8H93LQAJ2')
att=cl.query('Test/Attempt')
r=sr.Recognizer()
r.pause_threshold=0.7
r.energy_threshold=500
shell=wincl.Dispatch("WScript.Shell")
#v.Speak('Hello! For a list of commands, please say "Keyword list"...')
v.Speak('At your service Sir!')
#print('Hello! For a list of commands, please say "Keyword list"...')

#List of commands
google = 'search for'
youtube='search YouTube for'
acad = 'academic search'
wkp = 'wiki page for'
rdds = 'read the text'
t='what is the time'
d='what is the date'
say='say'
copy='copy the text'
sav = 'save the text'
bkmk = 'bookmark this page'
vid = 'video for'
wtis = 'what is'
wtar = 'what are'
whis = 'who is'
whws = 'who was'
when = 'when'
where = 'where'
how = 'how'
paint = 'open paint'
lsp = 'silence please'
lsc = 'resume listening'
stoplst = 'stop listening'
sc = 'deep search'


while True:
    with sr.Microphone() as source:
        try:
            v.Speak("How can I help you today?")
            print("Waiting for your command")
            audio=r.listen(source,Timeout=4)
            message=str(r.recognize_google(audio))
            message='wiki page for '
            print('You said: '+message)
            v.Speak('Hmmmmmmm.. Seems Interesting')
            if google in message:
                words=message.split()
                del words[0:2]
                st=' '.join(words)
                print('Google results for: '+str(st))
                url='https://google.com/search?q='+st
                webbrowser.open(url)
                v.Speak('Google Results for: '+str(st))
            elif acad in message:
                words=message.split()
                del words[0:2]
                st=' '.join(words)
                print('Academic results for: '+str(st))
                url='https://scholar.google.com/scholar?q='+st
                webbrowser.open(url)
                v.Speak('Academic Results for: '+str(st))
            elif wkp in message:
                try:
                    words=message.split()
                    del words[0:3]
                    st=' '.join(words)
                    wkpres=wikipedia.summary(st,sentences=2)
                    try:
                        print('\n'+str(wkpres)+'\n')
                        v.Speak(wkpres)
                    except UnicodeEncodeError:
                        v.Speak("Sorry! Please try searching again")
                except wikipedia.exceptions.DisambiguationError as e:
                    print(e.options)
                    v.Speak("Too many results for this keyword. Please be more specific and retry")
                    continue
                except wikipedia.exceptions.PageError as e:
                    print("This page doesn't exist")
                    v.Speak("No results found for: "+str(st))
                    continue
            elif rdds in message:
                words=message.split()
                del words[0:3]
                st=' '.join(words)
                print("Reading the text: "+str(st))
                v.Speak('Alright reading the text: '+ str(st))
            elif say in message:
                words=message.split()
                del words[0:1]
                st=' '.join(words)
                print("Repeating the text: "+str(st))
                v.Speak('Alright, Saying..: '+ str(st))
            elif copy in message:
                words=message.split()
                del words[0:3]
                st=' '.join(words)
                pyperclip.copy(st)
                print("The text "+str(st)+" is now copied to clipboard!")
                v.Speak('The text ..'+str(st)+' is now in your clipboard... Happy Pasting!')
            elif youtube in message:
                words=message.split()
                del words[0:3]
                st=' '.join(words)
                print("Searching for "+str(st)+" on Youtube")
                url='https://www.youtube.com/results?search_query='+str(st)
                webbrowser.open(url)
                v.Speak('Youtube Search results for: '+str(st))
            elif vid in message:
                words=message.split()
                del words[0:2]
                st=' '.join(words)
                print("Searching for "+str(st)+" on Youtube")
                url='https://www.youtube.com/results?search_query='+str(st)
                webbrowser.open(url)
                v.Speak('Youtube Search results for: '+str(st))
            elif t in message:
                c=time.ctime()
                words=c.split()
                v.Speak("The time is: "+str(words[3]))
            elif d in message:
                c=time.ctime()
                words=c.split()
                v.Speak("Today, it is"+words[2]+words[1]+words[4])

        except:
            break
        finally:
            break
