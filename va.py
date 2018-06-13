import speech_recognition as sr
import webbrowser #opening web
import wolframalpha #mathematical operations
import time
import os #file handling
import pyperclip #clipping
#import pyttsx3 as tts #text to speech_recognition
import win32com.client as wincl

v=wincl.Dispatch("SAPI.SpVoice")
#cl=wolframalpha.Client('YVH8AY-R8H93LQAJ2')
#att=cl.query('Test/Attempt')
r=sr.Recognizer()
r.pause_threshold=0.7
r.energy_threshold=500
shell=wincl.Dispatch("WScript.Shell")
#v.Speak('Hello! For a list of commands, please say "Keyword list"...')
v.Speak('At your service Sir!')
print('Hello! For a list of commands, please say "Keyword list"...')

#List of commands
google = 'search for'
acad = 'academic search'
sc = 'deep search'
wkp = 'wiki page for'
rdds = 'read this text'
sav = 'save this text'
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

while True:
    with sr.Microphone() as source:
        try:
            print("Please Speak")
            audio=r.listen(source,timeout=4)
            message=str(r.recognize_google(audio))
            print('You said: '+message)
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
                words=message.split()
                del words[0:3]
                st=' '.join(words)
                print("Wikipedia Page of : "+str(st))
                url='https://en.wikipedia.org/w/index.php?search='+st
                webbrowser.open(url)
                v.Speak('Wiki Page for: '+str(st))
        except:
            break
        finally:
            break
