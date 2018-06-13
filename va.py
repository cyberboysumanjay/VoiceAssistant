import speech_recognition as sr
import webbrowser #opening web
import wolframalpha #mathematical operations
import time
import os #file handling
import pyperclip #clipping
import pyttsx3 as tts #text to speech_recognition

#v=wincl.Dispatch("SAPI.SpVoice")
cl=wolframalpha.Client('YVH8AY-R8H93LQAJ2')
att=cl.query('Test/Attempt')
r=sr.Recognizer()
r.pause_threshold=0.7
r.energy_threshold=400
