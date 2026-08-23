import pyttsx3
from config.settings import AIRA_speach_rate,AIRA_speach_volume,AIRA_voice

class Tts():
    def __init__(self):
        self.engine = pyttsx3.init()
        self.voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', self.voices[AIRA_voice].id)
        self.engine.setProperty('rate',AIRA_speach_rate) 
        self.engine.setProperty('volume', AIRA_speach_volume)
    
    def speak(self,text):
        self.engine.stop()
        self.engine.say(text)
        print(f"Voice in  Use: {self.voices[AIRA_voice].id}")
        self.engine.runAndWait()
        self.engine.stop()