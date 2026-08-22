import pyttsx3
from config.settings import AIRA_speach_rate,AIRA_speach_volume,AIRA_voice
class Tts():
    def __init__(self):
        self.engine = pyttsx3.init()
    
    def speak(self,text):
        self.engine.setProperty('rate',AIRA_speach_rate) 
        self.engine.setProperty('volume', AIRA_speach_volume)
        voices = self.engine.getProperty('voices')
        print(f"Voices Found for AIRA:{voices}")
        self.engine.say(text)
        print(f"Voice in  Use: {voices[1].id}")
        self.engine.setProperty('voice', voices[1].id)
        self.engine.runAndWait()