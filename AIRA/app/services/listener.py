import sounddevice as sd
from scipy.io.wavfile import write
import numpy as np
import librosa

class Listen():
    def __init__(self,duration,fs):
        self.duration = duration
        self.fs = fs

    def run(self):
       print("Invoked Listen Class")
       return  self._sd_listener()

    def _sd_listener(self):
        print()
        print(f"Listening through {sd.query_devices()}...")
        audio = sd.rec(int(self.duration * self.fs), samplerate=self.fs, channels=1,device=9)
        sd.wait()
        print("Done")
        audio = audio.flatten()               # You need to flatten your audio becuase thats the input whisper expects while working with audio files
        audio = audio.astype(np.float32) 
        audio = librosa.resample(audio, orig_sr=self.fs, target_sr=16000)
        audio, _ = librosa.effects.trim(audio, top_db=20)
        print("Max amplitude:", np.max(np.abs(audio)))       
        audio = audio * 10
        if np.max(np.abs(audio)) > 1:
            audio = audio / np.max(np.abs(audio))
        return audio