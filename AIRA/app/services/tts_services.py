import os
import tempfile

import pyttsx3
from dotenv import load_dotenv
from elevenlabs.client import ElevenLabs

from config.settings import AIRA_speach_rate, AIRA_speach_volume, AIRA_voice, tts_engine

load_dotenv()


class Tts():
    def __init__(self, speech_model=tts_engine):
        self.speech_model = speech_model

        if self.speech_model == "pyttsx3":
            self.mime = "audio/wav"
            self._init_pyttsx3()

        elif self.speech_model == "elevenlabs":
            self.mime = "audio/mpeg"
            self._init_elevenlabs()

        else:
            raise ValueError(f"Unsupported speech model '{speech_model}'")

# ========================= Main ============
    def speak(self, text):
        if self.speech_model == "pyttsx3":
            return self._pyttsx3_speak(text)
        elif self.speech_model == "elevenlabs":
            return self._elevenlabs_speak(text)


# ============================ pyttsx3 ===================================
    def _init_pyttsx3(self):
        # Only probe the voice list here. The engine itself is built per call:
        # SAPI5 stops producing output if save_to_file/runAndWait is reused on a
        # long lived engine, which is exactly what a cached UI resource does.
        engine = pyttsx3.init()
        self.voice_id = engine.getProperty('voices')[AIRA_voice].id
        engine.stop()

    def _new_pyttsx3_engine(self):
        engine = pyttsx3.init()
        engine.setProperty('voice', self.voice_id)
        engine.setProperty('rate', AIRA_speach_rate)
        engine.setProperty('volume', AIRA_speach_volume)
        return engine

    def _pyttsx3_speak(self, text):
        print("Using Pyttsx3 for TTS Service")
        print(f"Voice in Use: {self.voice_id}")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as f:
            temp_path = f.name

        try:
            engine = self._new_pyttsx3_engine()
            engine.save_to_file(text, temp_path)
            engine.runAndWait()
            engine.stop()

            with open(temp_path, "rb") as f:
                return f.read()
        finally:
            try:
                os.unlink(temp_path)
            except OSError as e:
                print(f"Could not remove temp audio {temp_path}: {e}")

# ================================== Eleven Labs =========================
    def _init_elevenlabs(self):
        api_key = os.getenv("ELEVENLABS_API_KEY")
        if not api_key:
            raise RuntimeError("ELEVENLABS_API_KEY is not set. Add it to your .env or use the pyttsx3 engine.")

        self.elevenlabs = ElevenLabs(api_key=api_key)

    def _elevenlabs_speak(self, text):
        print("Using eleven labs for TTS Service")
        audio = self.elevenlabs.text_to_speech.convert(
            text=text,
            voice_id="JBFqnCBsd6RMkjVDRZzb",  # "George" - browse voices at elevenlabs.io/app/voice-library
            model_id="eleven_multilingual_v2",
            output_format="mp3_22050_32",
        )
        return b"".join(audio)
