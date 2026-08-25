from app.services.listener import Listen
from app.services.stt_services import Stt
from app.services.tts_services import Tts
from app.agents.workflow import Workflow
from config.settings import default_thread_id, listen_device, tts_engine

STT_FAILED_REPLY = "Sorry, I could not make out what you said. Could you try again?"


class VoiceController():
    def __init__(self, listen_duration, listen_fs, speech_model=tts_engine):
        self.stt = Stt()
        self.listener = Listen(listen_duration, listen_fs, listen_device)
        self.lst_duration = listen_duration
        self.lst_fs = listen_fs
        self.llm_agent = Workflow()
        self.tts = Tts(speech_model=speech_model)

    def run(self):
        """Main Pipeline"""
        while True:
             i = input("Press 'S' to Speak and 'C' to Interupt program:")
             if i=="S":
                audio = self._listen()
                print(f"Step 1 Complete:{audio}")
                # Listen already resamples to 16 kHz, so no conversion needed here
                audio_processed = self._speech_to_text(audio)
                print(f"Step 2 Complete: {audio_processed}")
                if not audio_processed:
                    print(STT_FAILED_REPLY)
                    continue
                response = self._get_agent_response(audio_processed)
                print(f"Step 3 Complete: {response}")
                self._speak(response)
                print(f"Process Complete")
             else:
                print("Program Interupted")
                break

    def process_once(self, audio_input=None, text_input=None, sample_rate=None,
                     history=None, thread_id=default_thread_id):
        """Single interaction pipeline (UI-friendly)"""

        if audio_input is not None:
            user_text = self._speech_to_text(audio_input, sample_rate=sample_rate)
        else:
            user_text = text_input

        user_text = (user_text or "").strip()

        # Never send an empty turn to the LLM, and never write one to the thread
        if not user_text:
            return {
                "user_text": None,
                "response_text": STT_FAILED_REPLY,
                "audio": None,
                "mime": self.tts.mime,
            }

        response = self._get_agent_response(user_text, history=history, thread_id=thread_id)

        # Step 3: TTS (NO playback)
        audio_bytes = self._speak(response)

        return {
            "user_text": user_text,
            "response_text": response,
            "audio": audio_bytes,
            "mime": self.tts.mime,
        }

    def set_tts_engine(self, speech_model):
        """Swap the voice without touching the LLM workflow or its checkpointer."""
        if speech_model == self.tts.speech_model:
            return

        self.tts = Tts(speech_model=speech_model)

    def reset_conversation(self, thread_id=default_thread_id):
        """Forget everything the agent remembers about a thread."""
        self.llm_agent.reset(thread_id)

    def _listen(self):
        """Listen to User Voice commands"""

        return self.listener.run()

    def _speech_to_text(self, voice_file, sample_rate=None):
        """Convert the Listen audio into text"""
        return self.stt.run(voice_file, sample_rate=sample_rate)

    def _get_agent_response(self, text, history=None, thread_id=default_thread_id):
        """Ask the orchestrator for a reply on this conversation thread"""

        return self.llm_agent.run(text, thread_id=thread_id, history=history)

    def _speak(self, llm_response):
        """load and speak the audi converted text"""
        return self.tts.speak(llm_response)
