from app.services.listener import Listen
from app.services.stt_services import Stt
from app.agents.orchestrator_agent import AIRA
from app.services.tts_services import Tts
from app.agents.workflow import Workflow

class VoiceController():
    def __init__(self,listen_duration,listen_fs):
        self.stt = Stt()
        self.listener = Listen(listen_duration,listen_fs)
        self.lst_duration = listen_duration
        self.lst_fs = listen_fs
        self.llm_agent = Workflow()
        self.tts = Tts()
    
    def run(self):
        """Main Pipeline"""
        while True:
             i = input("Press 'S' to Speak and 'C' to Interupt program:")
             if i=="S":
                audio = self._listen()
                print(f"Step 1 Complete:{audio}")
                audio_processed = self._speech_to_text(audio)
                print(f"Step 2 Complete: {audio_processed}")
                response = self._get_agent_response(audio_processed)
                print(f"Step 3 Complete: {response}")
                self._speak(response)
                print(f"Process Complete")
             else:
                print("Program Interupted")
                break
    
    def _listen(self):
        """Listen to User Voice commands"""

        return self.listener.run()

    def _speech_to_text(self,voice_file):
        """Convert the Listen audio into text"""
        return self.stt.run(voice_file)

    def _get_agent_response(self,text):
        """Convert the LLM text in audio"""

        return self.llm_agent.run(text)

    def _speak(self,llm_response):
        """load and speak the audi converted text"""
        return self.tts.speak(llm_response[:200]) # Remove later