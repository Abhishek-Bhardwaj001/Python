from app.controllers.voice_controller import VoiceController
from config.settings import listen_duration,listen_fs

# ==============================Initialize Voice Commands=====================================
print("Entered main")
voice_agent = VoiceController(listen_duration,listen_fs)

print(voice_agent.run())