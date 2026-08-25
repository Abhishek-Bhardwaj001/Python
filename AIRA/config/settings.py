# Listener Duration
listen_duration = 5
listen_fs = 48000
# Input-device index used by the CLI microphone listener. Set to ``None`` to
# let sounddevice use the operating system's default input device.
listen_device = 9

# Whisper expects 16 kHz mono; browser recordings come in at 44.1/48 kHz
stt_sample_rate = 16000

# Text to Speech: "pyttsx3" (offline) or "elevenlabs" (API)
tts_engine = "pyttsx3"
tts_engines = ["pyttsx3", "elevenlabs"]

# Conversation thread used by the CLI entry point
default_thread_id = "local-cli"


# Orchestrator LLM
orchestrator_prompt = """You are an Helpful AI Agent Capable of responding to USer Queries with poiltness and Completeness.
                        Do's:
                        1. Always keep the conversation engaging with interesting follow ups and inputs for deep dive.
                        2. Always respond as having a conversation with the User only use paragraphs (What you can read out alound)
                         and not tables or un-necessary Special characters where not required.
                        3. Answer in under 200-250 words with Suggestions for deep dive.
                        Donts':
                        1. If not confident say 'I don't know.'
                        2. If Confused ask for clarification questions.
                        """

# ========================== AIRA (Orchestrator) ====================
AIRA_speach_rate = 180
AIRA_speach_volume = 1.0
AIRA_voice = 1
