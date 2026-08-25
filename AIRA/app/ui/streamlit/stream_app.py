from uuid import uuid4

import streamlit as st

from app.controllers.voice_controller import VoiceController
from app.ui.streamlit.helpers import convert_audio_to_numpy
from config.settings import listen_duration, listen_fs, tts_engine, tts_engines


@st.cache_resource
def get_voice_agent():
    """Build the controller once for the whole server.

    Streamlit re-runs this script top to bottom on every interaction. Without
    caching, each turn would load Whisper again and hand the workflow a brand
    new MemorySaver, wiping the conversation. Threads are separated by
    thread_id, so one shared agent across sessions is safe.
    """
    return VoiceController(listen_duration, listen_fs)


voice_agent = get_voice_agent()

st.title("A.I.R.A Voice Assistant")

if "messages" not in st.session_state:
    st.session_state.messages = []

if "thread_id" not in st.session_state:
    st.session_state.thread_id = str(uuid4())

# ============================== Sidebar ==============================
with st.sidebar:
    st.subheader("Settings")

    selected_engine = st.selectbox(
        "Voice engine",
        tts_engines,
        index=tts_engines.index(tts_engine),
    )
    voice_agent.set_tts_engine(selected_engine)

    if st.button("New chat"):
        voice_agent.reset_conversation(st.session_state.thread_id)
        st.session_state.messages = []
        st.session_state.thread_id = str(uuid4())
        st.rerun()

    st.caption(f"Thread: {st.session_state.thread_id[:8]}")

# ============================== Transcript ==============================
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

        if msg.get("audio") is not None:
            st.audio(msg["audio"], format=msg.get("mime", "audio/wav"))

prompt = st.chat_input("Speak or type something...", accept_audio=True)

if prompt:
    result = None
    user_audio = None

    # 🟢 TEXT
    if prompt.text:
        result = voice_agent.process_once(
            text_input=prompt.text,
            history=st.session_state.messages,
            thread_id=st.session_state.thread_id,
        )

    # 🔵 AUDIO
    elif prompt.audio:
        audio_array, sample_rate, user_audio = convert_audio_to_numpy(prompt.audio)
        result = voice_agent.process_once(
            audio_input=audio_array,
            sample_rate=sample_rate,
            history=st.session_state.messages,
            thread_id=st.session_state.thread_id,
        )

    if result:
        user_text = result["user_text"]
        response = result["response_text"]
        audio_bytes = result["audio"]

        # Only record a user turn when we actually understood one, so a failed
        # transcription never enters the history the LLM rehydrates from.
        if user_text:
            st.session_state.messages.append({
                "role": "user",
                "content": user_text,
                "audio": user_audio,
                "mime": "audio/wav",
            })

        st.session_state.messages.append({
            "role": "assistant",
            "content": response,
            "audio": audio_bytes,
            "mime": result.get("mime", "audio/wav"),
        })

    st.rerun()
