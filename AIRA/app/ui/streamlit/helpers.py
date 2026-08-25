import io

import soundfile as sf


def convert_audio_to_numpy(uploaded_file):
    """Read a Streamlit audio upload into (samples, sample_rate).

    Rewinds first: rendering the clip with st.audio() consumes the buffer, so a
    plain .read() here would come back empty.
    """
    uploaded_file.seek(0)
    audio_bytes = uploaded_file.read()

    audio_array, sample_rate = sf.read(io.BytesIO(audio_bytes))
    audio_array = audio_array.astype("float32")
    return audio_array, sample_rate, audio_bytes
