import numpy as np
import whisper

from config.settings import stt_sample_rate


class Stt():
    def __init__(self):
        self.model = whisper.load_model("base")

    def run(self, audio_file, sample_rate=None):
        """Transcribe audio. Returns the text, or None if it was not usable.

        `sample_rate` is the rate of an incoming numpy array. Whisper assumes
        16 kHz mono when handed raw samples and does no resampling of its own,
        so browser recordings (44.1/48 kHz, often stereo) must be converted
        first or the transcript comes out garbled.
        """
        if isinstance(audio_file, np.ndarray):
            audio_file = self._prepare_array(audio_file, sample_rate)
            if audio_file is None:
                return None

        return self._whisper_wrapper(audio_file)

    def _prepare_array(self, audio, sample_rate):
        audio = np.asarray(audio, dtype="float32")

        # Downmix any multi-channel recording to mono
        if audio.ndim > 1:
            audio = audio.mean(axis=1)

        audio = audio.flatten()
        if audio.size == 0:
            print("Empty audio buffer received, rejecting conversion")
            return None

        if sample_rate and sample_rate != stt_sample_rate:
            audio = self._resample(audio, sample_rate, stt_sample_rate)

        return np.ascontiguousarray(audio, dtype="float32")

    def _resample(self, audio, source_rate, target_rate):
        print(f"Resampling audio {source_rate} Hz -> {target_rate} Hz")
        try:
            from math import gcd
            from scipy.signal import resample_poly

            divisor = gcd(int(source_rate), int(target_rate))
            resampled = resample_poly(
                audio, int(target_rate) // divisor, int(source_rate) // divisor
            )
        except ImportError:
            # Linear interpolation is good enough for speech recognition
            duration = audio.size / float(source_rate)
            target_size = int(round(duration * target_rate))
            resampled = np.interp(
                np.linspace(0.0, duration, target_size, endpoint=False),
                np.linspace(0.0, duration, audio.size, endpoint=False),
                audio,
            )

        return resampled.astype("float32")

    def _whisper_wrapper(self, audio_file):
        print("Started STT Service")
        try:
            transcript = self.model.transcribe(audio_file, temperature=0.0)
            print("RAW OUTPUT:", transcript)

            segments = transcript.get("segments") or []
            if not segments:
                print("No speech segments detected, rejecting Conversion")
                return None

            avg_logprob = sum(s["avg_logprob"] for s in segments) / len(segments)
            no_speech = max(s["no_speech_prob"] for s in segments)
            if avg_logprob < -1.0 or no_speech > 0.5:
                print("Low Clearity Sound Detected, rejecting Conversion")
                return None

            return (transcript["text"] or "").strip() or None
        except Exception as e:
            print(f"Failed a t STT Service with: {e}")
            return None
