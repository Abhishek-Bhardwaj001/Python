import whisper

class Stt():
    def __init__(self):
        self.model = whisper.load_model("base")
    
    def run(self,audio_file):
        result = self._whisper_wrapper(audio_file)
        return result
    
    def _whisper_wrapper(self,audio_file): 
        print("Started STT Service")
        try:
            transcript = self.model.transcribe(audio_file,temperature=0.0)
            print("RAW OUTPUT:", transcript)
            segments = transcript["segments"]
            avg_logprob = sum(s["avg_logprob"] for s in segments) / len(segments)
            no_speech = max(s["no_speech_prob"] for s in segments)
            if avg_logprob < -1.0 or no_speech>0.5:
                print("Low Clearity Sound Detected, rejecting Conversion")
                return None 
            else:  
                return transcript["text"]
        except Exception as e:
            print(f"Failed a t STT Service with: {e}")