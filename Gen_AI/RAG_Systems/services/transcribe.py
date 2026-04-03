# Audio / Video / AI
import whisper
import ffmpeg
import tiktoken

# Standard libraries
import math
import zipfile
import os
import shutil
import uuid
from pathlib import Path
import tempfile
import subprocess

class Transcribe():
    def __init__(self):
        pass

    def transcribe_multimedia(self,video_binary,verbose=False,timestamps=False):
        """
        Transcribes the text content from multimedia files, such as videos, by detecting the file type and extracting text. This function utilizes helper methods to determine the file format and extract textual data.

        Args:
            video_binary (bytes): The binary content of the multimedia file (e.g., video).
            verbose (bool, optional): If set to True, enables prints debug information to the console. Defaults to False.
            timestamps (bool, optional): If set to True, includes timestamps in the extracted transcript. Defaults to False.

        Returns:
            list: A list containing the transcribed text from the multimedia file. The format of this list may depend on the specific implementation of the text extraction 
                method used.
        """
        transcript = []
        ext = self._detect_multimedia_type(video_binary,verbose=verbose)
        if verbose:
            print(f"--------The Detected type of the file is: {ext}----------")
        try:
            transcript = self._extract_text(video_binary,file_ext=ext,verbose=verbose,timestamps=timestamps)
            return transcript
        except Exception as e:
           raise ValueError(f"Error extracting text from file: {e}")
    
    def _detect_multimedia_type(self,data,verbose=False):
        """
        Detects the type of multimedia file based on the provided data. This function can handle both file paths (as strings) and binary data, using file signatures to 
        identify the format.

        Args:
            data (str | bytes | bytearray): The multimedia file data to inspect, either as a file path (string) or binary data.
            verbose (bool, optional): If set to True, enables prints debug information to the console. Defaults to False.

        Returns:
            str: A string representing the detected multimedia file type. Possible return values 
                include:
                - 'mov': QuickTime Movie
                - 'mp4': MPEG-4 Video
                - 'mkv': Matroska Video
                - 'mp3': MPEG Audio Layer III
                - 'mpeg': MPEG Video
                - 'Unknown': If the file type cannot be determined.

        Raises:
            Exception: Raises an exception for invalid file paths, which is verboseged if enabled.
        """
        if isinstance(data, str):
            try:
                return Path(data).suffix[1:].lower()
            except Exception as e:
                if verbose:
                    print(f"Invalid file path: {e}")
        
        elif isinstance(data, (bytes, bytearray)):
            if data[4:8] == b"ftyp":
                if data[8:12] in [b"qt  ",b"moov"]:
                    return "mov"
                return "mp4"
            
            elif data.startswith(b"\x1A\x45\xDF\xA3"):
                return "mkv"
            
            elif data.startswith(b"ID3") or data[:2] in [b"\xFF\xFB", b"\xFF\xF3", b"\xFF\xF2"]:
                return "mp3"
                
            elif data.startswith(b"\x00\x00\x01\xBA") or data.startswith(b"\x00\x00\x01\xB3"):
                return "mpeg"
            else:
                return "Unknown"
        else:
            return "Unknown"
        
    def _extract_text(self,multimedia_input,file_ext: str, model_size: str = "base",timestamps=False,verbose=False):
        """
        Extracts and transcribes audio from multimedia input (e.g., video files) using a speech recognition model. The function converts the multimedia to audio format and then transcribes the audio.

        Args:
            multimedia_input (str | bytes | bytearray): The multimedia input to transcribe, specified as a file path (string) or binary data.
            file_ext (str): The file extension of the multimedia input (e.g., 'mp4', 'mov').
            model_size (str, optional): The size of the Whisper model to use for transcription. Defaults to "base".
            timestamps (bool, optional): If True, includes start and end times in the transcribed results. Defaults to False.
            verbose (bool, optional): If set to True, enables prints debug information to the console. Defaults to False.

        Returns:
            list | str | None: A list of dictionaries containing transcriptions with timestamps if `timestamps` is True. If `timestamps` is False, returns the transcribed text as a single string. Returns None if an error occurs during processing.

        Raises:
            ValueError: Raises a ValueError if the input type is neither a string, bytes, nor bytearray.
            Exception: Raises any exceptions encountered during audio extraction or transcription, which are logged if logging is enabled.
        """
        try:
            if isinstance(multimedia_input,str):
                temp_video_path = multimedia_input
            elif isinstance(multimedia_input,(bytes,bytearray)):
                with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as temp_video:
                    temp_video.write(multimedia_input)
                    temp_video_path = temp_video.name
            else:
                raise ValueError("Invalid input type. Expected str file path or binary data")
            if verbose:
                print("---Started audio extraction---")
            audio_path = temp_video_path.replace(file_ext, ".wav")
            (
                ffmpeg.input(temp_video_path)
                .output(audio_path,format="wav", acodec="pcm_s16le", ac=1, ar="16000")
                .overwrite_output()
                .run(quiet=True)
            )
            if verbose:
                print("Started transcription---")
            model = whisper.load_model(model_size)
            result = model.transcribe(audio_path)
            if timestamps:
                transcript=[]
                for seg in result["segments"]:
                    transcript.append({"start_time":seg['start'],
                                                "end_time":seg["end"],
                                                'transcript':seg["text"]})
            else:
                transcript = result["text"]
            os.remove(temp_video_path)
            os.remove(audio_path)
            if verbose:
                print("Transcription Complete---")
            return transcript
        except Exception as e:
            if verbose:
                print(f"Error during transcription: {e}")
            raise ValueError(f"Error extracting text from file: {e}")
        
