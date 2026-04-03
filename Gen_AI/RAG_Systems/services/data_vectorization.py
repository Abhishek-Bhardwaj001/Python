# Library imports
import logging
import tiktoken
import math
from langchain_text_splitters import RecursiveCharacterTextSplitter

# Import base classes
from .text_extractor import TextExtractor
from .transcribe import Transcribe

class DataVectorization(TextExtractor, Transcribe):
    def __init__(self):
        print = logging.getLogger(__name__)

    def count_token_size(self,text: str, model:str='gpt-3.5-turbo',verbose=False):
        try:
            if verbose:
                print("---Started token count---")
            encoding = tiktoken.encoding_for_model(model)
            num_tokens = len(encoding.encode(text))
            return num_tokens
        except Exception as e:
            if verbose:
                print.error(f"Error: {e}")
                print(f"Error: {e}")
            return None
    
    def text_splitter(self,text: str, embedding_model_token_limit:int,encoding_model:str='gpt-3.5-turbo',chunk_overlap_percent=0.05,verbose=False):
        try:
            encoding = tiktoken.encoding_for_model(encoding_model)
            total_tokens = len(encoding.encode(text))
            print(f"Total Tokens: {total_tokens}")
            
            if total_tokens>embedding_model_token_limit:
                Max_chunks = math.ceil(total_tokens/embedding_model_token_limit)
                Total_char=len(text.replace('\n', ''))
                Max_Char_len_per_chunk = math.ceil(Total_char/Max_chunks)
            else:
                Max_chunks = 1
                Total_char=len(text.replace('\n', ''))
                Max_Char_len_per_chunk = math.ceil(Total_char/Max_chunks)
            if verbose:
                print(f"-------Started text splitting-------\nTotal_Tokens:{total_tokens}\n\nEstimated_Chunk_split:{Max_chunks}\n\nMaximum Characters per chunk:{Max_Char_len_per_chunk}\n\n-------------")
            text_splitter = RecursiveCharacterTextSplitter(
                chunk_size=Max_Char_len_per_chunk,
                chunk_overlap=(Max_Char_len_per_chunk*chunk_overlap_percent),
                separators=["\n\n", "\n"],
                )
            if verbose:
                print(f"-------Text splitting Complete-------")
            if total_tokens>embedding_model_token_limit:
                return text_splitter.split_text(text)
            else:
                return [text]
        
        except Exception as e:
            if verbose:
                print.error(f"Error: {e}")
                print(f"Error: {e}")
            return None

if __name__ == "__main__":
    pass
