# Library imports
import base64
from PIL import Image
from io import BytesIO
import logging
import os
import subprocess
from datetime import datetime
import re
import json
import pandas as pd

# PySpark imports (optional — only available in Databricks/Spark environments)
try:
    from pyspark.sql import SparkSession
    from pyspark.sql.types import (
        LongType, IntegerType, FloatType, DoubleType,
        TimestampType, StringType
    )
except ImportError:
    SparkSession = None
    LongType = IntegerType = FloatType = DoubleType = None
    TimestampType = StringType = None

# Configuration
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "logs")
RED = "\033[31m"
RESET = "\033[0m"

def decode64(base64_string, path_save):
    """
    Decodes a base64-encoded image string and saves it to the specified path.

    Args:
        base64_string (str): The base64-encoded image string.
        path_save (str): The file path to save the decoded image.

    Returns:
        None
    """
    image_data = base64.b64decode(base64_string)
    image = Image.open(BytesIO(image_data))
    return image.save(path_save)

def setup_logger(file_name, file_mode, log_level):
    """
    Sets up a logger for the application, writing logs to a file.

    Args:
        file_name (str): Name of the log file (without extension).
        file_mode (str): File mode for logging (e.g., 'a' for append, 'w' for write).
        log_level (str): Logging level (e.g., 'INFO', 'DEBUG').

    Returns:
        None
    """
    if len(logging.getLogger().handlers) > 0:
        # Logger already set up, do not add another handler
        return

    file_name = f"{file_name}.log"
    log_file = os.path.join(LOG_DIR, file_name)
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    set_log_level = getattr(logging, log_level.upper(), logging.INFO)

    logging.basicConfig(
        filename=log_file,
        filemode=file_mode,
        level=set_log_level,
        format="%(asctime)s-%(name)s-%(levelname)s-%(message)s"
    )

    # Set log levels for noisy libraries
    logging.getLogger("py4j").setLevel(logging.ERROR)
    logging.getLogger("pyspark").setLevel(logging.ERROR)
    logging.getLogger("org.apache").setLevel(logging.ERROR)
    logging.getLogger("comm").setLevel(logging.CRITICAL)

def install_libreoffice(log=False):
    """
    Install LibreOffice on Databricks cluster.
    This function handles the installation process including repository updates.

    Args:
        log (bool): If True, logs installation steps.

    Returns:
        bool: True if installation successful, False otherwise
    """
    try:
        if log:
            logger = logging.getLogger("Libre_office")
            print("Starting LibreOffice installation...")
            logger.debug("------Starting LibreOffice installation-------")
        # Update package repositories
        result = subprocess.run(
            ["sudo", "apt", "update"],
            capture_output=True,
            text=True,
            check=True
        )
        if log:
            print("LibreOffice Package repositories updated successfully")
            logger.debug("------LibreOffice Package repositories updated successfully-------")
        # Add LibreOffice PPA repository
        subprocess.run(
            ["sudo", "add-apt-repository", "ppa:libreoffice/ppa", "-y"],
            capture_output=True,
            text=True,
            check=False  # Don't fail if PPA already exists
        )
        # Update again after adding PPA
        subprocess.run(
            ["sudo", "apt", "update"],
            capture_output=True,
            text=True,
            check=True
        )
        # Install LibreOffice components
        if log:
            print("Installing LibreOffice components...")
            logger.info("------Installing LibreOffice components-------")
        packages = [
            "libreoffice-common",
            "libreoffice-java-common",
            "libreoffice-writer",
            "libreoffice-impress",
            "libreoffice-calc",
            "openjdk-8-jre-headless",
        ]
        for package in packages:
            print(f"Installing {package}...")
            subprocess.run(
                ["sudo", "apt", "install", "-y", package],
                capture_output=True,
                text=True,
                check=True
            )
        # Verify installation
        result = subprocess.run(
            ["soffice", "--version"],
            capture_output=True,
            text=True,
            check=True
        )
        if log:
            print(f"LibreOffice installed successfully: {result.stdout.strip()}")
            logger.info(f"LibreOffice installed successfully: {result.stdout.strip()}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Installation failed: {e}")
        print(f"Error output: {e.stderr}")
        logger.critical(f"Installation failed: {e}")
        return False
    except Exception as e:
        print(f"Unexpected error during installation: {e}")
        logger.critical(f"Unexpected error during installation: {e}")
        return False

def convert_to_datetime(modified_at_str):
    """
    Converts a string with date/time attributes to a Python datetime object.

    Args:
        modified_at_str (str): String containing date/time attributes.

    Returns:
        datetime or None: The corresponding datetime object, or None if input is invalid.
    """
    if modified_at_str:
        # Extracting relevant attributes using regex
        year = int(re.search(r'YEAR=(\d+)', modified_at_str).group(1))
        month = int(re.search(r'MONTH=(\d+)', modified_at_str).group(1)) + 1  # MONTH is 0-based in Java Gregorian
        day = int(re.search(r'DAY_OF_MONTH=(\d+)', modified_at_str).group(1))
        hour = int(re.search(r'HOUR_OF_DAY=(\d+)', modified_at_str).group(1))
        minute = int(re.search(r'MINUTE=(\d+)', modified_at_str).group(1))
        second = int(re.search(r'SECOND=(\d+)', modified_at_str).group(1))
        # Create a datetime object
        dt = datetime(year, month, day, hour, minute, second)
        return dt
    return None


def  pandas_dtype_to_spark(dtype):
    """
    Maps a pandas dtype to the corresponding Spark SQL type.

    Args:
        dtype (str or numpy.dtype): The pandas data type.

:
        pyspark.sql.types.DataType: The corresponding Spark SQL type.
    """
    if dtype == 'int64':
        return LongType()
    elif dtype == 'int32':
        return IntegerType()
    elif dtype == 'float64':
        return DoubleType()
    elif dtype == 'float32':
        return FloatType()
    elif dtype == 'bool':
        return BooleanType()
    elif dtype.name.startswith('datetime'):
        return TimestampType()
    else:
        return StringType()


def merge_documents(processed_documents: list[dict]):
    
    full_text = ""
    metadata_tracker = []
    current_position = 0

    for doc in processed_documents:
        
        content = doc["content"]
        metadata = doc["metadata"]

        start = current_position
        end = start + len(content)

        full_text += content + "\n\n"

        metadata_tracker.append({
            "start": start,
            "end": end,
            "metadata": metadata
        })

        current_position = len(full_text)

    return full_text, metadata_tracker

def assign_metadata_to_chunks(chunks, metadata_tracker):
    
    chunk_documents = []
    cursor = 0

    for chunk in chunks:
        
        chunk_start = cursor
        chunk_end = cursor + len(chunk)

        chunk_metadata = []

        for meta in metadata_tracker:
            
            if not (chunk_end < meta["start"] or chunk_start > meta["end"]):
                chunk_metadata.append(meta["metadata"])

        chunk_documents.append({
            "content": chunk,
            "metadata": chunk_metadata
        })

        cursor += len(chunk)

    return chunk_documents

def convert_langchain_doc(chunked_documents):
    documents = [
        Document(
            page_content=doc["content"],
            metadata={
                "element_metadata": json.dumps(doc["metadata"])
            }
        )
        for doc in chunked_documents
    ]
    return documents
    
if __name__ == "__main__":
    # Entry point for script execution
    pass
