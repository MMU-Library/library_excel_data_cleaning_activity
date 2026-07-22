# config.py
"""
Configuration settings for the metadata cleanup pipeline.
"""

from pathlib import Path

# Paths
BASE_DIR = Path(__file__).parent.absolute()
INPUT_DIR = BASE_DIR / "inputs"
OUTPUT_DIR = BASE_DIR / "Outputs"

# Publisher standardisation
SIMILARITY_THRESHOLD_AUTO = 85
SIMILARITY_THRESHOLD_REVIEW = 65
AUTO_BOOST_THRESHOLD = 70
FREQ_RATIO_FOR_AUTO = 3

# Author standardisation
AUTHOR_SIMILARITY_THRESHOLD_AUTO = 85
AUTHOR_SIMILARITY_THRESHOLD_REVIEW = 75
