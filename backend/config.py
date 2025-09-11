import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    SQLALCHEMY_DATABASE_URI = os.getenv("DB_URI")
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    PROCESSED_FOLDER = os.getenv("PROCESSED_FOLDER", "processed")
