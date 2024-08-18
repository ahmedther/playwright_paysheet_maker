import os
from dotenv import load_dotenv

load_dotenv()

APP_USERNAME = os.environ.get("APP_USERNAME")
PASSWORD = os.environ.get("PASSWORD")
NAME = os.environ.get("NAME")
PDF_PASSWORD = os.environ.get("PDF_PASSWORD")
URL = os.environ.get("URL")
