# -*- coding: utf-8 -*-
from dotenv import load_dotenv
from pathlib import Path
import os

load_dotenv(verbose=True)
#cp .env.example .env

env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

USERNAME_LOGIN = os.getenv('USERNAME_LOGIN')
PASSWORD_LOGIN = os.getenv('PASSWORD_LOGIN')
SHEET_LINK = os.getenv('SHEET_LINK')
DESTINATION_PATH = os.getenv('DESTINATION_PATH')