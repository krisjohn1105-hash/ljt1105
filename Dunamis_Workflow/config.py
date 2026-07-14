from dotenv import load_dotenv
import os

load_dotenv()

# Enfusion
ENF_USERNAME = os.getenv("ENF_USERNAME")
ENF_PASSWORD = os.getenv("ENF_PASSWORD")

# REACT
PRELUDE_BLOCK = os.getenv("PRELUDE_BLOCK")
PRELUDE_IPO = os.getenv("PRELUDE_IPO")
PRELUDE_MACRO = os.getenv("PRELUDE_MACRO")
PRELUDE_EVENT = os.getenv("PRELUDE_EVENT")
PRELUDE_LS = os.getenv("PRELUDE_LS")
PRELUDE = [PRELUDE_BLOCK, PRELUDE_MACRO, PRELUDE_EVENT, PRELUDE_LS]

# Microsoft
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# Dooray
DOORAY_TOKEN = os.getenv("DOORAY_TOKEN")

# File Path
FINANCING_PATH = os.getenv("FINANCING_PATH")