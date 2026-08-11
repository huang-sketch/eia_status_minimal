import os


os.environ["EIA_LLM_DISABLE"] = "1"
os.environ["ENABLE_LLM_EXTRACTION"] = "false"
os.environ["ENABLE_SCHEMA_FALLBACK"] = "false"
os.environ["ENABLE_LLM_TEXT_POLISH"] = "false"
os.environ["EIA_SURFACE_WATER_WEB_SEARCH"] = "false"
os.environ.pop("EIA_LLM_API_KEY", None)
os.environ["WEB_CONCURRENCY"] = "1"
os.environ["UVICORN_WORKERS"] = "1"
