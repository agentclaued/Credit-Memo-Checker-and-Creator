"""Tests never touch the real data/: the database and uploads go to a temp folder before any app module is imported."""
import os
import tempfile

_tmp = tempfile.mkdtemp()
os.environ["CREDIT_DB_PATH"] = os.path.join(_tmp, "test.sqlite")
os.environ["CREDIT_UPLOAD_DIR"] = os.path.join(_tmp, "uploads")
