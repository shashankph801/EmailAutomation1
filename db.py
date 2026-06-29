import os
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

_client = None
_db = None


def get_db():
    global _client, _db
    if _db is None:
        uri = os.environ.get('MONGO_URI', 'mongodb://localhost:27017')
        db_name = os.environ.get('MONGO_DB_NAME', 'email_automation')
        _client = MongoClient(uri, serverSelectionTimeoutMS=5000)
        _db = _client[db_name]
    return _db


def get_hr_user(email):
    if not email:
        return None
    db = get_db()
    return db.hr_users.find_one({'email': email.lower().strip()}, {'_id': 0})
