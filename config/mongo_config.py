from pymongo import MongoClient
from mongoengine import connect

client = MongoClient()
db = client.auroville

connect('auroville')
