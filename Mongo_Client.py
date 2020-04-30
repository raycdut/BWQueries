# -*- coding: utf-8 -*-
import pymongo
from pymongo import MongoClient
import json


class Mongo_Client():
    def __init__(self):
        self.MONGODB_SERVER = 'localhost'
        #self.MONGODB_PORT = 32773
        self.MONGODB_PORT = 27017
        self.MONGODB_DB = "QueriesFields"

        self.client = MongoClient(
            self.MONGODB_SERVER, self.MONGODB_PORT)

        self.db = self.client[self.MONGODB_DB]
        pass
