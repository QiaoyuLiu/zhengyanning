import json
import dataHelper
import os

def getConfigs():
    with open("properties.json",'r') as properties:
        data = json.load(properties)
    return data

def getProperties(rules):
    with open(rules,'r') as properties:
        data = json.load(properties)
    return data
