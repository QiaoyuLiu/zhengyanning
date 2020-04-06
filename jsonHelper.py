import json
import dataHelper


def getProperties():
    with open("properties.json",'r') as properties:
        data = json.load(properties)
    return data
