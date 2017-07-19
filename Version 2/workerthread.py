import time
import threading
from Classroom import *

def get_data(id, inQ, outQ):
    while not inQ.empty():
        data = inQ.get()

        # Check Against any Excel cells that were blank instead of holding a ID value
        if data[1] == None:
            break
        health_class = Classroom(data[1])

        # Checking if Class Webpage is correctly loaded
        if health_class.grabClassPage():
            health_class.grabRoom()
            health_class.grabInstructor()
        outQ.put([data[0], health_class.room, health_class.instructor])
        time.sleep(2)

def printthread(id, data):
    print("Thread ",id,": ",data)

class workerThread(threading.Thread):
    def __init__(self, ID, inQ, outQ):
        threading.Thread.__init__(self)
        self.ID = ID
        self.inQ = inQ
        self.outQ = outQ

    def run(self):
        get_data(self.ID, self.inQ, self.outQ)