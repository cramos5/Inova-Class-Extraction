import threading
import time
import sys

class progressThread (threading.Thread):
    def __init__(self, inQ, outQ, size):
        threading.Thread.__init__(self)
        self.outQ = outQ
        self.inQ = inQ
        self.maxsize = size
    def run(self):
        getSize(self.inQ, self.outQ, self.maxsize)


def getSize(inQ, outQ, size):
    while not inQ.empty():
        processed = outQ.qsize()
        progressnum = int((processed/size) * 100)
        sys.stdout.write("\rProcessed : %d%%" % progressnum)
        sys.stdout.flush()
        time.sleep(.5)

    sys.stdout.write("\rProcessed : 100%")
    print()