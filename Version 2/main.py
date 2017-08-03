import queue
from workerthread import *
from progressthread import *
from Excel import *
from Menu import *
from multiprocessing import Queue




Menu_Instructions()
excelname = getXLfile()

# Getting Excel
data, workbook, sheet = loadexcel(excelname)

# Set Up Number of  Threads
numthread = 10

# Setting Up Queues
dataQueue = queue.Queue()
outputQueue = queue.Queue()

# Place Data in Queue
for i in data:
    dataQueue.put(i)


# Setting Up Progress Display
prog = progressThread(dataQueue, outputQueue, dataQueue.qsize())
prog.start()

# Setting Up Progress Display
threads = []

# Create Worker Threads
for i in range(0, numthread):
   temp = workerThread(i, dataQueue, outputQueue)
   threads.append(temp)

# Start all worker threads
for i in threads:
    i.start()

# Wait for all threads to finish(inputQueue is empty)
prog.join()
for i in threads:
    i.join()

newexcelname = excelname.split('.')[0] + "-Updated"
writeExcel(outputQueue, workbook, sheet, newexcelname)
print("\nWrote Excel File:", newexcelname)

