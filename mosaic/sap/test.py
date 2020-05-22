import time

import pyautogui
import numpy
from PIL import Image

import multiprocessing
import queue
import tqdm
import os

import screenshots
img = screenshots.getScreenShotCollection()

CO02_ORDER_LABEL_BUTTON = (18, 193, 115, 32)
CONFIRM_YES = (442, 288, 153, 26)

speedUp = False

def save(s, r):
    print(pyautogui.locateOnScreen(s))
    pyautogui.screenshot(s, region=r)


def test(s, r):
    original = numpy.array(Image.open(s))
    while 1:
        current = numpy.array(pyautogui.screenshot(region=r))
        if numpy.max(numpy.abs(original - current)) == 0:
            print("match")


def process(x):
    time.sleep(1)
    return(x*x)


def worker(i, o):
    try:
        for x in iter(i.get_nowait, "STOP"):
            o.put((multiprocessing.current_process().name, process(x)))
    except queue.Empty:
        print(multiprocessing.current_process().name, "DONE")


def testQueue():
    ITERATIONS = 29

    taskQueue = multiprocessing.Queue(ITERATIONS)
    doneQueue = multiprocessing.Queue(ITERATIONS)
    for i in range(ITERATIONS):
        taskQueue.put(i)

    with tqdm.tqdm(desc="bar", total=ITERATIONS) as progress:
        processes = []
        for i in range(os.cpu_count()):
            proc = multiprocessing.Process(target=worker, args=(taskQueue, doneQueue))
            processes.append(proc)
            proc.start()

        while 1:
            p, x = doneQueue.get()
            progress.update(1)
            print(p, "::", x)
            if x >= 400:
                for proc in processes:
                    proc.terminate()
                break

if __name__ == '__main__':
    testQueue()

    # test = [
    #     "1000692105",
    #     "1000692096",
    #     "1000692129",
    #     "1000692157",
    #     "1000692108",
    #     "1000692119",
    #     "1000692128",
    #     "1000692163",
    #     "1000692109",
    #     "1000692117",
    #     "1000692131",
    # ]
    #
    # pbar = tqdm.tqdm(test)
    # for x in pbar:
    #     pbar.write(x + "test")
    #     time.sleep(0.5)
