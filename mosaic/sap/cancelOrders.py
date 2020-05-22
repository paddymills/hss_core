import pyautogui
import time

import os
import sys
import tqdm

def main():
    pyautogui.PAUSE = 0.5
    cancelOrders()


def cancelOrders():
    homePath = os.path.dirname(os.path.realpath(__file__))
    with open(os.path.join(homePath, "orders.txt"), "r") as ordersFile:
        orderList = ordersFile.read().split("\n")

    if not orderList[-1]:
        orderList.pop()

    transaction = input("Open transaction CO13 and press enter (s to skip)")
    if not transaction.startswith("s"):
        pyautogui.hotkey("alt", "tab")
        cancelOrderConfirmation(orderList)

    transaction = input("Open transaction CO02 and press enter (s to skip)")
    if not transaction.startswith("s"):
        pyautogui.hotkey("alt", "tab")
        deleteOrder(orderList)


def cancelOrderConfirmation(orders):
    with tqdm.tqdm(orders) as progress:
        for order in progress:
            pyautogui.press(["tab", "tab"])
            pyautogui.typewrite(order)
            pyautogui.hotkey("enter")
            pyautogui.hotkey("enter")
            pyautogui.hotkey("ctrl", "s")
            time.sleep(0.5)
            pyautogui.hotkey("esc")
            time.sleep(0.5)

            progress.write(f"{order} confirmation cancelled")


def deleteOrder(orders):
    with tqdm.tqdm(orders) as progress:
        for order in progress:
            pyautogui.typewrite(order)
            pyautogui.hotkey("enter")
            pyautogui.keyDown("alt")
            pyautogui.typewrite("nls", interval=0.25)
            pyautogui.keyUp("alt")
            pyautogui.hotkey("ctrl", "s")
            # time.sleep(0.5)

            progress.write(f"{order} deleted")



if __name__ == '__main__':
    main()
