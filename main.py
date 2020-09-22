import time

import win32com.client

from plyer import notification
# I had Toast installed originally but failed to get it to work
# https://stackoverflow.com/questions/50758709/the-win10toast-distribution-was-not-found-is-displayed-while-i-execute-a-python

from SMWinservice import SMWinservice


def check_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts

    for account in accounts:
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        outbox = inbox.Folders['Outbox']
        for email in outbox.Items:
            print(email)
            # notification.notify("Title", "Body")


class PythonCornerExample(SMWinservice):
    _svc_name_ = "PythonCornerExample"
    _svc_display_name_ = "Python Corner's Winservice Example"
    _svc_description_ = "That's a great winservice! :)"

    def start(self):
        self.isrunning = True

    def stop(self):
        self.isrunning = False

    def main(self):
        i = 0
        while self.isrunning:
            check_outlook()
            time.sleep(5)


if __name__ == '__main__':
    PythonCornerExample.parse_command_line()
