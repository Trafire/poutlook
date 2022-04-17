import win32com.client as win32

from .constants import DISPLAY, MAIL_ITEM


def emailer(recipients, subject, text) -> None:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(MAIL_ITEM)
    mail.To = recipients
    mail.Subject = subject
    mail.HTMLBody = text

    if DISPLAY:
        mail.Display(DISPLAY)
    else:
        mail.send()
