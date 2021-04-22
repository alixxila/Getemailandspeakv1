import win32com.client
import pyttsx3

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetLast()

# print(message.SenderName)
# print(message.subject)
# print(message.Body)

engine = pyttsx3.init()
rate = engine.setProperty('rate', 150)
engine.say("Vous avez un mail de {} , Le sujet du mail est {}, le mail est le suivant : {}".format(message.SenderName, message.subject, message.Body))
engine.runAndWait()
