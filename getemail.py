import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import string
import pyttsx3

#Les pass du mail
username = "youremail"
password = "yourpassword"

def clean(text):

    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

#Créer une class IMAP4 avec SSL
imap = imaplib.IMAP4_SSL("outlook.office365.com")

#Authentification
imap.login(username, password)

#Selection des mails dans "Inbox", il ne va pas lire le message 
status, messages = imap.select("INBOX", "UNSEEN")

# Le total des mails que la variable selectionne
N = 1

messages = int(messages[0])

for i in range(messages, messages-N, -1):
    # Récuperer le mail avec un ID + utilisation du protocole de représentation de message (RFC822)
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # Envoie de l'email dans l'objet message 
            msg = email.message_from_bytes(response[1])

            # Décodage du sujet de l'email
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                # Si c'est en bytes, décode en str
                subject = subject.decode(encoding)

            # Décodage de email sender
            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)

            # Envoie de la variable dans une variable
            f = From

            # Envoie des textes en variable
            subjectis = "le sujet est "
            destinataireis = " et Il provient du destinataire "
            
            # Envoie des textes en variable
            subjectisisclean = subject.translate({ord(i): None for i in '-'})

            # Nettoyage du message
            destinataireisclean = f.translate({ord(i): None for i in '<>'})

            # Concaténation des variables du dessus
            Cleanshit = subjectis + subjectisisclean + destinataireis + destinataireisclean

            #  Engine va convertir les mots en voix variable puis initialisation de la librairie
            engine = pyttsx3.init()

            #  Rate va définir la rapidité de la voix
            rate = engine.setProperty('rate', 132)

            #  Lit le texte entre " "
            engine.save_to_file(Cleanshit, "test.mp3")

            #  Execute et attend pour une prochaine commande
            engine.runAndWait()

imap.close()
imap.logout() 