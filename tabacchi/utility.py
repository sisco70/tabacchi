#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Copyright (C) Francesco Guarnieri 2020 <francesco@guarnie.net>
#

import os
import random
import re
import struct
import locale
import sys
import traceback
import tempfile
import threading
import keyring
import hashlib
import ftplib
import configparser
import base64
from datetime import date, datetime
from urllib import request
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_CENTER
from Crypto.Cipher import AES
import gi

from . import config
from .config import log

gi.require_version('Gtk', '3.0')
gi.require_version('Gdk', '3.0')
from gi.repository import Gdk, Gtk, Pango, GLib, Gio  # noqa: E402


PANGO_SCALE_SMALL = 0.8333333333333
STR, INT, FLOAT, CURRENCY, DATE, BOOL, DICT = ("str", "int", "float", "currency", "date", "bool", "dict")


# Festività italiane nel formato M,G (aggiunto il patrono di Firenze)
ITALIAN_HOLIDAYS = {(1, 1), (1, 6), (4, 25), (5, 1), (6, 2), (6, 24), (8, 15), (11, 1), (12, 8), (12, 25), (12, 26)}


# Calcola il giorno di Pasqua
def __easter(year):
    a = year % 19
    b = year // 100
    c = year % 100
    d = (19 * a + b - b // 4 - ((b - (b + 8) // 25 + 1) // 3) + 15) % 30
    e = (32 + 2 * (b % 4) + 2 * (c // 4) - d - (c % 4)) % 7
    f = d + e - 7 * ((a + 11 * d + 22 * e) // 451) + 114
    month = f // 31
    day = f % 31 + 1
    return month, day


# Controlla se e' una festività italiana (con patrono di Firenze)
def italian_holiday(data):
    m, d = data.month, data.day
    if (m, d) in ITALIAN_HOLIDAYS:
        return True
    else:
        me, de = __easter(data.year)
        # Calcolo il giorno di pasquetta
        dem = de % 31 + 1
        mem = me + (de // 31)

        return (m == me and d == de) or (m == mem and d == dem)


def changeWidgetFontSize(widget, size):
    widget_font = widget.get_pango_context().get_font_description()
    widget_font.unset_fields(Pango.FontMask.SIZE)
    widget_font.set_size(size * Pango.SCALE)
    widget.modify_font(widget_font)


class Preferences:
    BACKUP_STR = "Backup FTP password for {}"
    ENCRYPTION_STR = "Backup Encryption Key for {}"

    def __init__(self, program_desc, program_filename):
        self.program_desc = program_desc
        self.program_filename = program_filename
        self.backupFolder = ""
        self.backupHost = ""
        self.backupUser = ""
        self.backup = True
        self.history = 10
        self.isDBDirty = False

    def getPwd(self):
        val = keyring.get_password(self.program_desc, self.ENCRYPTION_STR.format(self.backupFolder))
        return "" if not val else val

    def setPwd(self, value):
        keyring.set_password(self.program_desc, self.ENCRYPTION_STR.format(self.backupFolder), value)

    def getBackupPwd(self):
        val = keyring.get_password(self.program_desc, self.BACKUP_STR.format(self.backupFolder))
        return "" if not val else val

    def setBackupPwd(self, value):
        keyring.set_password(self.program_desc, self.BACKUP_STR.format(self.backupFolder), value)

    def setDBDirty(self):
        self.isDBDirty = True

    # Ritorna una nuova connessione (da implementare nelle sottoclassi)
    def getConn(self):
        raise NotImplementedError

    # Ritorna un cursore data una certa connessione (da implementare nelle sottoclassi)
    def getCursor(self, conn):
        raise NotImplementedError

    # Legge le preferenze dal file indicato in PREFERENCES_FILENAME
    def load(self):
        config = configparser.ConfigParser()
        config.read(self.program_filename)

        if config.has_section("Backup"):
            self.backupHost = config.get('Backup', 'host')
            self.backupFolder = config.get('Backup', 'folder')
            self.backupUser = config.get('Backup', 'user')
            self.history = config.getint('Backup', 'history')
            self.backup = config.getboolean('Backup', 'backup')
            self.isDBDirty = config.getboolean('Backup', 'modified')

        return config

    # Modifica solo lo stato del database sul file di configurazione
    def saveModified(self, modified):
        self.isDBDirty = modified
        config = configparser.ConfigParser()
        config.read(self.program_filename)
        if config.has_section("Backup"):
            config['Backup']['modified'] = self.isDBDirty
        with open(self.program_filename, 'w') as cfgfile:
            config.write(cfgfile)

    # Salva preferenze su file, tranne le pwd che vanno nel keyring
    def save(self, config=None):
        if not config:
            config = configparser.ConfigParser()

        config['Backup'] = {'host': self.backupHost,
                            'folder': self.backupFolder,
                            'user': self.backupUser,
                            'history': self.history,
                            'backup': self.backup,
                            'modified': self.isDBDirty
                            }

        with open(self.program_filename, 'w') as cfgfile:
            config.write(cfgfile)

    # Controlla se e' necessario (o e' possibile) fare un backup
    def checkBackup(self, parent):
        result = False
        if self.backup and self.isDBDirty:
            if not self.getPwd():
                msgDialog = Gtk.MessageDialog(parent=parent, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT,
                                              type=Gtk.MessageType.ERROR, buttons=Gtk.ButtonsType.CLOSE,
                                              message_format="Il backup non è andato a buon fine.")
                msgDialog.format_secondary_text("La password per criptare i dati non è stata impostata.")
                msgDialog.set_title("Errore")
                msgDialog.run()
                msgDialog.destroy()
            else:
                result = True
        return result


def encryptFile(key, in_filename, out_filename=None, chunksize=64 * 1024):
    ''' Encrypts a file using AES (CBC mode) with the given key.
            key:
                The encryption key - a string that must be either 16, 24 or 32 bytes long. Longer keys
                are more secure.
            in_filename:
                Name of the input file
            out_filename:
                If None, '<in_filename>.enc' will be used.
            chunksize:
                Sets the size of the chunk which the function uses to read and encrypt the file. Larger chunk
                sizes can be faster for some files and machines. chunksize must be divisible by 16.
    '''
    if not out_filename:
        out_filename = in_filename + '.enc'

    iv = ''.join(chr(random.randint(0, 0xFF)) for i in range(16))
    encryptor = AES.new(key, AES.MODE_CBC, iv)
    filesize = os.path.getsize(in_filename)

    with open(in_filename, 'rb') as infile:
        with open(out_filename, 'wb') as outfile:
            outfile.write(struct.pack('<Q', filesize))
            outfile.write(iv)

            while True:
                chunk = infile.read(chunksize)
                size = len(chunk)
                if size == 0:
                    break
                elif size % 16 != 0:
                    chunk += ' ' * (16 - size % 16)

                outfile.write(encryptor.encrypt(chunk))


def decryptFile(key, in_filename, out_filename=None, chunksize=64 * 1024):
    ''' Decripta un file usando AES (CBC mode) con la chiave passata per parametro.
        Se out_filename non è specificato allora sarà il_filename senza l'ultima estensione.
        (Per esempio se in_filename è '20120129.tar.bz2.enc' allora out_filename sarà '20120129.tar.bz2')
    '''
    if not out_filename:
        out_filename = os.path.splitext(in_filename)[0]

    with open(in_filename, 'rb') as infile:
        origsize = struct.unpack('<Q', infile.read(struct.calcsize('Q')))[0]
        iv = infile.read(16)
        decryptor = AES.new(key, AES.MODE_CBC, iv)

        with open(out_filename, 'wb') as outfile:
            while True:
                chunk = infile.read(chunksize)
                if len(chunk) == 0:
                    break
                outfile.write(decryptor.decrypt(chunk))

            outfile.truncate(origsize)

# Classe generica per thread


class WorkerThread(threading.Thread):
    STARTED, ERROR, STOPPED, DONE = (0, 1, 2, 3)

    def __init__(self):
        super(WorkerThread, self).__init__()
        self.progressDialog = None
        self.status = self.STARTED
        self.error = None

    def stop(self):
        self.status = self.STOPPED

    # Callback method per aggiornare lo stato del trasferimento, ignorando il parametro
    def update(self, par=None):
        if (self.status == self.STOPPED):
            raise StopIteration
        self.progressDialog.updateProgress()

    def setProgressDialog(self, dialog):
        self.progressDialog = dialog

    def setError(self, error):
        self.error = error
        self.status = self.ERROR
        log.error(f"Thread exception: {error}")


class ProgressDialog(Gtk.Dialog):
    def __init__(self, parent, msg, secMsg, title, thread):
        super().__init__(title=title, parent=parent, modal=True)
        self.set_default_size(400, 150)
        self.set_border_width(6)
        self.set_position(Gtk.WindowPosition.CENTER_ON_PARENT)
        stopButton = self.add_button("Stop", Gtk.ResponseType.CANCEL)
        contentArea = self.get_content_area()
        box = Gtk.Box()
        box.set_orientation(Gtk.Orientation.VERTICAL)
        msgLabel = Gtk.Label()
        msgLabel.set_alignment(0, 0.5)
        msgLabel.set_padding(4, 0)
        secondaryLabel = Gtk.Label()
        secondaryLabel.set_alignment(0, 0.5)
        secondaryLabel.set_padding(4, 4)
        self.progressBar = Gtk.ProgressBar()
        self.progressBar.set_show_text(True)
        msgLabel.set_markup(f"<b>{msg}</b>")
        secondaryLabel.set_text(secMsg)

        box.pack_start(msgLabel, True, True, 0)
        box.pack_start(secondaryLabel, True, True, 0)
        box.pack_start(self.progressBar, True, True, 10)
        contentArea.pack_start(box, False, True, 6)

        self.step = 0
        self.progress = 0
        self.thread = thread

        self.errorCallback = None
        self.responseCallback = None
        self.stopCallback = None

        stopButton.connect("clicked", self.forcedClose)
        self.connect("delete-event", self.forcedClose)
        self.show_all()

    def setErrorCallback(self, errorCallback):
        self.errorCallback = errorCallback

    def setResponseCallback(self, responseCallback):
        self.responseCallback = responseCallback

    def setStopCallback(self, stopCallback):
        self.stopCallback = stopCallback

    def setSteps(self, steps):
        self.step = float(1) / steps if steps > 0 else float(1)

    def updateProgress(self, step=1):
        self.progress += (step * self.step)
        if self.progress > 1:
            self.progress = 1
        GLib.idle_add(self.progressBar.set_fraction, self.progress)

    def __pulse(self):
        if (self.thread.status in [WorkerThread.STOPPED, WorkerThread.DONE]):
            return False  # Evita che sia richiamata
        self.progressBar.pulse()
        return True

    def startPulse(self, text=''):
        self.thread.setProgressDialog(self)
        self.progressBar.set_text(text)
        self.progressBar.set_pulse_step(0.1)
        GLib.timeout_add(100, self.__pulse)
        self.thread.start()

    def start(self):
        self.thread.setProgressDialog(self)
        self.progressBar.set_fraction(0.0)
        self.thread.start()

    def forcedClose(self, *args):
        log.debug("Forced close")
        self.thread.stop()
        if self.thread.status == WorkerThread.STOPPED:
            if self.stopCallback:
                self.stopCallback()
        self.destroy()

    def close(self, *args):
        log.debug("REGULAR close")
        if self.thread.status == WorkerThread.STOPPED:
            if self.stopCallback:
                self.stopCallback()
        elif self.thread.status == WorkerThread.ERROR:
            gtkErrorMsg(self.thread.error, self)
            if self.errorCallback:
                self.errorCallback()
        elif self.thread.status == WorkerThread.DONE:
            if self.responseCallback:
                self.responseCallback(*args)
        self.destroy()

# Effettua il backup del programma su Cloud


class BackupThread(WorkerThread):
    DEFAULT_CHUNK_SIZE = 16 * 1024         # 16kb
    # Regular expression per estrarre dimensione, data, filename dalle stringhe ritornate dal comando MLSD
    MLSD = re.compile(r'type=file;size=(\d+);modify=(\d+);.*\s(\S+)$')

    def __init__(self, preferences, history=None):
        super(BackupThread, self).__init__()
        self.preferences = preferences
        self.pwd = preferences.getPwd()
        self.backupPwd = preferences.getBackupPwd()
        self.backupHost = self.preferences.backupHost
        self.backupUser = self.preferences.backupUser
        self.backupFolder = self.preferences.backupFolder
        self.history = history if history else self.preferences.history
        self.ftp = None
        self.fileList = []

    def __changeFolder(self, d):
        try:
            self.ftp.cwd(d)
        except ftplib.error_perm:
            # Nel caso di errore nel cambiare directory si prova a crearla
            self.ftp.mkd(d)
            self.ftp.cwd(d)

    # Mostra message dialog per avvertire del backup non terminato
    def showWarningMessage(self):
        msgDialog = Gtk.MessageDialog(
            parent=self.progressDialog.progressDialog, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT, type=Gtk.MessageType.WARNING,
            buttons=Gtk.ButtonsType.CLOSE, message_format="Il backup non è andato a buon fine.")
        msgDialog.format_secondary_text("Le modifiche potrebbero non essere state salvate")
        msgDialog.set_title("Attenzione")
        msgDialog.run()
        msgDialog.destroy()
        return False

    # Da implementare nelle superclassi
    def __dataBackup(self, workdir, backupDir):
        raise NotImplementedError

    # Comprime la directory del programma e la cripta
    def __dataCrypt(self):
        encFileName = None
        encPathName = None

        # Directory temporanea accessibile solo all'utente
        # sarà cancellata al prossimo riavvio
        workdir = tempfile.mkdtemp()

        try:
            backupDir = os.path.dirname(os.path.realpath(__file__))
            backupDir = os.path.split(backupDir)[0]

            dataPathName = self.__dataBackup(workdir, backupDir)

            now = datetime.now()
            dataStr = now.strftime("%Y%m%d_%H%M%S")
            encFileName = f"{dataStr}.tar.bz2.enc"
            encPathName = f"{workdir}/{encFileName}"

            encryptFile(hashlib.sha256(self.pwd + dataStr).digest(), dataPathName + ".tar.bz2", out_filename=encPathName)
        except Exception as e:
            encFileName = None
            encPathName = None
            GLib.idle_add(gtkErrorMsg, e, self.progressDialog)

        return (encFileName, encPathName)

    # Callback per estrarre la data, la dimensione e il nome dei file dalle stringhe ritornate dal comando ftp 'MLSD'
    def __getFileInfo(self, string):
        result = self.MLSD.match(string)
        if result:
            self.fileList.append([result.group(1), result.group(2), result.group(3)])

    def __authenticate(self, host, user, pwd):
        self.ftp = ftplib.FTP_TLS(self.backupHost)
        # self.ftp.set_debuglevel(2)
        self.ftp.login(self.backupUser, self.backupPwd)

    def __transferFile(self, filename, f):
        self.ftp.storbinary('STOR ' + filename, f, blocksize=self.DEFAULT_CHUNK_SIZE, callback=self.update)
        self.ftp.retrlines('MLSD', self.__getFileInfo)

    def __removeFile(self, filename):
        self.ftp.delete(filename)

    # Trasferisce un file su un server FTP, in uno specifico folder, con un specifico nome
    def run(self):
        f = None
        try:
            (encFileName, encPathName) = self.__dataCrypt()
            if not encFileName or not encPathName:
                return

            f = open(encPathName, 'rb')
            file_size = os.path.getsize(f.name)

            self.progressDialog.setSteps(5 + (file_size / self.DEFAULT_CHUNK_SIZE) + len(self.fileList))
            self.update()

            self.__authenticate(self.backupHost, self.backupUser, self.backupPwd)

            self.update()
            self.__changeFolder(self.backupFolder)
            self.update()

            self.__transferFile(os.path.basename(f.name), f)

            # Ordina la lista documenti e mantiene solo i più recenti (in base alle preferenze)
            self.fileList.sort(key=lambda fileInfo: fileInfo[1])
            del self.fileList[-self.history:]

            self.update()

            # Elimina definitivamente i documenti più vecchi
            for fileInfo in self.fileList:
                self.__removeFile(fileInfo[2])
                self.update()

        except StopIteration:
            log.debug("Stop iteration: Backup")
        except Exception as e:
            self.setError(e)
        else:
            self.status = self.DONE
        finally:
            if f:
                f.close
            try:
                if self.ftp:
                    self.ftp.quit()
            except Exception:
                pass
            GLib.idle_add(self.progressDialog.close)

        return False  # In questo modo si evita che sia eseguito altre volte...


class GladeWindow():
    def __init__(self, parent, uiFileName):
        self.parent = parent
        self.builder = Gtk.Builder()
        self.add_from_file(uiFileName)

    def add_from_file(self, uiFileName):
        uiPathName = config.RESOURCE_PATH / uiFileName
        try:
            self.builder.add_from_file(str(uiPathName))
        except Exception:
            errormsg = f"Failed to load UI XML file: {uiPathName}"
            log.error(errormsg)
            msgDialog = Gtk.MessageDialog(parent=self.parent, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT, type=Gtk.MessageType.ERROR, buttons=Gtk.ButtonsType.CLOSE, message_format=errormsg)
            msgDialog.set_title("Errore")
            msgDialog.run()
            msgDialog.destroy()
            sys.exit(1)


# Dialog per l'inserimento delle impostazioni
class PreferencesDialog(GladeWindow):
    UNDEF_STRING = "**********************"

    def __init__(self, parent, preferences):
        super().__init__(parent, "preferencesDialog.glade")
        self.preferences = preferences

        self.preferencesDialog = self.builder.get_object("preferencesDialog")
        self.preferencesDialog.set_transient_for(parent)

        self.backupFrame = self.builder.get_object("backupFrame")
        self.encryptFrame = self.builder.get_object("encryptFrame")

        self.enableBackupSwitch = self.builder.get_object("enableBackupSwitch")
        self.backupUserEntry = self.builder.get_object("backupUserEntry")
        self.backupPwdEntry = self.builder.get_object("backupPwdEntry")
        self.hostEntry = self.builder.get_object("hostEntry")
        self.folderEntry = self.builder.get_object("folderEntry")
        self.historyBackupSpinbutton = self.builder.get_object("historyBackupSpinbutton")
        self.changePwdSwitch = self.builder.get_object("changePwdSwitch")
        self.cryptPwdEntry = self.builder.get_object("cryptPwdEntry")
        self.confCryptPwdEntry = self.builder.get_object("confCryptPwdEntry")

        self.builder.connect_signals({"on_preferencesDialog_delete_event": self.close,
                                      "on_okButton_clicked": self.check,
                                      "on_cancelButton_clicked": self.close
                                      })

        adjustmentHistory = Gtk.Adjustment(value=self.preferences.history, lower=1, upper=20, step_incr=1)
        self.historyBackupSpinbutton.set_adjustment(adjustmentHistory)
        self.historyBackupSpinbutton.set_value(self.preferences.history)

        self.enableBackupSwitch.set_active(self.preferences.backup)
        self.backupFrame.set_sensitive(self.preferences.backup)
        self.changePwdSwitch.set_active(self.preferences.getPwd() is None)
        self.changePwdActive(self.preferences.getPwd() is None)

        self.folderEntry.set_text(self.preferences.backupFolder)
        self.backupUserEntry.set_text(self.preferences.backupUser)
        self.hostEntry.set_text(self.preferences.backupHost)
        self.backupPwdEntry.set_text(self.preferences.getBackupPwd())

        self.enableBackupSwitch.connect("notify::active", self.__enableBackupActive)
        self.changePwdSwitch.connect("notify::active", self.__changePwdActive)

    def __enableBackupActive(self, switch, gparam):
        self.backupFrame.set_sensitive(switch.get_active())

    def changePwdActive(self, status):
        if not status:
            self.cryptPwdEntry.set_text(self.UNDEF_STRING)
            self.confCryptPwdEntry.set_text(self.UNDEF_STRING)
        else:
            self.cryptPwdEntry.set_text("")
            self.confCryptPwdEntry.set_text("")
        self.encryptFrame.set_sensitive(status)

    def __changePwdActive(self, switch, gparam):
        self.changePwdActive(switch.get_active())

    # Utilizza un backup thread specializzato nelle superclassi
    def __getBackupThread(self, preferences, history):
        raise NotImplementedError

    # Controlla prima di salvare la configurazione
    def check(self, widget, other=None):
        if self.enableBackupSwitch.get_active():
            if self.changePwdSwitch.get_active():
                pwd = self.cryptPwdEntry.get_text().strip()
                confPwd = self.confCryptPwdEntry.get_text().strip()
                if (pwd == confPwd) and (len(pwd)) > 0:
                    msgDialog = Gtk.MessageDialog(
                        parent=self.preferencesDialog, flags=Gtk.DialogFlags.MODAL, message_type=Gtk.MessageType.WARNING,
                        buttons=Gtk.ButtonsType.YES_NO, text="Cambiando la password i vecchi backup saranno cancellati.")
                    msgDialog.format_secondary_text("Sarà generato un nuovo backup. Sei sicuro?")
                    msgDialog.set_title("Attenzione")
                    response = msgDialog.run()
                    msgDialog.destroy()
                    if (response == Gtk.ResponseType.YES):
                        self.preferences.setPwd(pwd)
                        # Prima del backup salvo le preferenze
                        self.save()
                        thread = self.__getBackupThread(self.preferences, 1)
                        progressDialog = ProgressDialog(self.preferencesDialog, "Backup in corso..", "", "Backup", thread)
                        progressDialog.setResponseCallback(self.__close)
                        progressDialog.setStopCallback(self.close)
                        progressDialog.setErrorCallback(self.close)
                        progressDialog.start()
                        return True  # Evita che il segnale destroy o delete si propaghi..
                    else:
                        return False
                else:
                    msgDialog = Gtk.MessageDialog(
                        parent=self.preferencesDialog, flags=Gtk.DialogFlags.MODAL, message_type=Gtk.MessageType.ERROR, buttons=Gtk.ButtonsType.CLOSE,
                        text="Le password non coincidono oppure sono troppo brevi.")
                    msgDialog.set_title("Errore")
                    msgDialog.run()
                    msgDialog.destroy()
                    return False
        self.save()
        self.close()

    def save(self):
        self.preferences.backup = self.enableBackupSwitch.get_active()
        self.preferences.history = self.historyBackupSpinbutton.get_value_as_int()
        self.preferences.backupFolder = self.folderEntry.get_text()
        self.preferences.backupUser = self.backupUserEntry.get_text()
        self.preferences.backupHost = self.hostEntry.get_text()
        self.preferences.setBackupPwd(self.backupPwdEntry.get_text())
        self.preferences.save()

    def close(self, widget=None, other=None):
        self.preferencesDialog.destroy()

    def __close(self):
        self.preferences.saveModified(False)
        self.preferencesDialog.destroy()

    def run(self):
        result = self.preferencesDialog.run()
        return Gtk.ResponseType.OK if (result == 0) else Gtk.ResponseType.CANCEL


# Finestra Errore, se usata per gli errori DB, non passare i messaggi
def gtkErrorMsg(e, parent, errStr=None):
    secMsg = ""
    num = len(e.args)
    if num == 0:
        msg = str(e)
    else:
        msg = e.args[0]
        if (num > 1):
            secMsg = e.args[1]

    # Messaggio d'errore forzato
    if errStr:
        secMsg = msg
        msg = errStr

    _, _, exc_traceback = sys.exc_info()
    if exc_traceback is not None:
        traceList = traceback.extract_tb(exc_traceback)[-1]
        log.error(f"({traceList[0]}:{traceList[1]}) {msg}")

    msgDialog = Gtk.MessageDialog(parent=parent, modal=True, message_type=Gtk.MessageType.ERROR, buttons=Gtk.ButtonsType.CLOSE, text=msg)
    msgDialog.format_secondary_text(secMsg)
    msgDialog.set_title("Errore")
    msgDialog.run()
    msgDialog.destroy()


class DownloadThread(WorkerThread):
    CHUNK_SIZE = 8192

    def __init__(self, url, filename, user=None, pwd=None):
        super(DownloadThread, self).__init__()
        self.url = url
        self.user = user
        self.pwd = pwd
        self.filename = filename

    def run(self):
        myresponse = None
        fileObj = None
        try:
            # Prepara i dati per l'autenticazione
            myrequest = request.Request(self.url)

            if (self.user and self.pwd):
                base64string = base64.encodestring('%s:%s' % (self.user, self.pwd))[:-1]
                myrequest.add_header("Authorization", "Basic %s" % base64string)

            # Inizia il download del file
            myresponse = request.urlopen(myrequest)

            lenght = int(myresponse.headers['Content-Length'].strip())
            self.progressDialog.setSteps((lenght / self.CHUNK_SIZE) + 1)

            # Apre il file in scrittura
            fileObj = open(self.filename, "wb")
            while True:
                chunk = myresponse.read(self.CHUNK_SIZE)
                if not chunk:
                    break
                fileObj.write(chunk)
                self.update()

        except StopIteration:
            pass
        except Exception as e:
            self.setError(e)
            if fileObj:
                fileObj.close()
                os.remove(self.filename)
        else:
            self.status = self.DONE
        finally:
            if myresponse:
                myresponse.close()
            if fileObj:
                fileObj.close()
            GLib.idle_add(self.progressDialog.close, self.filename)

        return False

# Numeric Entry


class NumEntry(Gtk.Entry):
    INT_DIGITS = 9
    DEC_DIGITS = 0
    decimal_point = locale.localeconv()["decimal_point"]
    decimal_point_value = ord(decimal_point)
    numeric_keyset = {Gdk.KEY_0, Gdk.KEY_1, Gdk.KEY_2, Gdk.KEY_3, Gdk.KEY_4, Gdk.KEY_5, Gdk.KEY_6, Gdk.KEY_7, Gdk.KEY_8, Gdk.KEY_9,
                      Gdk.KEY_KP_0, Gdk.KEY_KP_1, Gdk.KEY_KP_2, Gdk.KEY_KP_3, Gdk.KEY_KP_4, Gdk.KEY_KP_5, Gdk.KEY_KP_6, Gdk.KEY_KP_7, Gdk.KEY_KP_8, Gdk.KEY_KP_9}
    extra_keyset = {Gdk.KEY_Escape, Gdk.KEY_Return, Gdk.KEY_Right, Gdk.KEY_Left, Gdk.KEY_Tab}
    edit_keyset = {Gdk.KEY_BackSpace, Gdk.KEY_Delete}
    spin_keyset = {Gdk.KEY_Up, Gdk.KEY_Down}

    def __init__(self, box=None, box_pos=0, int_digits=INT_DIGITS, decimal_digits=DEC_DIGITS):
        super(NumEntry, self).__init__()
        self.extra_keyset = self.extra_keyset | self.spin_keyset
        self.decimal_digits = decimal_digits
        self.int_digits = int_digits
        self.set_max_length(int_digits + (0 if decimal_digits == 0 else decimal_digits + 1))
        self.set_width_chars(int_digits + (0 if decimal_digits == 0 else decimal_digits + 1))
        self.connect("key-press-event", self.__numberKeyCheck)
        if box:
            box.add(self)
            box.reorder_child(self, box_pos)
            self.show_all()

    def get_value(self):
        txt = self.get_text()
        return locale.atof(txt) if len(txt) > 0 else 0

    def set_value(self, value):
        self.set_text(locale.format(("%%.%sf" % self.decimal_digits), value))

    def __numberKeyCheck(self, widget, event):
        if (event.keyval in self.extra_keyset) or (event.keyval in self.edit_keyset):
            return False
        text = self.get_text()
        lenText = len(text)
        pos = self.get_position()
        commaPos = text.find(self.decimal_point)
        noComma = (commaPos == -1)

        # Se è stato premuto un tasto numerico
        if event.keyval in self.numeric_keyset:
            # Nel caso in cui è selezionato del testo si comporta nel modo standard
            if len(self.get_selection_bounds()) > 0:
                return False
            else:
                if noComma:
                    if lenText < self.int_digits:
                        return False
                else:
                    # Se la posizione del cursore è a sinistra o a destra della virgola
                    if pos <= commaPos:
                        if commaPos < self.int_digits:
                            return False  # Nel caso in cui le cifre intere siano minori del max consentito
                    elif (lenText - commaPos - 1) < self.decimal_digits:
                        return False  # Nel caso in cui le cifre decimali siano minori del max consentito
        elif event.keyval == self.decimal_point_value:
            # Solo se ci devono essere cifre decimali, altrimenti ignoro
            if (self.decimal_digits > 0):
                # Se non è presente nel testo e si rispetta il numo max di decimali
                # consento la pressione del tasto separatore decimali
                if noComma and (lenText - pos <= self.decimal_digits):
                    return False

        return True

# Costruisce il modello e la treeview (con possibilita di editare il contenuto di una riga)


class ExtTreeView():
    EDITABLE = '*'
    SORTABLE = '^'
    RESIZABLE = '!'
    EXPANDABLE = '+'

    modifier_chars = [EDITABLE, SORTABLE, RESIZABLE, EXPANDABLE]

    def __init__(self, modelInfoList, dataTreeview, modelCallback=None, edit_callbacks=None, properties=None, formats=None, custom_cell_funcs=None):
        self.modelDescs = []
        size = len(modelInfoList)
        self.modelTypes = [None] * size
        self.modelId = [None] * size
        self.dictionaries = dict()

        self.digits = dict()
        self.propList = []
        self.adjMin = dict()
        self.adjMax = dict()
        self.adjValues = dict()
        self.adjIndexes = dict()

        dataModel = self.__parseInfoList(modelInfoList)

        if modelCallback is not None:
            dataModel = modelCallback(dataModel)

        dataTreeview.set_model(dataModel)

        self.editable = None
        self.decimal_digits = NumEntry.DEC_DIGITS
        self.int_digits = NumEntry.INT_DIGITS
        self.min = None
        self.max = None
        self.spin = None

        i = 0
        for desc in self.modelDescs:
            adjMin = self.adjMin.get(i)
            adjMax = self.adjMax.get(i)
            adjValues = self.adjValues.get(i)
            adjIndexes = self.adjIndexes.get(i)
            column = None
            cell = None
            editCol, sortCol, expandCol, resizeCol = self.propList[i]
            digits = self.digits.get(i)
            modelId = self.modelId[i]
            dataType = self.modelTypes[modelId]
            ed_callback = None
            custom_cell_f = None
            if edit_callbacks:
                # Se inserisco una callback con l'indice -1 vale per tutte le colonne (editabili)
                if -1 in edit_callbacks:
                    ed_callback = edit_callbacks[-1]
                elif (modelId in edit_callbacks):
                    ed_callback = edit_callbacks[modelId]
            if custom_cell_funcs:
                # Se inserisco una cell_func con l'indice -1 vale per tutte le colonne
                if -1 in custom_cell_funcs:
                    custom_cell_f = custom_cell_funcs[-1]
                elif (modelId in custom_cell_funcs):
                    custom_cell_f = custom_cell_funcs[modelId]
            dataDict = None
            if modelId in self.dictionaries:
                dataDict = self.dictionaries[modelId]
            property_dict = None
            if properties and modelId in properties:
                property_dict = properties[modelId]
            format_text = None
            if formats and modelId in formats:
                format_text = formats[modelId]
            if dataType == INT:
                if desc:
                    cell = Gtk.CellRendererText()
                    cell.set_property("xalign", 1)
                    cell.set_property("editable", editCol)
                    if editCol:
                        cell.connect("edited", self.onNumberEdited, dataModel, modelId, ed_callback)
                        cell.connect("editing-started", self.onNumberStartEditing, dataModel, modelId, digits, adjMin, adjMax, adjValues, adjIndexes)
                    column = Gtk.TreeViewColumn(desc, cell, text=modelId)

            elif dataType == BOOL:
                if desc:
                    cell = Gtk.CellRendererToggle()
                    cell.set_property('activatable', editCol)
                    if editCol:
                        cell.connect("toggled", self.onToggled, dataModel, modelId, ed_callback)
                    column = Gtk.TreeViewColumn(desc, cell, active=modelId)

            elif dataType == FLOAT:
                if desc:
                    cell = Gtk.CellRendererText()
                    cell.set_property("xalign", 1)
                    cell.set_property("editable", editCol)
                    if editCol:
                        cell.connect("edited", self.onNumberEdited, dataModel, modelId, ed_callback)
                        cell.connect("editing-started", self.onNumberStartEditing, dataModel, modelId, digits, adjMin, adjMax, adjValues, adjIndexes)

                    column = Gtk.TreeViewColumn(desc, cell)
                    column.set_cell_data_func(cell, self.__floatFormat, (modelId, format_text))
            elif dataType == STR:
                if desc:
                    cell = Gtk.CellRendererText()
                    cell.set_property("xalign", 0)
                    cell.set_property("editable", editCol)
                    if editCol:
                        cell.connect("edited", self.onStringEdited, dataModel, modelId, ed_callback)
                    column = Gtk.TreeViewColumn(desc, cell, text=modelId)
            elif dataType == DICT:
                if desc:
                    if editCol:
                        cell = Gtk.CellRendererCombo()
                        dictModel = Gtk.ListStore(int, str)
                        for key, value in dataDict.items():
                            dictModel.append([key, value])
                        cell.set_property("model", dictModel)
                        cell.set_property("text-column", 1)
                        cell.set_property("has-entry", False)
                        cell.connect("changed", self.onComboChanged, dataModel, modelId, ed_callback)
                    else:
                        cell = Gtk.CellRendererText()
                        cell.set_property("xalign", 0)

                    cell.set_property("editable", editCol)
                    column = Gtk.TreeViewColumn(desc, cell, text=modelId)
                    column.set_cell_data_func(cell, self.__dictFormat, (modelId, dataDict))

            elif dataType == CURRENCY:
                if desc:
                    cell = Gtk.CellRendererText()
                    cell.set_property("xalign", 1)
                    cell.set_property("editable", editCol)
                    if editCol:
                        cell.connect("edited", self.onNumberEdited, dataModel, modelId, ed_callback)
                        cell.connect("editing-started", self.onNumberStartEditing, dataModel, modelId, digits, adjMin, adjMax, adjValues, adjIndexes)
                    column = Gtk.TreeViewColumn(desc, cell)
                    column.set_cell_data_func(cell, self.__currencyFormat, modelId)
            elif dataType == DATE:
                if desc:
                    cell = Gtk.CellRendererText()
                    cell.set_property("xalign", 1)
                    column = Gtk.TreeViewColumn(desc, cell)
                    column.set_cell_data_func(cell, self.__dateFormat, (modelId, format_text))
                    if sortCol:
                        dataModel.set_sort_func(modelId, self.__sort_date_func)
            else:
                log.warning("Unknown datamodel !!")
                if desc:
                    cell = Gtk.CellRendererText()
                    cell.set_property("xalign", 1)
                    column = Gtk.TreeViewColumn(desc, cell)

            if column:
                # By-passa le function cell predefinite per il tipo di dato
                if custom_cell_f:
                    column.set_cell_data_func(cell, custom_cell_f, modelId)
                column.set_property('alignment', 0.5)
                if cell and property_dict:
                    for key in property_dict:
                        cell.set_property(key, property_dict[key])
                column.set_resizable(resizeCol)
                column.set_expand(expandCol)
                if sortCol:
                    column.set_sort_column_id(modelId)
                dataTreeview.append_column(column)

            i += 1

    # Funzione ordinamento date
    def __sort_date_func(self, model, iter1, iter2, user_data=None):
        sort_column, _ = model.get_sort_column_id()
        data1 = model.get_value(iter1, sort_column)
        data2 = model.get_value(iter2, sort_column)
        if not data2:
            return 1
        if data1 < data2:
            return -1
        elif data1 > data2:
            return 1
        else:
            return 0

    # Analizza la lista dei tipi di dato
    def __parseInfoList(self, modelInfoList):
        # Genera il modello e controlla se ci sono un numero di cifre decimali personalizzate
        dataModelTypes = [None] * len(modelInfoList)
        pos = 0
        for info in modelInfoList:
            desc = info[0]
            dataType = info[1]
            modelId = pos if len(info) < 3 else info[2]

            # Nel caso il tipo di dati sia un dizionario, si passa direttamente il dizionario, al posto di una stringa
            if isinstance(dataType, dict):
                self.dictionaries[modelId] = dataType
                dataType = DICT
            self.modelId[pos] = modelId

            mod_str = ''

            # Se campo nascosto lo ignoro
            if (desc is not None):
                i = 0
                while (i < len(desc)) and (desc[i] in self.modifier_chars):
                    i += 1
                mod_str = desc[0:i]
                desc = desc[i:]

            # Controllo se campo editabile
            edit_prop = (self.EDITABLE in mod_str)
            # Controllo se campo ordinabile
            sort_prop = (self.SORTABLE in mod_str)
            # Controllo se campo resizable
            resize_prop = (self.RESIZABLE in mod_str)
            # Controllo se campo expandable
            expand_prop = (self.EXPANDABLE in mod_str)

            self.modelDescs.append(desc)

            self.propList.append([edit_prop, sort_prop, expand_prop, resize_prop])

            # Controllo se campo numerico con incrementi fissi
            # per es. "*float#1/i5,0,10" oppure "*int/5,1,100" o "*float#2,2/0.01,0,100" o "*float/0.5"
            place = dataType.find('/')
            if place > -1:
                adj = dataType[place + 1:].split(',')
                dataType = dataType[:place]
                spin = adj[0]
                if spin.startswith("i"):           # Nel caso che l'incremento stia in un altro campo, si cerca l'indice
                    self.adjIndexes[pos] = int(spin[1:])
                else:
                    self.adjValues[pos] = float(spin)      # Nel caso che l'incremento sia una cifra precisa

                self.adjMin[pos] = float(adj[1]) if len(adj) > 1 else float('-inf')
                self.adjMax[pos] = float(adj[2]) if len(adj) > 2 else float('inf')

            # Controllo se c'è un numero di cifre decimali personalizzato
            place = dataType.find('#')
            if place > -1:
                digstr = dataType[place + 1:].split(',')
                dataType = dataType[:place]

                intDig = int(digstr[0]) if len(digstr) > 0 else NumEntry.INT_DIGITS
                decDig = int(digstr[1]) if len(digstr) > 1 else NumEntry.DEC_DIGITS
                self.digits[pos] = (intDig, decDig)

            if dataType == STR:
                dataModelTypes[modelId] = str
            elif dataType == INT:
                if pos not in self.digits:
                    self.digits[pos] = (NumEntry.INT_DIGITS, 0)
                dataModelTypes[modelId] = int
            elif dataType == DICT:
                dataModelTypes[modelId] = int
            elif dataType == BOOL:
                dataModelTypes[modelId] = bool
            elif dataType == FLOAT:
                if pos not in self.digits:
                    self.digits[pos] = (NumEntry.INT_DIGITS, 3)
                dataModelTypes[modelId] = float
            elif dataType == CURRENCY:
                if pos not in self.digits:
                    self.digits[pos] = (NumEntry.INT_DIGITS, locale.localeconv()["frac_digits"])
                dataModelTypes[modelId] = float
            elif dataType == DATE:
                dataModelTypes[modelId] = object
            else:
                raise ValueError('DataType "%s" unknown. Legal values: "str", "int", "bool", "float", "object", "dict"' % dataType)

            self.modelTypes[modelId] = dataType
            pos += 1

        return Gtk.ListStore(*dataModelTypes)

    def onComboChanged(self, cellrenderer, path, iterator, model, col_id, callback):
        value = cellrenderer.props.model.get_value(iterator, 0)
        if callback:
            callback(cellrenderer, path, value, model, col_id)
        else:
            model[path][col_id] = value

    def onNumberStartEditing(self, cellrenderer, editable, path, model, col_id, digits, minV=None, maxV=None, spinValue=None, spinId=None):
        self.editable = editable
        self.min = minV
        self.max = maxV
        self.spin = spinValue if spinId is None else model[path][spinId]
        self.int_digits = digits[0]
        self.decimal_digits = digits[1]
        self.editable.set_max_length(self.int_digits + (0 if self.decimal_digits == 0 else self.decimal_digits + 1))
        self.editable.set_width_chars(self.int_digits + (0 if self.decimal_digits == 0 else self.decimal_digits + 1))
        editable.set_text(locale.format(("%%.%sf" % self.decimal_digits), model[path][col_id]))
        editable.connect("key-press-event", self.__numberKeyCheck)

    def __numberKeyCheck(self, widget, event):
        if (event.keyval in NumEntry.extra_keyset) or (event.keyval in NumEntry.edit_keyset):
            return False
        ed = self.editable
        text = ed.get_text()
        lenText = len(text)
        pos = ed.get_position()
        commaPos = text.find(NumEntry.decimal_point)
        noComma = (commaPos == -1)

        # Se è impostata la modalità spinbutton
        if self.spin:
            # Se è stato premuto un tasto che emula lo spinbutton
            if event.keyval in NumEntry.spin_keyset:
                val = locale.atof(text)
                if event.keyval == Gdk.KEY_Up:
                    val += self.spin
                    if val > self.max:
                        val = self.max
                elif event.keyval == Gdk.KEY_Down:
                    val -= self.spin
                    if val < self.min:
                        val = self.min
                self.editable.set_text(locale.format(("%%.%sf" % self.decimal_digits), val))
        # Se è stato premuto un tasto numerico
        elif event.keyval in NumEntry.numeric_keyset:
            # Nel caso in cui è selezionato del testo si comporta nel modo standard
            if len(ed.get_selection_bounds()) > 0:
                return False
            else:
                if noComma:
                    if lenText < self.int_digits:
                        return False
                else:
                    # Se la posizione del cursore è a sinistra o a destra della virgola
                    if pos <= commaPos:
                        if commaPos < self.int_digits:
                            return False  # Nel caso in cui le cifre intere siano minori del max consentito
                    elif (lenText - commaPos - 1) < self.decimal_digits:
                        return False  # Nel caso in cui le cifre decimali siano minori del max consentito
        elif event.keyval == NumEntry.decimal_point_value:
            # Solo se ci devono essere cifre decimali, altrimenti ignoro
            if (self.decimal_digits > 0):
                # Se non è presente nel testo e si rispetta il numo max di decimali
                # consento la pressione del tasto separatore decimali
                if noComma and (lenText - pos <= self.decimal_digits):
                    return False

        return True

    def onNumberEdited(self, widget, path, value, model, col_id, callback=None):
        self.editable.disconnect_by_func(self.__numberKeyCheck)
        self.editable = None
        self.int_digits = NumEntry.INT_DIGITS
        self.decimal_digits = NumEntry.DEC_DIGITS
        self.min = None
        self.max = None
        self.spin = None
        if len(value) == 0:
            value = "0"
        if callback:
            callback(widget, path, locale.atof(value), model, col_id)
        else:
            model[path][col_id] = locale.atof(value)

    def onStringEdited(self, widget, path, value, model, col_id, callback=None):
        self.editable = None
        if callback:
            callback(widget, path, value, model, col_id)
        else:
            model[path][col_id] = value

    def onToggled(self, widget, path, model, col_id, callback=None):
        if callback:
            callback(widget, path, model, col_id)
        else:
            model[path][col_id] = not model[path][col_id]

    def __floatFormat(self, column, cell, model, iterator, param):
        col_id, format_text = param
        data = model.get_value(iterator, col_id)
        if data is None:
            cell.set_property('text', '')
        else:
            cell.set_property('text', locale.format_string(format_text if format_text else "%.3f", data))
        return

    def __currencyFormat(self, column, cell, model, iterator, col_id):
        data = model.get_value(iterator, col_id)
        if data is None:
            cell.set_property('text', '')
        else:
            cell.set_property('text', "%s" % locale.currency(data, True, True))
        return

    def __dateFormat(self, column, cell, model, iterator, param):
        col_id, format_text = param
        data = model.get_value(iterator, col_id)
        if data is None:
            cell.set_property('text', '')
        else:
            cell.set_property('text', data.strftime(format_text if format_text else "%d %b %Y"))
        return

    def __dictFormat(self, column, cell, model, iterator, param):
        col_id, dataDict = param
        data = model.get_value(iterator, col_id)
        cell.set_property('text', dataDict[data] if data in dataDict else '')
        return


# Esempio:
# modelInfoList [("Descrizione", "str"), ("Valore", "float"), ("Prezzo", "currency")]
class ExtMsgDialog(GladeWindow):
    CANCEL, OK_CANCEL, YES_NO = (0, 1, 2)

    def __init__(self, parent, modelInfoList, msg, title, iconName, result_column_id=-1, buttons=OK_CANCEL):
        super().__init__(parent, "extMsgDialog.glade")

        self.extMsgDialog = self.builder.get_object("extMsgDialog")
        if buttons == ExtMsgDialog.OK_CANCEL:
            self.extMsgDialog.add_button("OK", Gtk.ResponseType.OK)
            self.extMsgDialog.add_button("Cancel", Gtk.ResponseType.CANCEL)
            self.result = Gtk.ResponseType.CANCEL
        elif buttons == ExtMsgDialog.YES_NO:
            self.extMsgDialog.add_button("Yes", Gtk.ResponseType.YES)
            self.extMsgDialog.add_button("No", Gtk.ResponseType.NO)
            self.result = Gtk.ResponseType.NO
        elif buttons == ExtMsgDialog.CANCEL:
            self.extMsgDialog.add_button("Cancel", Gtk.ResponseType.CANCEL)
            self.result = Gtk.ResponseType.CANCEL

        self.extMsgDialog.set_transient_for(parent)
        self.msgLabel = self.builder.get_object("msgLabel")
        self.secondaryLabel = self.builder.get_object("secondaryLabel")
        self.msgImage = self.builder.get_object("msgImage")
        self.msgImage.set_from_gicon(Gio.ThemedIcon(name=iconName), Gtk.IconSize.DIALOG)
        self.msgLabel.set_text(msg)
        self.extMsgDialog.set_title(title)
        self.dataTreeview = self.builder.get_object("dataTreeview")
        ExtTreeView(modelInfoList, self.dataTreeview)
        self.dataTreeview.get_selection().set_mode(Gtk.SelectionMode.SINGLE if result_column_id > -1 else Gtk.SelectionMode.NONE)
        self.result_column_id = result_column_id
        self.selectedValue = None

        self.builder.connect_signals({"on_extMsgDialog_destroy_event": self.close
                                      })

    def setSecondaryLabel(self, text):
        self.secondaryLabel.set_text(text)

    def setData(self, dataList):
        model = self.dataTreeview.get_model()
        for data in dataList:
            model.append(data)

    def run(self):
        self.result = self.extMsgDialog.run()
        if self.result in (Gtk.ResponseType.OK, Gtk.ResponseType.YES):
            if self.result_column_id > -1:
                model, iterator = self.dataTreeview.get_selection().get_selected()
                if iterator:
                    self.selectedValue = model.get_value(iterator, self.result_column_id)
        self.close()
        return self.result

    def close(self, widget=None, other=None):
        self.extMsgDialog.destroy()


# Entry per l'inserimento delle date
class DataEntry:
    def __init__(self, button, data, dataFormat, callBack=None):
        self.data = data
        self.__format = dataFormat
        self.callBack = callBack

        self.button = button
        self.button.connect("clicked", self.__on_click)
        self.button.set_always_show_image(True)
        self.button.set_relief(Gtk.ReliefStyle.HALF)
        self.button.set_image_position(Gtk.PositionType.RIGHT)
        self.button.set_image(Gtk.Image.new_from_icon_name(Gtk.STOCK_GO_DOWN, 0))

        self.__popover = Gtk.Popover.new(self.button)
        self.__popover.set_transitions_enabled(False)

        self.__calendar = Gtk.Calendar()
        self.__calendar.set_display_options(Gtk.CalendarDisplayOptions.SHOW_DAY_NAMES | Gtk.CalendarDisplayOptions.SHOW_HEADING)
        self.__popover.add(self.__calendar)

        if data:
            self.setDate(data)
        else:
            self.button.set_label("         ")

        self.__calendar.connect('day-selected', self.__daySelected, None)
        self.__calendar.connect('next-month', self.__dateChanged, None)
        self.__calendar.connect('prev-month', self.__dateChanged, None)
        self.__calendar.connect('prev-year', self.__dateChanged, None)
        self.__calendar.connect('next-year', self.__dateChanged, None)
        self.__calendar.connect('button-release-event', self.__buttonReleased, None)
        self.changed = False

    def setDate(self, data):
        self.data = data
        self.__calendar.select_month(data.month - 1, data.year)
        self.__calendar.select_day(data.day)
        self.button.set_label(data.strftime(self.__format))

    # Pulsante rilasciato sul calendario
    def __buttonReleased(self, widget, event, Other=None):
        if self.changed:
            self.changed = False
        else:
            self.__popover.hide()
            if self.callBack:
                self.callBack()

    def __dateChanged(self, widget, Other=None):
        self.changed = True

    def __daySelected(self, widget, Other=None):
        (year, month, day) = self.__calendar.get_date()
        self.data = date(year, month + 1, day)
        self.button.set_label(self.data.strftime(self.__format))

    def __on_click(self, widget):
        if self.__popover.get_visible():
            self.__popover.hide()
            if self.callBack:
                self.callBack()
        else:
            self.__popover.show_all()


class Report:
    HEADER, TABLE, FOOTER = (0, 1, 2)

    def __init__(self, parent, preferences, query, title, numRows, colDesc, colWidths, rowHeight, pageFooter=True):
        self.parent = parent
        self.columnsDesc = colDesc
        self.numCols = len(colDesc)
        self.numRows = numRows
        self.rowHeight = rowHeight
        self.colWidths = colWidths
        self.tableData = None
        self.preferences = preferences
        self.list = []
        self.title = title
        self.query = query
        self.count = 0
        self.row = 0
        self.__showFooter = False
        self.__showReportFooter = False
        self.changeField = None
        self.isChangedField = False
        self.tableStyle = []
        self.tableHeaderStyle = []
        self.footerStyle = []
        self.reportFooterStyle = []
        self.changeFieldStyle = []
        self.columnStyle = [(None, None, None)] * self.numCols
        self.style = []
        self.rowStyle = []
        self.pageSizeX, self.pageSizeY = A4
        self.tablePosX = 0.6 * cm
        self.tablePosY = 1.4 * cm
        self.pageHeaderFont = 'Helvetica'
        self.pageHeaderFontSize = 8
        self.pageHeaderAlignment = TA_CENTER
        self.pageFooter = pageFooter

    def setPageHeaderStyle(self, font='Helvetica', fontSize=9, alignment=TA_CENTER):
        self.pageHeaderFont = font
        self.pageHeaderFontSize = fontSize
        self.pageHeaderAlignment = alignment

    def setTableHeaderStyle(self, tableHeaderStyle):
        self.tableHeaderStyle = tableHeaderStyle

    def setTableStyle(self, tableStyle):
        self.tableStyle = tableStyle

    # Imposta lo stile del footer e di conseguenza ne abilita la visualizzazione
    def setFooterStyle(self, footerStyle):
        self.__showFooter = True
        self.footerStyle = footerStyle

    # Imposta lo stile del footer del report  e di conseguenza ne abilita la visualizzazione
    def setReportFooterStyle(self, reportFooterStyle):
        self.__showReportFooter = True
        self.reportFooterStyle = reportFooterStyle

    def setTableColumnStyle(self, index, headerStyle, tableStyle, footerStyle=[]):
        self.columnStyle[index] = (headerStyle, tableStyle, footerStyle)

    def initialize(self):
        pass

    def __fetchData(self):
        conn = None
        cursor = None
        try:
            conn = self.preferences.getConn()
            cursor = self.preferences.getCursor(conn)
            cursor.execute(self.query)
            self.list = cursor.fetchall()
        except Exception as e:
            gtkErrorMsg(e, self.parent)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
        return len(self.list)

    def __printPageFooter(self, canvas, page):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.drawString(self.tablePosX, 0.60 * cm, "Pagina %d" % page)
        canvas.restoreState()

    def __printPageHeader(self, canvas, text):
        canvas.saveState()
        header_style = ParagraphStyle('header_style', fontName=self.pageHeaderFont,
                                      fontSize=self.pageHeaderFontSize, alignment=self.pageHeaderAlignment)
        p = Paragraph(text, style=header_style)
        _, height = p.wrapOn(canvas, self.pageSizeX, self.pageSizeY)
        p.drawOn(canvas, 0, self.pageSizeY - height)
        canvas.restoreState()

    def __clean(self):
        del self.style[:]
        del self.rowStyle[:]
        self.row = 0
        for i in range(self.numRows):
            for j in range(self.numCols):
                self.tableData[i][j] = None

    # Mostra il report
    def show(self, fileObj):
        if fileObj:
            Gio.Subprocess.new(["gio", "open", fileObj.name], Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_MERGE)

    def writeTableRow(self):
        pass

    def setRowStyle(self, style):
        self.rowStyle.append([self.row, style])

    # Applica lo stile da a riga a riga ed eventualmente da colonna a colonna
    def __applyStyle(self, style, startRow, endRow, startCol=0, endCol=-1):
        if startRow <= endRow:
            for item in style:
                item = (item[0],) + ((startCol, startRow), (endCol, endRow),) + (item[1:])
                self.style.append(item)

    # Scrive le righe con i dati e applica lo stile della tabella
    def __writeTableRows(self, size):
        start = self.row
        self.isChangedField = False
        while self.row < self.numRows:
            self.writeTableRow()
            self.count += 1
            self.row += 1
            if self.count == size:
                if self.changeField:
                    self.isChangedField = True
                break
            if self.changeField and (self.list[self.count][self.changeField] != self.list[self.count - 1][self.changeField]):
                self.isChangedField = True
                break
        self.__applyStyle(self.tableStyle, start, self.row - 1)
        for i in range(self.numCols):
            style = self.columnStyle[i][self.TABLE]
            if style:
                self.__applyStyle(style, start, self.row - 1, i, i)
        # Applica (se ci sono) gli stili per riga
        for rowStyle in self.rowStyle:
            row = rowStyle[0]
            self.__applyStyle(rowStyle[1], row, row)

    def __writeTableHeader(self):
        if self.row < (self.numRows - 1 - (1 if self.changeField else 0)):
            if self.changeField:
                self.tableData[self.row][0] = self.list[self.count][self.changeField]
                self.__applyStyle(self.changeFieldStyle, self.row, self.row)
                self.row += 1
            for i in range(self.numCols):
                self.tableData[self.row][i] = self.columnsDesc[i]
            self.__applyStyle(self.tableHeaderStyle, self.row, self.row)
            for i in range(self.numCols):
                style = self.columnStyle[i][self.HEADER]
                if style:
                    self.__applyStyle(style, self.row, self.row, i, i)
            self.row += 1
        else:
            self.row = self.numRows

    def setChangeField(self, field, style=[]):
        self.changeField = field
        self.changeFieldStyle = style

    # Da implementare nelle sottoclassi
    def writeTableFooter(self):
        pass

    def __writeTableFooter(self):
        if self.row < self.numRows:
            self.writeTableFooter()
            self.__applyStyle(self.footerStyle, self.row, self.row)
            for i in range(self.numCols):
                style = self.columnStyle[i][self.FOOTER]
                if style:
                    self.__applyStyle(style, self.row, self.row, i, i)
            self.row += 1

    # Da implementare nelle sottoclassi
    def writeReportFooter(self):
        pass

    def __writeReportFooter(self, size):
        if self.row < self.numRows:
            self.writeReportFooter()
            self.__applyStyle(self.reportFooterStyle, self.row, self.row)

    def getField(self, field):
        return self.list[self.count][field]

    def writeCell(self, value, column):
        self.tableData[self.row][column] = value

    def writeFieldToCell(self, field, column):
        self.tableData[self.row][column] = self.list[self.count][field]

    def __buildPage(self, c):
        self.__printPageHeader(c, self.title)
        t = Table(self.tableData, self.colWidths, self.rowHeight, style=TableStyle(self.style))

        t.wrapOn(c, self.pageSizeX, self.pageSizeY)
        self.tablePosX = (self.pageSizeX - t.minWidth()) / 2
        t.drawOn(c, self.tablePosX, self.tablePosY)

        if self.pageFooter:
            self.__printPageFooter(c, self.page)

        c.showPage()
        self.__clean()
        self.page += 1

    # Genera il report
    def build(self):
        self.initialize()
        size = self.__fetchData()
        tmpFile = None
        if size > 0:
            # Crea la tabella per i dati
            self.tableData = [[""] * self.numCols for i in range(self.numRows)]
            tmpFile = tempfile.NamedTemporaryFile(delete=False)
            tmpFile.close()

            c = Canvas(tmpFile.name, pagesize=(self.pageSizeX, self.pageSizeY))
            c.setAuthor(config.__author__)
            c.setTitle(self.title)

            self.count = 0
            self.row = 0
            self.page = 1
            self.reportBuilded = False

            while self.count < size:
                self.__writeTableHeader()
                self.__writeTableRows(size)

                if self.row == self.numRows:
                    self.__buildPage(c)
                if self.__showFooter and self.isChangedField:
                    self.__writeTableFooter()

                if self.count == size:
                    # nel caso in cui si sia appena generata la pagina
                    if self.row == 0 and self.__showFooter:
                        self.__writeTableFooter()
                    # Nel caso in cui fossimo arrivati all'ultima riga
                    if self.row == self.numRows:
                        self.__buildPage(c)

                    if self.__showReportFooter:
                        self.__writeReportFooter(size)
                    self.__buildPage(c)
            c.save()

        return tmpFile
