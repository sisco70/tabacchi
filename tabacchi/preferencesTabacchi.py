#
# Copyright (C) Francesco Guarnieri 2020 <francesco@guarnie.net>
#

import configparser
import keyring
import sqlite3
import subprocess
import bluetooth
import gi
import datetime

from . import config
from .config import log
from . import utility
from .utility import WorkerThread

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk, GLib  # noqa: E402


class InitBarcodeThread(WorkerThread):
    def __init__(self):
        super().__init__()

    # Inizializza
    def run(self):
        resultList = []
        try:
            devices = bluetooth.discover_devices()
            for device in devices:
                name = bluetooth.lookup_name(device)
                services = bluetooth.find_service(address=device)
                for svc in services:
                    if (svc['protocol'] == 'RFCOMM'):
                        resultList.append([name, device, svc['port']])
                    log.debug("Device found: %s %s %s" % (name, device, svc['port']))
        except Exception as e:
            self.setError(e)
        else:
            self.status = self.DONE
        finally:
            GLib.idle_add(self.progressDialog.close, resultList)

        return False


# Classe per la gestione del thread per la  connessione con il lettore di barcode bluetooth
class ConnectBarcodeThread(WorkerThread):
    def __init__(self, addr, port):
        super().__init__()
        self.addr = addr
        self.port = port
        self.sock = None

    def stop(self):
        log.debug("Barcode connect thread: STOP")
        WorkerThread.stop(self)
        self.__closeSocket(True)

    def __closeSocket(self, force=False):
        if self.sock:
            if force:
                try:
                    self.sock.shutdown(2)  # blocks both sides of the socket
                except Exception:
                    pass
                log.debug("Barcode connect thread: socket shutdown!")
            try:
                self.sock.close()
            except Exception:
                pass
            log.debug("Barcode connect thread: socket closed.")
            self.sock = None

    # Inizializza
    def run(self):
        try:
            self.sock = bluetooth.BluetoothSocket(bluetooth.RFCOMM)
            log.debug("Connecting Socket..")
            self.sock.connect((self.addr, self.port))
        except Exception as e:
            self.setError(e)
            self.__closeSocket()
        else:
            self.status = self.DONE
        finally:
            GLib.idle_add(self.progressDialog.close, self.sock)

        return False


# Classe per la gestione del thread di lettura codici a barre
class ReadBarcodeThread(WorkerThread):
    BARCODE_SUFFIX = '\r\n'

    def __init__(self, updateCallback, errorCallback, sock):
        super().__init__()
        self.sock = sock
        self.updateCallback = updateCallback
        self.errorCallback = errorCallback

    def stop(self):
        log.debug("Barcode read thread: STOP")
        WorkerThread.stop(self)
        self.__closeSocket(True)

    def __closeSocket(self, force=False):
        if self.sock:
            if force:
                try:
                    self.sock.shutdown(2)  # blocks both sides of the socket
                except Exception:
                    pass
                log.debug("Barcode read thread: socket shutdown!")
            try:
                self.sock.close()
            except Exception:
                pass
            log.debug("Barcode read thread: socket closed.")
            self.sock = None

    # Legge i dati dal dispositivo bluetooth
    def run(self):
        try:
            # Si esce solo se il socket viene chiuso (con shutdown) dall'esterno
            while not (self.status == self.STOPPED):
                data = self.sock.recv(1024)
                if data:
                    # log.debug("data: %s", data)
                    if (len(data) > 0) and data.endswith(self.BARCODE_SUFFIX):
                        GLib.idle_add(self.updateCallback, data[:-len(self.BARCODE_SUFFIX)])
                    else:
                        raise Exception("Invalid data from barcode reader.")
        except Exception as e:
            self.setError(e)
            GLib.idle_add(self.errorCallback, e)
        finally:
            self.__closeSocket()

        return False


# Opzioni del programma
class Preferences(utility.Preferences):
    TABACCHI_STR = "Logista Website password"
    DB_PATHNAME = config.user_data_dir / f'{config.PACKAGE_NAME}.sqlite'

    def __init__(self):
        super().__init__(config.__desc__, config.CONF_PATHNAME)
        self.numRivendita = ""
        self.codCliente = ""
        self.nome = ""
        self.cognome = ""
        self.citta = ""
        self.telefono = ""
        self.pianoConsegneDaSito = False
        self.giornoLevata = 3   # Default Giovedi, ma si può cambiare
        self.ggPerOrdine = 2
        self.oraInvio = 11
        self.timbro = ""
        self.firma = ""
        self.u88 = ""
        self.u88urg = ""
        self.timbroW = 0
        self.timbroH = 0
        self.firmaW = 0
        self.firmaH = 0
        self.tabacchiPwd = ""
        self.tabacchiUser = ""
        self.catalogoUrl = ""
        self.loginUrl = ""
        self.dataCatalogo = datetime.date.today()
        self.defaultBarcode = -1
        self.barcodeList = []
        self.pianoConsegneList = []

    # Controlla se è possibile ottenere una connessione con la configurazione attuale
    def checkDB(self):
        return self.DB_PATHNAME.exists()

    # Ritorna una nuova connessione
    def getConn(self):
        conn = sqlite3.connect(self.DB_PATHNAME, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
        conn.row_factory = sqlite3.Row
        conn.text_factory = str
        conn.execute("pragma foreign_keys=ON;")
        return conn

    # Ritorna un cursore, data una connessione
    def getCursor(self, conn):
        return conn.cursor()

    # Legge le preferenze dal file di configurazione
    def load(self):
        config = super().load()

        self.barcodeList[:] = []
        if config.has_section('Barcode'):
            barcode = config['Barcode']
            i = 0
            self.defaultBarcode = barcode.getint('defaultbarcode', -1)
            while config.has_option('Barcode', f'device{i}'):
                device = barcode[f'device{i}']
                port = barcode.getint(f'port{i}')
                addr = barcode[f'addr{i}']
                i += 1
                self.barcodeList.append([device, addr, port])

        if config.has_section("Tabacchi"):
            tabacchi = config['Tabacchi']
            self.numRivendita = tabacchi.get('numRivendita', '')
            self.codCliente = tabacchi.get('codCliente', '')
            self.dataCatalogo = datetime.datetime.strptime(tabacchi.get('dataCatalogo', datetime.date.today()), '%Y-%m-%d')
            self.nome = tabacchi.get('nome', '')
            self.cognome = tabacchi.get('cognome', '')
            self.citta = tabacchi.get('citta', '')
            self.telefono = tabacchi.get('telefono')
            self.pianoConsegneDaSito = tabacchi.getboolean('pianoConsegneDaSito')
            self.giornoLevata = tabacchi.getint('giornoLevata', 3)
            self.oraInvio = tabacchi.getint('oraInvio', 11)
            self.ggPerOrdine = tabacchi.getint('ggPerOrdine', 2)
            self.timbro = tabacchi.get('timbro', '')
            self.firma = tabacchi.get('firma', '')
            self.u88 = tabacchi.get('u88', '')
            self.u88urg = tabacchi.get('u88urg', '')
            self.timbroW = tabacchi.getfloat('timbroW', 0)
            self.timbroH = tabacchi.getfloat('timbroH', 0)
            self.firmaW = tabacchi.getfloat('firmaW', 0)
            self.firmaH = tabacchi.getfloat('firmaH', 0)
            self.tabacchiUser = tabacchi.get('user', '')
            self.catalogoUrl = tabacchi.get('catalogoUrl', '')
            self.loginUrl = tabacchi.get('loginUrl', '')

            value = keyring.get_password(self.TABACCHI_STR, self.tabacchiUser)
            if value:
                self.tabacchiPwd = value

            #   Carica nelle preferenze una list con il piano levate
            #   che sarà usato per generare la treeview del piano consegne nella mainwindow
            #    usare funzione a basso livello per aggiornare i dati e ad alto per rinnovare model e treeview
            if self.pianoConsegneDaSito:
                if config.has_section("PianoLevate"):
                    tmpList = config.items("PianoLevate")
                    for row in tmpList:
                        consegna = datetime.datetime.strptime(row[0], "%d/%m/%Y").date()
                        dataLimiteStr, ordine, stato, canale, tipo = row[1].split(',')
                        dataLimite = datetime.datetime.strptime(dataLimiteStr, "%d/%m/%Y - %H:%M")
                        self.pianoConsegneList.append([consegna, dataLimite, ordine, stato, canale, tipo])

    # Salva le preferenze per questo progetto
    def save(self):
        config = configparser.ConfigParser()

        if len(self.barcodeList) > 0:
            config['Barcode'] = {}
            barcode = config['Barcode']
            i = 0
            barcode['defaultbarcode'] = str(self.defaultBarcode)
            for code in self.barcodeList:
                barcode[f'device{i}'] = code[0]
                barcode[f'addr{i}'] = code[1]
                barcode[f'port{i}'] = str(code[2])
                i += 1
        config['Tabacchi'] = {'numRivendita': self.numRivendita,
                              'codCliente': self.codCliente,
                              'nome': self.nome,
                              'cognome': self.cognome,
                              'citta': self.citta,
                              'telefono': self.telefono,
                              'pianoConsegneDaSito': self.pianoConsegneDaSito,
                              'giornoLevata': self.giornoLevata,
                              'oraInvio': self.oraInvio,
                              'ggPerOrdine': self.ggPerOrdine,
                              'dataCatalogo': self.dataCatalogo.strftime('%Y-%m-%d'),
                              'timbro': self.timbro,
                              'firma': self.firma,
                              'u88': self.u88,
                              'u88urg': self.u88urg,
                              'timbroW': self.timbroW,
                              'timbroH': self.timbroH,
                              'firmaW': self.firmaW,
                              'firmaH': self.firmaH,
                              'user': self.tabacchiUser,
                              'catalogoUrl': self.catalogoUrl,
                              'loginUrl': self.loginUrl
                              }

        keyring.set_password(self.TABACCHI_STR, self.tabacchiUser, self.tabacchiPwd)

        if self.pianoConsegneDaSito:
            config["PianoLevate"] = {}
            for row in self.pianoConsegneList:
                consegna = datetime.datetime.strftime(row[0], "%d/%m/%Y")
                dataLimite = datetime.datetime.strftime(row[1], "%d/%m/%Y - %H:%M")
                config["PianoLevate"][consegna] = dataLimite + ',' + ','.join(map(str, row[2:]))

        super().save(config)


# Istanza globale
prefs = Preferences()


# Presa una data di riferimento qualsiasi, e un giorno della settimana, restituisce il giorno della
# settimana successivo alla data di riferimento (nel caso si passi anche un'ora, confronta anche quella)
#
# Per esempio:
# nextWeekday('venerdì 12 Dicembre 9:00', giovedì) = > giovedì 18 Dicembre 00:00
# nextWeekday('martedì 3 Maggio 10:27', martedi, '11:00') = > martedi 3 Maggio 11:00
def nextWeekday(d, weekday, time=datetime.time(0, 0, 0, 0)):
    days_ahead = weekday - d.weekday()
    if (days_ahead < 0) or ((days_ahead == 0) and (d.time() >= time)):
        days_ahead += 7
    return datetime.datetime.combine(d.date() + datetime.timedelta(days_ahead), time)


# In base alla data passata per parametro, restituisce il timestamp entro il quale
# si può inviare l'ordine e la data di quando l'ordine sarà consegnato (levata).
#
# Condizioni:
# - Gg lavorativi per preparare l'ordine (prefs.ggPerOrdine)
# - Non si consegna di sabato e domenica,
# - Il giorno di consegna è modificabile nelle prefs (prefs.giornoLevata)
# - Lunedi = 0, .. Domenica = 6
# - Non oltre un certo orario (prefs.oraInvio)
def dataLimiteOrdine(data):
    dataOrdine = None
    levata = None
    if prefs.pianoConsegneDaSito:
        log.debug("[dataLimiteOrdine] check prefs.pianoConsegneList: %s" % prefs.pianoConsegneList)
        for row in prefs.pianoConsegneList:
            if data < row[1]:
                dataOrdine = row[1]
                levata = row[0]
                break
    # Nel caso sia disabilitata la scansione del piano consegne dal sito Logista
    # oppure sia vuota la lista del piano consegne
    if dataOrdine is None or levata is None:
        # Giorno della settimana entro il quale ordinare
        giornoOrdine = (prefs.giornoLevata - prefs.ggPerOrdine) % 5
        # Data esatta entro la quale ordinare
        dataOrdine = nextWeekday(data, giornoOrdine, datetime.time(prefs.oraInvio, 0, 0, 0))
        # Data esatta quando sarà consegnato l'ordine
        levata = nextWeekday(dataOrdine, prefs.giornoLevata)
        levata = levata.date()

    return (dataOrdine, levata)


# Dialog per impostare le opzioni del programma
class PreferencesDialog(utility.PreferencesDialog):
    def __init__(self, parent):
        super().__init__(parent, prefs)

        self.add_from_file("preferencesTabacchi.glade")

        self.preferencesNotebook = self.builder.get_object("preferencesNotebook")
        self.tabacchiVbox = self.builder.get_object("tabacchiVbox")
        self.barcodeVbox = self.builder.get_object("barcodeVbox")

        self.preferencesNotebook.append_page(self.tabacchiVbox)
        self.preferencesNotebook.append_page(self.barcodeVbox)

        self.preferencesNotebook.set_tab_label_text(self.tabacchiVbox, "Tabacchi")
        self.preferencesNotebook.set_tab_label_text(self.barcodeVbox, "Lettore codici a barre")

        self.__buildTabacchi(self.builder)
        self.__buildBarcode(self.builder)
        self.__loadBarcode(prefs.barcodeList)

        self.builder.connect_signals({
            "on_refreshBarcodeButton_clicked": self.refreshBarcode,
        })

        self.allfilter = Gtk.FileFilter()
        self.allfilter.set_name("All files")
        self.allfilter.add_pattern("*")
        self.pdffilter = Gtk.FileFilter()
        self.pdffilter.set_name("PDF files")
        self.pdffilter.add_pattern("*.pdf")
        self.imgfilter = Gtk.FileFilter()
        self.imgfilter.set_name("Images")
        self.imgfilter.add_mime_type("image/png")
        self.imgfilter.add_mime_type("image/jpeg")
        self.imgfilter.add_mime_type("image/gif")
        self.imgfilter.add_pattern("*.png")
        self.imgfilter.add_pattern("*.jpg")
        self.imgfilter.add_pattern("*.gif")
        self.imgfilter.add_pattern("*.tif")
        self.imgfilter.add_pattern("*.xpm")

    # Backup thread specializzato
    def __getBackupThread(self, preferences, history):
        return BackupThread(preferences, history)

    def __buildTabacchi(self, builder):
        self.numRivenditaEntry = builder.get_object("numRivenditaEntry")
        self.codClienteEntry = builder.get_object("codClienteEntry")
        self.cittaEntry = builder.get_object("cittaEntry")
        self.nomeEntry = builder.get_object("nomeEntry")
        self.cognomeEntry = builder.get_object("cognomeEntry")
        self.telefonoEntry = builder.get_object("telefonoEntry")
        self.levataCombobox = builder.get_object("levataCombobox")
        self.timbrofchooser = builder.get_object("timbroFilechooserbutton")
        self.firmafchooser = builder.get_object("firmaFilechooserbutton")
        self.u88fchooser = builder.get_object("u88Filechooserbutton")
        self.u88urgfchooser = builder.get_object("u88urgFilechooserbutton")
        timbroBox = builder.get_object("timbroBox")
        firmaBox = builder.get_object("firmaBox")
        self.giorniConsegnaCombobox = builder.get_object("giorniConsegnaCombobox")
        self.ordineEntroCombobox = builder.get_object("ordineEntroCombobox")
        self.consegneFrame = builder.get_object("consegneFrame")
        self.pianoConsegneDaSitoSwitch = builder.get_object("pianoConsegneDaSitoSwitch")

        self.firmaWEntry = utility.NumEntry(firmaBox, 1, 2, 2)
        self.firmaHEntry = utility.NumEntry(firmaBox, 3, 2, 2)
        self.timbroWEntry = utility.NumEntry(timbroBox, 1, 2, 2)
        self.timbroHEntry = utility.NumEntry(timbroBox, 3, 2, 2)

        self.tabacchiUserEntry = builder.get_object("tabacchiUserEntry")
        self.catalogoUrlEntry = builder.get_object("catalogoUrlEntry")
        self.loginUrlEntry = builder.get_object("loginUrlEntry")
        self.tabacchiPwdEntry = builder.get_object("tabacchiPwdEntry")

        self.dateModel = Gtk.ListStore(str)
        self.dateModel.append(["Lunedì"])
        self.dateModel.append(["Martedì"])
        self.dateModel.append(["Mercoledì"])
        self.dateModel.append(["Giovedì"])
        self.dateModel.append(["Venerdì"])

        self.giorniConsegnaModel = Gtk.ListStore(str)
        for x in range(0, 6):
            self.giorniConsegnaModel.append([str(x)])

        self.ordiniEntroModel = Gtk.ListStore(str)
        for x in range(0, 24):
            self.ordiniEntroModel.append([str(x)])

        self.numRivenditaEntry.set_text(prefs.numRivendita)
        self.codClienteEntry.set_text(prefs.codCliente)
        self.cittaEntry.set_text(prefs.citta)
        self.nomeEntry.set_text(prefs.nome)
        self.cognomeEntry.set_text(prefs.cognome)
        self.telefonoEntry.set_text(prefs.telefono)
        self.pianoConsegneDaSitoSwitch.set_active(prefs.pianoConsegneDaSito)
        self.consegneFrame.set_sensitive(not prefs.pianoConsegneDaSito)

        self.levataCombobox.set_model(self.dateModel)
        self.giorniConsegnaCombobox.set_model(self.giorniConsegnaModel)
        self.ordineEntroCombobox.set_model(self.ordiniEntroModel)
        if prefs.timbro:
            self.timbrofchooser.set_filename(prefs.timbro)
        if prefs.firma:
            self.firmafchooser.set_filename(prefs.firma)
        if prefs.u88:
            self.u88fchooser.set_filename(prefs.u88)
        if prefs.u88urg:
            self.u88urgfchooser.set_filename(prefs.u88urg)

        self.firmaWEntry.set_value(prefs.firmaW)
        self.firmaHEntry.set_value(prefs.firmaH)
        self.timbroWEntry.set_value(prefs.timbroW)
        self.timbroHEntry.set_value(prefs.timbroH)

        cell = Gtk.CellRendererText()
        self.levataCombobox.pack_start(cell, True)
        self.levataCombobox.add_attribute(cell, 'text', 0)
        self.levataCombobox.set_active(prefs.giornoLevata)

        cell = Gtk.CellRendererText()
        self.giorniConsegnaCombobox.pack_start(cell, True)
        self.giorniConsegnaCombobox.add_attribute(cell, 'text', 0)
        self.giorniConsegnaCombobox.set_active(prefs.ggPerOrdine)

        cell = Gtk.CellRendererText()
        self.ordineEntroCombobox.pack_start(cell, True)
        self.ordineEntroCombobox.add_attribute(cell, 'text', 0)
        self.ordineEntroCombobox.set_active(prefs.oraInvio)

        self.numRivenditaEntry.set_text(prefs.numRivendita)
        self.codClienteEntry.set_text(prefs.codCliente)

        self.catalogoUrlEntry.set_text(prefs.catalogoUrl)
        self.loginUrlEntry.set_text(prefs.loginUrl)

        self.tabacchiUserEntry.set_text(prefs.tabacchiUser)
        self.tabacchiPwdEntry.set_text(prefs.tabacchiPwd)

        self.pianoConsegneDaSitoSwitch.connect("notify::active", self.consegneToggled)

    #
    def consegneToggled(self, switch, gparam):
        toggled = switch.get_active()
        self.consegneFrame.set_sensitive(not toggled)

    def refreshBarcode(self, widget):
        initBarcodeThread = InitBarcodeThread()
        progressDialog = utility.ProgressDialog(self.preferencesDialog, "Searching bluetooth devices..", "", "RFCOMM Bluetooth devices", initBarcodeThread)
        progressDialog.setResponseCallback(self.__loadBarcode)
        progressDialog.setStopCallback(self.barcodeModel.clear)
        progressDialog.setErrorCallback(self.barcodeModel.clear)
        progressDialog.startPulse()

    def __loadBarcode(self, barcodeList):
        self.barcodeModel.clear()
        for row in barcodeList:
            self.barcodeModel.append(row)
        self.barcodeCombobox.set_active(0)

    def __buildBarcode(self, builder):
        self.barcodeCombobox = builder.get_object("barcodeCombobox")
        self.barcodeModel = Gtk.ListStore(str, str, int)
        self.barcodeCombobox.set_model(self.barcodeModel)
        cell = Gtk.CellRendererText()
        self.barcodeCombobox.pack_start(cell, True)
        self.barcodeCombobox.add_attribute(cell, 'text', 0)

    def check(self, widget, other=None):
        utility.PreferencesDialog.check(self, widget, other)

    def save(self):
        prefs.numRivendita = self.numRivenditaEntry.get_text()
        prefs.codCliente = self.codClienteEntry.get_text()
        prefs.citta = self.cittaEntry.get_text()
        prefs.nome = self.nomeEntry.get_text()
        prefs.cognome = self.cognomeEntry.get_text()
        prefs.telefono = self.telefonoEntry.get_text()
        prefs.pianoConsegneDaSito = self.pianoConsegneDaSitoSwitch.get_active()
        prefs.oraInvio = self.ordineEntroCombobox.get_active()
        prefs.giornoLevata = self.levataCombobox.get_active()
        prefs.ggPerOrdine = self.giorniConsegnaCombobox.get_active()

        # Se il primo operatore e' None restituisce il secondo (False or True = True)
        prefs.timbro = self.timbrofchooser.get_filename() or ''
        prefs.firma = self.firmafchooser.get_filename() or ''
        prefs.u88 = self.u88fchooser.get_filename() or ''
        prefs.u88urg = self.u88urgfchooser.get_filename() or ''

        prefs.firmaW = self.firmaWEntry.get_value()
        prefs.firmaH = self.firmaHEntry.get_value()
        prefs.timbroW = self.timbroWEntry.get_value()
        prefs.timbroH = self.timbroHEntry.get_value()

        prefs.catalogoUrl = self.catalogoUrlEntry.get_text()
        prefs.loginUrl = self.loginUrlEntry.get_text()
        prefs.tabacchiUser = self.tabacchiUserEntry.get_text()
        prefs.tabacchiPwd = self.tabacchiPwdEntry.get_text()

        prefs.barcodeList[:] = []
        for row in self.barcodeModel:
            prefs.barcodeList.append(list(row))

        prefs.defaultBarcode = self.barcodeCombobox.get_active()

        super().save()


# Effettua il backup del programma su Cloud esteso
class BackupThread(utility.BackupThread):
    def __init__(self, preferences, history=None):
        super().__init__(preferences, history)
    #

    def __dataBackup(self, workdir, backupDir):
        dataPathName = f"{workdir}/{config.__desc__.replace(' ', '_')}"
        command = f"tar cjfP {dataPathName}.tar.bz2 --exclude=.* {backupDir}"
        subprocess.check_call(command, shell=True)
        return dataPathName
