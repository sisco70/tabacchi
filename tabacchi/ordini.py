#!/usr/bin/python
# coding: UTF-8
#
# Copyright (C) Francesco Guarnieri 2020 <francesco@guarnie.net>
#

import locale
import sqlite3
import datetime

from playsound import playsound
import gi

import config
from config import log
import utility
import stats
from preferencesTabacchi import prefs
import preferencesTabacchi

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk   # noqa: E402

BEEP_SOUND = str(config.RESOURCE_PATH / 'beep.ogg')
ERROR_SOUND = str(config.RESOURCE_PATH / 'error.ogg')

EDIT_MODE, VIEW_MODE, REVIEW_MODE, NEW_MODE = (0, 1, 2, 3)

MODIFICABILE, IN_LAVORAZIONE, EVASO = ("Modificabile", "In lavorazione", "Evaso")
STATI_VALIDI = [MODIFICABILE, IN_LAVORAZIONE, EVASO]

IN_CORSO, RICEVUTO, INVIATO = (0, 1, 2)
STATO_ORDINE = {IN_CORSO: "In corso", RICEVUTO: "Ricevuto", INVIATO: "Inviato"}

ORDINARIO, URGENTE, STRAORDINARIO = (0, 1, 2)
TIPO_ORDINE = {ORDINARIO: "", URGENTE: "Urgente", STRAORDINARIO: "Straordinario"}

# Dialog per verificare un ordine ricevuto


class RicezioneOrdineDialog(utility.GladeWindow):
    ID, DESCRIZIONE, PESO, CARICO, COSTO, UNITA_MIN, VERIFICA, PREZZO_KG, ELIMINATO, BARCODE = (0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    modelInfoList = [
        ("Verifica", "bool", VERIFICA),
        ("Codice", "str", ID),
        ("!+Descrizione", "str", DESCRIZIONE),
        ("*In Carico", "float#3,3/i5,0,12", CARICO),
        ("Ordine", "float", PESO),
        ("Costo", "currency", COSTO),
        (None, "float", UNITA_MIN),
        (None, "currency", PREZZO_KG),
        (None, "bool", ELIMINATO),
        (None, "str", BARCODE)]

    def __init__(self, parent, idOrdine, data):
        super().__init__(parent, "verificaOrdineDialog.glade")
        self.idOrdine = idOrdine
        self.fullView = False
        self.dirty = False
        self.readBarcodeThread = None
        self.totCarico = 0
        self.totPeso = 0
        self.totEuroCarico = 0
        self.totEuroPeso = 0
        barcode = prefs.barcodeList[prefs.defaultBarcode]
        self.barcodeDevice = barcode[0]
        self.barcodeAddr = barcode[1]
        self.barcodePort = barcode[2]

        self.verificaOrdineDialog = self.builder.get_object("verificaOrdineDialog")
        self.costoTotaleLabel = self.builder.get_object("costoTotaleLabel")
        self.pesoTotaleLabel = self.builder.get_object("pesoTotaleLabel")
        self.deleteMenuitem = self.builder.get_object("deletemenuitem")
        self.allineaMenuitem = self.builder.get_object("allineamenuitem")
        self.verificaOrdineDialog.set_transient_for(parent)
        self.bluetoothStatusImage = self.builder.get_object('bluetoothStatusImage')
        self.bluetoothStatusImage.hide()
        self.ordineTreeview = self.builder.get_object("ordineTreeView")
        self.verificaPopupMenu = self.builder.get_object("verificaPopupMenu")

        listino_formats = {self.PESO: "%.3f kg"}
        listino_properties = {self.ID: {"xalign": 1, "scale": utility.PANGO_SCALE_SMALL}, self.COSTO: {"xalign": 1, "scale": utility.PANGO_SCALE_SMALL}}

        utility.ExtTreeView(self.modelInfoList, self.ordineTreeview, formats=listino_formats,
                            edit_callbacks={self.CARICO: self.caricoCallback}, properties=listino_properties)
        self.ordineModel = self.ordineTreeview.get_model()

        self.ordineTreeview.connect("button-press-event", self.showPopup, self.verificaPopupMenu)

        self.ordineDict = dict()
        self.listinoDict = dict()
        self.deletedList = []
        self.ordineTreeview.set_model(None)
        self.load(self.ordineModel)
        self.ordineTreeview.set_model(self.ordineModel)
        self.editable = None
        self.result = None
        self.selectedPath = None
        self.verificaOrdineDialog.set_title(data.strftime("Ricezione ordine di %A %d %B %Y"))

        self.builder.connect_signals({"on_verificaOrdineDialog_delete_event": self.close,
                                      "on_verificaOrdineDialog_destroy_event": self.close,
                                      "on_barcodeToolbutton_clicked": self.enableBarcode,
                                      "on_deletemenuitem_activate": self.deleteArticolo,
                                      "on_allineamenuitem_activate": self.changePeso,
                                      "on_okButton_clicked": self.okClose,
                                      "on_cancelButton_clicked": self.close
                                      })

    # Mostra menu popup per la gestione ordini
    def showPopup(self, treeview, event, popupMenu):
        if event.button == 3:
            x = int(event.x)
            y = int(event.y)
            pthinfo = treeview.get_path_at_pos(x, y)
            if pthinfo is not None:
                path, col, _, _ = pthinfo
                self.selectedPath = path
                selection = treeview.get_selection()
                self.checkAllinea(path)
                if not selection.path_is_selected(path):
                    treeview.set_cursor(path, col, 0)
                treeview.grab_focus()
                popupMenu.popup(None, None, None, None, event.button, event.time)
            return True

    def enableBarcode(self, widget):
        if self.readBarcodeThread:
            self.readBarcodeThread.stop()
            self.readBarcodeThread = None
            self.bluetoothStatusImage.hide()
        else:
            connectBarcodeThread = preferencesTabacchi.ConnectBarcodeThread(self.barcodeAddr, self.barcodePort)
            progressDialog = utility.ProgressDialog(self.verificaOrdineDialog, "Connecting to device %s.." %
                                                    self.barcodeDevice, "", "Bluetooth Barcode reader", connectBarcodeThread)
            progressDialog.setResponseCallback(self.connectCallback)
            progressDialog.setErrorCallback(self.bluetoothStatusImage.hide)
            progressDialog.startPulse()

    def connectCallback(self, sock):
        if sock:
            self.bluetoothStatusImage.show()
            self.readBarcodeThread = preferencesTabacchi.ReadBarcodeThread(self.readDataCallback, self.errorCallback, sock)
            self.readBarcodeThread.start()
        else:
            self.bluetoothStatusImage.hide()

    def errorCallback(self, e):
        self.bluetoothStatusImage.hide()
        utility.gtkErrorMsg(e, self.verificaOrdineDialog)

    # Metodo che viene invocato dal thread di comunicazione bluetooth ogni volta che si legge un codice
    def readDataCallback(self, data):
        if data in self.ordineDict:
            playsound(BEEP_SOUND)
            path = self.ordineDict[data]
            selection = self.ordineTreeview.get_selection()
            selection.select_path(path)
            self.ordineTreeview.scroll_to_cell(path)
            carico = self.ordineModel[path][self.CARICO] + self.ordineModel[path][self.UNITA_MIN]
            self.__setValue(path, round(carico, 3), self.ordineModel)
        elif data in self.listinoDict:
            playsound(ERROR_SOUND)
            descrizione = self.listinoDict[data][1]
            msgDialog = Gtk.MessageDialog(parent=self.verificaOrdineDialog, flags=Gtk.DialogFlags.MODAL, type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.YES_NO, message_format="Articolo non presente nell'ordine.")
            msgDialog.format_secondary_text("Aggiungere %s?" % descrizione)
            msgDialog.set_title("Attenzione")
            result = msgDialog.run()
            msgDialog.destroy()
            if result == Gtk.ResponseType.YES:
                self.dirty = True
                iterator = self.ordineModel.append()
                unitaMin = self.listinoDict[data][2]
                prezzoKg = self.listinoDict[data][3]
                costo = round(prezzoKg * unitaMin, 3)
                self.ordineModel.set_value(iterator, self.ID, self.listinoDict[data][0])
                self.ordineModel.set_value(iterator, self.DESCRIZIONE, descrizione)
                self.ordineModel.set_value(iterator, self.PESO, unitaMin)
                self.ordineModel.set_value(iterator, self.COSTO, costo)
                self.ordineModel.set_value(iterator, self.UNITA_MIN, unitaMin)
                self.ordineModel.set_value(iterator, self.CARICO, unitaMin)
                self.ordineModel.set_value(iterator, self.VERIFICA, True)
                self.ordineModel.set_value(iterator, self.PREZZO_KG, prezzoKg)
                self.ordineModel.set_value(iterator, self.BARCODE, data)
                self.totCarico += unitaMin
                self.totPeso += unitaMin
                self.totEuroCarico += costo
                self.totEuroPeso += costo
                self.updateLabels()
                del self.listinoDict[data]  # elimina il valore dal dizionario del listino, avendolo messo nell'ordine..
                self.ordineDict[data] = self.ordineModel.get_path(iterator)
        else:
            playsound(ERROR_SOUND)
            msgDialog = Gtk.MessageDialog(parent=self.verificaOrdineDialog, flags=Gtk.DialogFlags.MODAL, type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.CANCEL, message_format="Codice a barre non riconosciuto: %s" % data)
            msgDialog.format_secondary_text("Devi associarlo ad un articolo.")
            msgDialog.set_title("Attenzione")
            result = msgDialog.run()
            msgDialog.destroy()

    def caricoCallback(self, widget, path, value, model, col_id):
        self.__setValue(path, value, model)

    def updateLabels(self):
        self.pesoTotaleLabel.set_text("%skg / %skg" % (locale.format_string("%.3f", self.totCarico), locale.format_string("%.3f", self.totPeso)))
        self.costoTotaleLabel.set_text("%s / %s" % (locale.currency(self.totEuroCarico, True, True), locale.currency(self.totEuroPeso, True, True)))

    def __setValue(self, path, value, model):
        row = model[path]
        peso = row[self.PESO]
        carico = row[self.CARICO]
        old_costo = round(carico * row[self.PREZZO_KG], 3)
        costo = round(value * row[self.PREZZO_KG], 3)
        if value <= peso:
            self.dirty = True
            self.totCarico = self.totCarico - carico + value
            self.totEuroCarico = self.totEuroCarico - old_costo + costo
            row[self.CARICO] = value
            row[self.VERIFICA] = (value == peso)
        else:
            playsound(ERROR_SOUND)
            msgDialog = Gtk.MessageDialog(parent=self.verificaOrdineDialog, flags=Gtk.DialogFlags.MODAL, type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.YES_NO, message_format="Quantità consegnata superiore all'ordine.")
            msgDialog.format_secondary_text("Incremento l'ordine?")
            msgDialog.set_title("Attenzione")
            result = msgDialog.run()
            msgDialog.destroy()
            if result == Gtk.ResponseType.YES:
                self.dirty = True
                self.totCarico = self.totCarico - carico + value
                self.totPeso = self.totPeso - peso + value
                self.totEuroCarico = self.totEuroCarico - old_costo + costo
                self.totEuroPeso = self.totEuroPeso - old_costo + costo
                row[self.CARICO] = value
                row[self.PESO] = value
                row[self.COSTO] = costo
                row[self.VERIFICA] = True
        self.updateLabels()

    def deleteArticolo(self, widget):
        model = self.ordineModel
        row = model[self.selectedPath]
        msgDialog = Gtk.MessageDialog(parent=self.verificaOrdineDialog, flags=Gtk.DialogFlags.MODAL, type=Gtk.MessageType.WARNING,
                                      buttons=Gtk.ButtonsType.YES_NO, message_format="Elimino dall'ordine l'articolo %s?" % row[self.DESCRIZIONE])
        msgDialog.set_title("Attenzione")
        result = msgDialog.run()
        msgDialog.destroy()
        if result == Gtk.ResponseType.YES:
            self.dirty = True
            carico = row[self.CARICO]
            peso = row[self.PESO]
            prezzoKg = row[self.PREZZO_KG]
            idOrdine = row[self.ID]
            barcode = row[self.BARCODE]
            self.totCarico -= carico
            self.totPeso -= peso
            self.totEuroCarico -= round(carico * prezzoKg, 3)
            self.totEuroPeso -= round(peso * prezzoKg, 3)
            # Aggiunge al dizionario del listino l'articolo eliminato
            if barcode and (len(barcode) > 0):
                self.listinoDict[barcode] = [idOrdine, row[self.DESCRIZIONE], row[self.UNITA_MIN], prezzoKg]
            iterator = model.get_iter(self.selectedPath)
            model.remove(iterator)
            self.deletedList.append(idOrdine)
            # Aggiorna tutti i riferimenti ai path del dizionario dell'ordine
            self.ordineDict.clear()
            iterator = model.get_iter_first()
            while iterator:
                barcode = model.get_value(iterator, self.BARCODE)
                if barcode and (len(barcode) > 0):
                    self.ordineDict[barcode] = model.get_path(iterator)
                iterator = model.iter_next(iterator)
            self.updateLabels()

    def checkAllinea(self, path):
        row = self.ordineModel[path]
        self.allineaMenuitem.set_sensitive(not row[self.VERIFICA] and (row[self.CARICO] > 0))

    def changePeso(self, widget):
        row = self.ordineModel[self.selectedPath]
        carico = row[self.CARICO]
        msgDialog = Gtk.MessageDialog(parent=self.verificaOrdineDialog, flags=Gtk.DialogFlags.MODAL, type=Gtk.MessageType.QUESTION,
                                      buttons=Gtk.ButtonsType.YES_NO, message_format="Allineo l'ordine alla quantità consegnata?")
        msgDialog.set_title("Attenzione")
        result = msgDialog.run()
        msgDialog.destroy()
        if result == Gtk.ResponseType.YES:
            self.dirty = True
            peso = row[self.PESO]
            old_costo = row[self.COSTO]
            row[self.PESO] = carico
            row[self.COSTO] = round(carico * row[self.PREZZO_KG], 3)
            row[self.VERIFICA] = True
            self.totPeso = self.totPeso - peso + carico
            self.totEuroPeso = self.totEuroPeso - old_costo + row[self.COSTO]
            self.updateLabels()

    def save(self, model):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            for _id in self.deletedList:
                cursor.execute("update verificaOrdine set Carico=?, Peso=?, Eliminato=? where ID=? and ID_Ordine=?", (0, 0, True, _id, self.idOrdine))
                if cursor.rowcount == 0:
                    cursor.execute("INSERT INTO verificaOrdine (Carico, Peso, Eliminato, ID, ID_Ordine) VALUES (?, ?, ?, ?, ?)", (0, 0, True, _id, self.idOrdine))
            for row in model:
                riga = (row[self.CARICO], row[self.PESO], row[self.ELIMINATO], row[self.ID], self.idOrdine)
                cursor.execute("update verificaOrdine set Carico=?, Peso=?, Eliminato=? where ID=? and ID_Ordine=?", riga)
                if cursor.rowcount == 0:
                    cursor.execute("INSERT INTO verificaOrdine (Carico, Peso, Eliminato, ID, ID_Ordine) VALUES (?, ?, ?, ?, ?)", riga)
            conn.commit()
        except sqlite3.Error as e:
            if conn:
                conn.rollback()
            utility.gtkErrorMsg(e, self.verificaOrdineDialog)
        else:
            self.dirty = False
            prefs.setDBDirty()
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def load(self, model):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute("SELECT count(ID) as size from verificaOrdine where ID_Ordine = ?", (self.idOrdine,))
            row = cursor.fetchone()
            if row["size"] > 0:
                cursor.execute(
                    "SELECT t.Descrizione, t.ID, t.barcode, v.Peso as Ordine, v.Carico, t.unitaMin, t.prezzoKg, v.Eliminato FROM (tabacchi t LEFT JOIN (select * from verificaOrdine where ID_Ordine = ?) as v on v.ID = t.ID) order by t.Tipo desc, t.Descrizione",
                    (self.idOrdine,))
            else:
                # Inizializza la tabella per la verifica
                cursor.execute(
                    "INSERT INTO verificaOrdine (ID, ID_Ordine, Carico, Peso, Eliminato) SELECT t.ID, r.ID_Ordine, 0, r.Ordine, 0 FROM (tabacchi t JOIN (select * from rigaOrdineTabacchi where ID_Ordine = ?) as r on r.ID = t.ID)",
                    (self.idOrdine,))
                conn.commit()
                # Legge i dati per popolare la treeview
                cursor.execute(
                    "SELECT t.Descrizione, t.ID, t.barcode, r.Ordine, 0 as Carico, t.unitaMin, t.prezzoKg, 0 as Eliminato FROM (tabacchi t LEFT JOIN (select * from rigaOrdineTabacchi where ID_Ordine = ?) as r on r.ID = t.ID) order by t.Tipo desc, t.Descrizione",
                    (self.idOrdine,))

            resultList = cursor.fetchall()

            model.clear()
            self.ordineDict.clear()
            self.listinoDict.clear()
            self.totPeso = 0
            for row in resultList:
                peso = row["Ordine"]
                barcode = row["barcode"]
                eliminato = row["Eliminato"]
                idOrdine = row["ID"]
                prezzoKg = round(row["prezzoKg"], 3) if row["prezzoKg"] else 0
                descrizione = row["Descrizione"]
                unitaMin = round(row["unitaMin"], 3) if row["unitaMin"] else 0
                if not eliminato and peso and (peso > 0):
                    peso = round(peso, 3)
                    carico = round(row["Carico"], 3) if row["Carico"] else 0
                    costo = round(prezzoKg * peso, 3)
                    self.totPeso += peso
                    self.totCarico += carico
                    self.totEuroCarico += round(prezzoKg * carico, 3)
                    self.totEuroPeso += costo
                    iterator = model.append()
                    model.set_value(iterator, self.ID, idOrdine)
                    model.set_value(iterator, self.DESCRIZIONE, descrizione)
                    model.set_value(iterator, self.PESO, peso)
                    model.set_value(iterator, self.COSTO, costo)
                    model.set_value(iterator, self.UNITA_MIN, unitaMin)
                    model.set_value(iterator, self.CARICO, carico)
                    model.set_value(iterator, self.VERIFICA, carico == peso)
                    model.set_value(iterator, self.PREZZO_KG, prezzoKg)
                    model.set_value(iterator, self.BARCODE, barcode)
                    if barcode and (len(barcode) > 0):
                        self.ordineDict[barcode] = model.get_path(iterator)
                else:
                    if eliminato:
                        self.deletedList.append(idOrdine)
                    if barcode and (len(barcode) > 0):
                        self.listinoDict[barcode] = [idOrdine, descrizione, unitaMin, prezzoKg]
        except sqlite3.Error as e:
            if conn:
                conn.rollback()
            utility.gtkErrorMsg(e, self.verificaOrdineDialog)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
            self.updateLabels()

    def run(self):
        self.verificaOrdineDialog.run()

    # Aggiorna l'ordine inviato con quello che si è realmente ricevuto
    def __updateOrdine(self, model):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            for idDeleted in self.deletedList:
                cursor.execute("update rigaOrdineTabacchi set Ordine=? where ID=? and ID_Ordine=?", (0, idDeleted, self.idOrdine))
            for row in model:
                cursor.execute("update rigaOrdineTabacchi set Ordine=? where ID=? and ID_Ordine=?", (row[self.PESO], row[self.ID], self.idOrdine))
                if cursor.rowcount == 0:
                    cursor.execute("insert into rigaOrdineTabacchi (ID, Descrizione, ID_Ordine, Ordine, Prezzo) values (?, ?, ?, ?, ?)",
                                   (row[self.ID], row[self.DESCRIZIONE], self.idOrdine, row[self.PESO], row[self.PREZZO_KG]))
            cursor.execute("update ordineTabacchi set Stato = ? where ID = ?", (RICEVUTO, self.idOrdine))
            cursor.execute("delete from verificaOrdine where ID_Ordine = ?", (self.idOrdine,))
            conn.commit()
        except sqlite3.Error as e:
            if conn:
                conn.rollback()
            utility.gtkErrorMsg(e, self.verificaOrdineDialog)
        else:
            self.dirty = False
            prefs.setDBDirty()
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def okClose(self, widget):
        self.save(self.ordineModel)
        allVerified = True
        for row in self.ordineModel:
            if not row[self.VERIFICA]:
                allVerified = False
                break
        if allVerified and (round(self.totCarico, 3) == round(self.totPeso, 3)) and (round(self.totEuroCarico, 2) == round(self.totEuroPeso, 2)):
            msgDialog = Gtk.MessageDialog(parent=self.verificaOrdineDialog, flags=Gtk.DialogFlags.MODAL, type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.YES_NO, message_format="Gli articoli sono stati verificati e l'importo ed il peso corrispondono.")
            msgDialog.format_secondary_text("Procedo con l'archiviazione?")
            msgDialog.set_title("Attenzione")
            result = msgDialog.run()
            msgDialog.destroy()
            if (result == Gtk.ResponseType.YES):
                self.__updateOrdine(self.ordineModel)

        self.close(self, widget)

    def close(self, widget, other=None):
        reallyClose = True
        if self.dirty:
            msgDialog = Gtk.MessageDialog(parent=self.verificaOrdineDialog, flags=Gtk.DialogFlags.MODAL, type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.YES_NO, message_format="Saranno perse le verifiche effettuate.")
            msgDialog.format_secondary_text("Sei sicuro?")
            msgDialog.set_title("Attenzione")
            reallyClose = (msgDialog.run() == Gtk.ResponseType.YES)
            msgDialog.destroy()
        if reallyClose:
            if self.readBarcodeThread:
                log.debug("Closing Verifica Ordine Dialog: Stop readBarcodeThread..")
                self.readBarcodeThread.stop()
            self.verificaOrdineDialog.destroy()
            return False
        else:
            return True

#


class SuppletivoDialog(utility.GladeWindow):
    def __init__(self, parent, idOrdine, tipo, data=None):
        super().__init__(parent, "suppletivoDialog.glade")
        self.idOrdine = idOrdine
        self.dirtyFlag = False

        self.suppletivoDialog = self.builder.get_object("suppletivoDialog")
        self.suppletivoDialog.set_transient_for(parent)
        self.tipoCombobox = self.builder.get_object("tipoCombobox")
        self.ordineTreeview = self.builder.get_object('tabacchiTreeView')
        self.totaleLabel = self.builder.get_object('totaleLabel')

        self.__loadCombo(tipo)

        modelInfoList = [
            ("Codice", "str"),
            ("+Descrizione", "str"),
            ("Tipo", "str"),
            ("*Quantità", "float#3,3/i6,0,12"),
            ("Prezzo conf.", "currency"),
            ("Importo", "currency"),
            (None, "float"),
            (None, "float")]
        suppletivoProp = {2: {"xalign": 0.5, "scale": utility.PANGO_SCALE_SMALL}}
        utility.ExtTreeView(modelInfoList, self.ordineTreeview, edit_callbacks={3: self.__changeValue}, properties=suppletivoProp)
        self.ordineModel = self.ordineTreeview.get_model()
        self.dataEntry = utility.DataEntry(self.builder.get_object("dataButton"), data, " %d %B %Y ", self.__update)
        self.suppletivoDialog.set_title("Ordine suppletivo")

        self.__loadModelFromDB()
        self.tipoCombobox.connect("changed", self.__update)
        self.totaleLabel.set_text("Peso: %s kg   Importo: %s" % (locale.format_string("%.3f", self.totaleKg), locale.currency(self.totale, True, True)))

    def __update(self, widget=None):
        self.dirtyFlag = True

    def __loadCombo(self, tipo):
        # Nel caso stessimo inserendo un ordine suppletivo per la prima volta
        # il tipo sarebbe ordinario
        if tipo == ORDINARIO:
            self.dirtyFlag = True
            tipo = URGENTE  # Valore default
        tipo_store = Gtk.ListStore(int, str)
        cell = Gtk.CellRendererText()
        self.tipoCombobox.pack_start(cell, True)
        self.tipoCombobox.add_attribute(cell, 'text', 1)
        i = 0
        index = 0
        for cod in TIPO_ORDINE.keys():
            if cod == tipo:
                index = i
            if cod != ORDINARIO:
                tipo_store.append([cod, TIPO_ORDINE[cod]])
                i += 1
        self.tipoCombobox.set_model(tipo_store)
        self.tipoCombobox.set_active(index)

    def __changeValue(self, widget, path, value, model, col_id):
        row = model[path]
        self.dirtyFlag = True
        oldQuantita = row[3]
        prezzoKg = row[7]
        quantita = value
        importo = quantita * prezzoKg
        self.totale = self.totale - (oldQuantita * prezzoKg) + importo
        self.totaleKg = self.totaleKg - oldQuantita + quantita
        row[3] = quantita
        row[5] = importo
        self.totaleLabel.set_text(f"Peso: {locale.format_string('%.3f', self.totaleKg)} kg   Importo: {locale.currency(self.totale, True, True)}")

    def __loadModelFromDB(self):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute(
                "SELECT t.ID, t.Descrizione, t.PrezzoKg, t.Tipo, t.UnitaMin, r.Ordine FROM (tabacchi t LEFT JOIN (select * from rigaOrdineSuppletivo where ID_Ordine = ?) as r on r.ID = t.ID) where t.InMagazzino order by t.Tipo, t.Descrizione",
                (self.idOrdine,))
            result_set = cursor.fetchall()
            self.ordineModel.clear()
            self.totale = 0
            self.totaleKg = 0
            for row in result_set:
                idArt = row["ID"]
                unitaMin = row["UnitaMin"]
                prezzoKg = row["PrezzoKg"]
                quantita = row["Ordine"]
                if quantita is None:
                    quantita = 0
                self.ordineModel.append([idArt, row["Descrizione"], row["Tipo"], quantita, prezzoKg * unitaMin, quantita * prezzoKg, unitaMin, prezzoKg])
                self.totaleKg += quantita
                self.totale += (quantita * prezzoKg)
        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self.suppletivoDialog)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def __saveModelToDB(self):
        data = self.dataEntry.data
        tipo = URGENTE
        iterator = self.tipoCombobox.get_active_iter()
        if iterator:
            tipo = self.tipoCombobox.get_model()[iterator][0]
        if data is None:
            msgDialog = Gtk.MessageDialog(parent=self.suppletivoDialog, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT,
                                          type=Gtk.MessageType.WARNING, buttons=Gtk.ButtonsType.OK, message_format="La data di consegna è obbligatoria.")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()
            return False
        elif self.totaleKg <= 0:
            msgDialog = Gtk.MessageDialog(parent=self.suppletivoDialog, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT,
                                          type=Gtk.MessageType.WARNING, buttons=Gtk.ButtonsType.OK, message_format="L'ordine suppletivo è vuoto.")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()
            return False
        else:
            cursor = None
            conn = None

            try:
                conn = prefs.getConn()
                cursor = prefs.getCursor(conn)
                cursor.execute("delete from rigaOrdineSuppletivo where ID_Ordine = ?", (self.idOrdine,))
                for row in self.ordineModel:
                    quantita = round(row[3], 3)  # Per evitare prob. con gli arrotondamenti
                    if quantita > 0:
                        _id = row[0]
                        descrizione = row[1]
                        cursor.execute(
                            "INSERT INTO rigaOrdineSuppletivo(ID, Descrizione, ID_Ordine, Ordine) VALUES(?, ?, ?, ?)",
                            (_id, descrizione, self.idOrdine, quantita))
                cursor.execute("Update ordineTabacchi set DataSuppletivo = ?, Suppletivo = ? where ID = ?", (data, tipo, self.idOrdine))
                conn.commit()
            except sqlite3.Error as e:
                if conn:
                    conn.rollback()
                utility.gtkErrorMsg(e, self.suppletivoDialog)
            else:
                prefs.setDBDirty()
                self.dirtyFlag = False
            finally:
                if cursor:
                    cursor.close()
                if conn:
                    conn.close()
        return True

    # RUN corretto per gestione DELETE
    def run(self):
        redo = True
        while redo:
            redo = False
            result = self.suppletivoDialog.run()
            log.debug("Return from suppletivoDialog.run() = %s" % result)
            doSave = (result == Gtk.ResponseType.OK)
            if self.dirtyFlag:
                if not doSave:
                    msgDialog = Gtk.MessageDialog(
                        parent=self.suppletivoDialog, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT, type=Gtk.MessageType.QUESTION,
                        buttons=Gtk.ButtonsType.YES_NO, message_format="Ci sono modifiche non salvate.")
                    msgDialog.format_secondary_text("Vuoi salvarle?")
                    msgDialog.set_title("Attenzione")
                    doSave = (Gtk.ResponseType.YES == msgDialog.run())
                    msgDialog.destroy()
                if doSave and not self.__saveModelToDB():
                    redo = True
        self.suppletivoDialog.destroy()


class OrdineDialog(utility.GladeWindow):
    ID_QUANTITA, ID_ORDINE, ID_COSTO, ID_CONSUMO = (0, 1, 2, 3)

    def __init__(self, parent, ordineID=None, mode=EDIT_MODE):
        super().__init__(parent, "ordineDialog.glade")
        self.mode = mode
        self.dirtyFlag = False
        self.tabacchiList = list()
        self.ordineDict = dict()
        self.ordinePrecDict = dict()
        self.mainIndex = 0
        self.mainTotal = 0
        self.ordineID = ordineID

        self.error = True
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            # Legge la lista dei record dei tabacchi
            cursor.execute(
                "SELECT t.ID, t.Descrizione, t.UnitaMin, t.PrezzoKg, t.PezziUnitaMin, t.LivelloMin, t.Tipo FROM tabacchi t where t.InMagazzino order by t.Tipo desc, t.Descrizione")
            self.tabacchiList = cursor.fetchall()
            self.mainSize = len(self.tabacchiList)

            if (self.mainSize == 0):
                raise UserWarning()

            # Se si sta creando un nuovo ordine..
            if not self.ordineID:
                dataOrdine = datetime.datetime.now().replace(microsecond=0)
                (_, levata) = preferencesTabacchi.dataLimiteOrdine(dataOrdine)
                log.debug("Nuovo ordine levata: %s" % levata)
                cursor.execute("insert into ordineTabacchi (Data, Levata) values (?, ?)", (dataOrdine, levata))
                self.ordineID = cursor.lastrowid
                conn.commit()
                prefs.setDBDirty()
            else:
                cursor.execute("SELECT ID, Ordine, Giacenza, Prezzo, Consumo FROM rigaOrdineTabacchi where ID_Ordine = ?", (self.ordineID,))
                resultset = cursor.fetchall()
                for row in resultset:
                    ordine = row["Ordine"]
                    costo = row["Prezzo"] * ordine
                    self.mainTotal += costo
                    self.ordineDict[row["ID"]] = [row["Giacenza"], ordine, costo, row["Consumo"]]

            cursor.execute("select Data as 'Data [timestamp]', Stato, LastPos from ordineTabacchi where ID = ?", (self.ordineID,))
            row = cursor.fetchone()
            self.date = row["Data"]
            self.stato = row["Stato"]
            self.mainIndex = row["LastPos"]

            # memorizza tutte le quantita e il tabacco ordinato nell'ordine precedente
            cursor.execute(
                "SELECT ID, Ordine, Giacenza FROM rigaOrdineTabacchi where ID_Ordine = (select ID from ordineTabacchi where Data = (select max(Data) from ordineTabacchi where Data < ? order by Data desc))",
                (self.date,))
            resultset = cursor.fetchall()
            for row in resultset:
                self.ordinePrecDict[row["ID"]] = [row["Giacenza"], row["Ordine"]]

        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self.parent)
            if conn:
                conn.rollback()
        except UserWarning:
            msgDialog = Gtk.MessageDialog(parent=self.parent, flags=Gtk.DialogFlags.MODAL, message_type=Gtk.MessageType.ERROR,
                                          buttons=Gtk.ButtonsType.OK, text="Non è possibile creare un ordine.")
            msgDialog.format_secondary_text("Caricare catalogo Logista e creare un listino.")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()
        else:
            self.error = False
            self.ordineDialog = self.builder.get_object("ordineDialog")
            self.tipoLabel = self.builder.get_object("tipoLabel")
            self.codiceLabel = self.builder.get_object("codiceLabel")
            self.descLabel = self.builder.get_object("descLabel")
            self.unitaminLabel = self.builder.get_object("unitaminLabel")
            self.pezziUnitaLabel = self.builder.get_object("pezziUnitaLabel")
            self.prezzoLabel = self.builder.get_object("prezzoLabel")
            self.prezzoConfLabel = self.builder.get_object("prezzoConfLabel")
            self.livelloLabel = self.builder.get_object("livelloLabel")
            self.consumoLabel = self.builder.get_object("consumoLabel")
            self.consumoMinLabel = self.builder.get_object("consumoMinLabel")
            self.consumoMaxLabel = self.builder.get_object("consumoMaxLabel")
            self.quantitaSpinbutton = self.builder.get_object("quantitaSpinbutton")
            self.ordineSpinbutton = self.builder.get_object("ordineSpinbutton")
            self.totaleLabel = self.builder.get_object("totaleLabel")
            self.counterLabel = self.builder.get_object("counterLabel")

            self.ordineDialog.set_transient_for(parent)

            self.builder.connect_signals({"on_ordineDialog_delete_event": self.close,
                                          "on_closeButton_clicked": self.close,
                                          "on_fwdToolbutton_clicked": self.forward,
                                          "on_rewToolbutton_clicked": self.rewind,
                                          "on_nxtToolbutton_clicked": self.next,
                                          "on_preToolbutton_clicked": self.previous,
                                          "on_statsToolbutton_clicked": self.showStats,
                                          "on_quantitaSpinbutton_value_changed": self.quantitaChange,
                                          "on_ordineSpinbutton_value_changed": self.ordineChange})
            self.quantitaSpinbutton.set_sensitive(mode != VIEW_MODE)
            self.ordineSpinbutton.set_sensitive(mode == EDIT_MODE)

            self.ordineDialog.set_title("Ordine %s" % self.date.strftime("%d %b %Y - %H:%M"))

            if self.mainSize > self.mainIndex:
                self.__showInfo(self.mainIndex)
            else:
                self.mainIndex = self.mainSize - 1
                self.__showInfo(self.mainIndex)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    # Calcola consumo per un articolo

    def __calcolaConsumo(self, idArt, quantita):
        precQuantita = 0
        precOrdine = 0
        if idArt in self.ordinePrecDict:
            precTuple = self.ordinePrecDict[idArt]
            precQuantita = precTuple[self.ID_QUANTITA]
            precOrdine = precTuple[self.ID_ORDINE]
        return (precQuantita + precOrdine) - quantita

    # Aggiorna il DB e il dizionario con le quantità e gli ordini modificati
    def __update(self):
        if self.dirtyFlag:
            self.dirtyFlag = False
            idArt = self.tabacchiList[self.mainIndex]["ID"]
            ordine = round(self.ordineSpinbutton.get_value(), 3)
            prezzo = self.tabacchiList[self.mainIndex]["PrezzoKg"]
            costo = ordine * prezzo
            quantita = round(self.quantitaSpinbutton.get_value(), 3)
            consumo = self.__calcolaConsumo(idArt, quantita)

            # se questo articolo esiste (ha quantità o ordine inseriti)
            # si toglie il vecchio costo dal totale
            if idArt in self.ordineDict:
                data = self.ordineDict[idArt]
                self.mainTotal -= data[self.ID_COSTO]

            self.ordineDict[idArt] = [quantita, ordine, costo, consumo]
            self.mainTotal += costo

            cursor = None
            conn = None
            try:
                conn = prefs.getConn()
                cursor = prefs.getCursor(conn)
                cursor.execute("UPDATE rigaOrdineTabacchi SET Ordine=?, Giacenza=?, Consumo=? WHERE ID = ? and ID_Ordine = ?",
                               (ordine, quantita, float(consumo), idArt, self.ordineID))
                if cursor.rowcount == 0:
                    descrizione = self.tabacchiList[self.mainIndex]["Descrizione"]
                    cursor.execute(
                        "INSERT INTO rigaOrdineTabacchi(ID, ID_Ordine, Descrizione, Ordine, Prezzo, Giacenza, Consumo) VALUES(?, ?, ?, ?, ?, ?, ?)",
                        (idArt, self.ordineID, descrizione, ordine, float(prezzo),
                         quantita, float(consumo)))

                conn.commit()
            except sqlite3.Error as e:
                if conn:
                    conn.rollback()
                utility.gtkErrorMsg(e, self.ordineDialog)
            else:
                prefs.setDBDirty()
            finally:
                if cursor:
                    cursor.close()
                if conn:
                    conn.close()

    # Mostra statistiche sull'articolo corrente
    def showStats(self, widget):
        idArt = self.tabacchiList[self.mainIndex]["ID"]
        statsDialog = stats.StatsDialog(self.ordineDialog, idArt)
        statsDialog.run()

    # Evento scatenato dalla pressione dello spinbutton quantita
    def quantitaChange(self, widget):
        log.debug("dirtyFlag = TRUE")
        self.dirtyFlag = True
        # aggiornamento campi ordineSpinButton e consumoLabel
        if self.mode == EDIT_MODE:
            self.ordineSpinbutton.set_value(self.tabacchiList[self.mainIndex]["LivelloMin"] - widget.get_value())
        idList = self.tabacchiList[self.mainIndex]["ID"]
        self.consumoLabel.set_text(locale.format_string("%.3f kg", self.__calcolaConsumo(idList, widget.get_value())))

    # Evento scatenato dalla pressione dello spinbutton ordine
    def ordineChange(self, widget):
        log.debug("dirtyFlag = TRUE")
        self.dirtyFlag = True
        prezzoKg = self.tabacchiList[self.mainIndex]["PrezzoKg"]
        costo = widget.get_value() * prezzoKg

        ordineId = self.tabacchiList[self.mainIndex]["ID"]
        if ordineId in self.ordineDict:
            data = self.ordineDict[ordineId]
            self.totaleLabel.set_text("%s su %s" % (locale.currency(costo, True, True), locale.currency(
                self.mainTotal - (data[self.ID_ORDINE] * prezzoKg) + costo, True, True)))
        else:
            self.totaleLabel.set_text(f"{locale.currency(costo, True, True)} su {locale.currency(self.mainTotal + costo, True, True)}")

    def rewind(self, widget):
        self.__update()
        self.mainIndex = (self.mainIndex - 1) % self.mainSize
        self.__showInfo(self.mainIndex)

    def forward(self, widget):
        self.__update()
        self.mainIndex = (self.mainIndex + 1) % self.mainSize
        self.__showInfo(self.mainIndex)

    def previous(self, widget):
        self.__update()
        self.mainIndex = (self.mainIndex - 10) % self.mainSize
        self.__showInfo(self.mainIndex)

    def next(self, widget):
        self.__update()
        self.mainIndex = (self.mainIndex + 10) % self.mainSize
        self.__showInfo(self.mainIndex)

    def __showInfo(self, index):
        idArticolo = self.tabacchiList[index]["ID"]
        unitaMin = self.tabacchiList[index]["UnitaMin"]
        prezzoKg = self.tabacchiList[index]["PrezzoKg"]
        pezziUnitaMin = self.tabacchiList[index]["PezziUnitaMin"]
        prezzoConf = 0 if (pezziUnitaMin == 0) else (prezzoKg * unitaMin) / pezziUnitaMin
        self.counterLabel.set_text("%i di %i" % (index, self.mainSize))

        self.tipoLabel.set_text(self.tabacchiList[index]["Tipo"])
        self.codiceLabel.set_text(idArticolo)
        self.descLabel.set_text(self.tabacchiList[index]["Descrizione"])

        self.unitaminLabel.set_text(locale.format_string("%.3f kg", unitaMin))
        self.pezziUnitaLabel.set_text("%i" % pezziUnitaMin)
        self.prezzoLabel.set_text(locale.currency(prezzoKg, True, True))
        self.prezzoConfLabel.set_text(locale.currency(prezzoConf, True, True))
        self.livelloLabel.set_text(locale.format_string("%.3f kg", self.tabacchiList[index]["LivelloMin"]))

        # Se l'articolo e' stato ordinato
        if idArticolo in self.ordineDict:
            data = self.ordineDict[idArticolo]
            quantita = data[self.ID_QUANTITA]
            ordine = data[self.ID_ORDINE]
            consumo = float(data[self.ID_CONSUMO])
            costo = data[self.ID_COSTO]
        else:
            quantita = 0
            ordine = 0
            consumo = 0
            costo = 0

        # Blocca il propagarsi degli eventi
        self.quantitaSpinbutton.handler_block_by_func(self.quantitaChange)
        adjustmentQuantita = Gtk.Adjustment(value=0, lower=0, upper=(50 * unitaMin), step_incr=unitaMin)
        self.quantitaSpinbutton.set_adjustment(adjustmentQuantita)
        self.quantitaSpinbutton.set_value(quantita)
        # Ripristina il propagarsi degli eventi
        self.quantitaSpinbutton.handler_unblock_by_func(self.quantitaChange)

        self.ordineSpinbutton.handler_block_by_func(self.ordineChange)
        adjustmentOrdine = Gtk.Adjustment(value=0, lower=0, upper=(50 * unitaMin), step_incr=unitaMin)
        self.ordineSpinbutton.set_adjustment(adjustmentOrdine)
        self.ordineSpinbutton.set_value(ordine)
        self.ordineSpinbutton.handler_unblock_by_func(self.ordineChange)

        self.consumoLabel.set_text(locale.format_string("%.3f kg", consumo))
        self.totaleLabel.set_text(f"{locale.currency(costo, True, True)} su {locale.currency(self.mainTotal, True, True)}")
        self.quantitaSpinbutton.grab_focus()

    def run(self):
        self.result = self.ordineDialog.run()
        return self.result

    def close(self, widget, event=None):
        self.__update()

        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            # Mantiene la memoria dell'ultima posizione aggiornata nell'elenco tabacchi
            cursor.execute("update ordineTabacchi set LastPos = ? where ID = ?", (self.mainIndex, self.ordineID))
            conn.commit()
        except sqlite3.Error as e:
            if conn:
                conn.rollback()
            utility.gtkErrorMsg(e, self.ordineDialog)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
        self.ordineDialog.destroy()
