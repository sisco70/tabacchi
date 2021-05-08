#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) Francesco Guarnieri 2020 <francesco@guarnie.net>
#

import datetime
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import xlrd
import xlwt
import keyring
import base64
import locale
import gi

import browserWebkit2
from browserWebkit2 import Browser
import config
from config import log
import ordini
import preferencesTabacchi
from preferencesTabacchi import prefs
import stampe
import stats
import utility

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk, Gdk, Gio, GLib, GdkPixbuf, Pango  # noqa: E402

MENU_XML = """
<?xml version="1.0" encoding="UTF-8"?>
<interface>
    <menu id="app-menu">
        <submenu>
            <attribute name="label">Tabacchi</attribute>
            <section>
                <item>
                  <attribute name="label">Importa documenti Logista...</attribute>
                  <attribute name="action">app.importa</attribute>
                </item>
                <item>
                    <attribute name="label">Ricalcola consumi...</attribute>
                    <attribute name="action">app.ricalcola</attribute>
                </item>
                <item>
                    <attribute name="label">Statistiche...</attribute>
                    <attribute name="action">app.statistiche</attribute>
                </item>
            </section>
        </submenu>
        <submenu>
            <attribute name="label">Excel</attribute>
            <section>
                <item>
                    <attribute name="label">Genera inventario...</attribute>
                    <attribute name="action">app.excel</attribute>
                </item>
                <item>
                    <attribute name="label">Genera modello ordine...</attribute>
                    <attribute name="action">app.ordine_excel</attribute>
                </item>
            </section>
        </submenu>
        <submenu>
            <attribute name="label">Stampe</attribute>
            <section>
                <item>
                  <attribute name="label">Etichette prezzi...</attribute>
                  <attribute name="action">app.etichette</attribute>
                </item>
                <item>
                    <attribute name="label">Elenco articoli...</attribute>
                    <attribute name="action">app.articoli</attribute>
                </item>
            </section>
        </submenu>
        <section>
            <item>
                <attribute name="action">app.preferences</attribute>
                <attribute name="label" translatable="yes">_Preferences...</attribute>
            </item>
            <item>
                <attribute name="action">app.about</attribute>
                <attribute name="label" translatable="yes">_About</attribute>
            </item>
            <item>
                <attribute name="action">app.quit</attribute>
                <attribute name="label" translatable="yes">_Quit</attribute>
                <attribute name="accel">&lt;Primary&gt;q</attribute>
            </item>
        </section>
    </menu>
</interface>
"""


# Dialog per importare ordini o fatture dal sito Logista
class ImportDialog(utility.GladeWindow):
    ID_CODICE, ID_DESC, ID_PESO, ID_COSTO, ID_PREZZO_KG = (0, 1, 2, 3, 4)
    POS_CODICE, POS_DESC, POS_PESO, POS_COSTO = (1, 2, 3, 5)
    UNKNOWN, FATTURA, ORDINE = (-1, 1, 2)
    TIPO_DOC = {FATTURA: "Fattura", ORDINE: "Ordine"}

    modelInfoList = [
        ("Codice", "str", ID_CODICE),
        ("!+Descrizione", "str", ID_DESC),
        ("Peso", "float", ID_PESO),
        ("Costo", "currency", ID_COSTO),
        (None, "currency", ID_PREZZO_KG)]

    def __init__(self, parent):
        super().__init__(parent, "importDialog.glade")

        self.importDialog = self.builder.get_object("importDialog")
        self.importDialog.set_transient_for(parent)

        self.importTreeview = self.builder.get_object("importTreeview")
        self.dataLabel = self.builder.get_object("dataLabel")
        self.fileChooser = self.builder.get_object("fileChooserButton")
        self.tipoDocLabel = self.builder.get_object("tipoDocLabel")
        self.totaleKgLabel = self.builder.get_object("totaleKgLabel")
        self.totaleEuroLabel = self.builder.get_object("totaleEuroLabel")

        listino_formats = {self.ID_PESO: "%.3f kg"}

        listino_properties = {self.ID_CODICE: {"xalign": 1, "scale": utility.PANGO_SCALE_SMALL},
                              self.ID_COSTO: {"xalign": 1, "scale": utility.PANGO_SCALE_SMALL}}

        utility.ExtTreeView(self.modelInfoList, self.importTreeview, formats=listino_formats, properties=listino_properties)
        self.importModel = self.importTreeview.get_model()

        fileFilter = Gtk.FileFilter()
        fileFilter.set_name("Pdf files")
        fileFilter.add_pattern("*.pdf")
        self.fileChooser.add_filter(fileFilter)

        self.builder.connect_signals({"on_importDialog_delete_event": self.close,
                                      "on_okButton_clicked": self.okClose,
                                      "on_cancelButton_clicked": self.close,
                                      "on_fileChooserButton_file_set": self.importDoc
                                      })

        self.dataOrdine = None
        self.dataFattura = None
        self.totalKg = 0
        self.totalEuro = 0
        self.result = None
        self.tipo = self.UNKNOWN
        self.tabacchiDict = self.__getTabacchiDict()
        self.dataOrdineEntry = utility.DataEntry(self.builder.get_object("dataOrdineButton"), None, " %d %B %Y ")

    # Genera un dizionario con l'attuale listino tabacchi
    def __getTabacchiDict(self):
        conn = None
        cursor = None
        tabacchiDict = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute("SELECT ID, Descrizione, PrezzoKg FROM tabacchi")
            resultset = cursor.fetchall()
            tabacchiDict = dict()
            for row in resultset:
                tabacchiDict[row["ID"]] = [row["Descrizione"], row["PrezzoKg"]]
        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self.importDialog)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        return tabacchiDict

    # Importa un documento Logista
    def importDoc(self, widget):
        (self.dataFattura, self.tipo, self.totalEuro, self.totalKg) = self.__importDoc(self.importModel, self.tabacchiDict)

        self.totaleEuroLabel.set_text(locale.currency(self.totalEuro, True, True))
        self.totaleKgLabel.set_text(locale.format_string("%.3f", self.totalKg))

        if self.dataFattura:
            self.dataLabel.set_text(self.dataFattura.strftime("%A %d %B %Y"))
            self.dataOrdine = self.dataFattura - datetime.timedelta(days=prefs.ggPerOrdine)
            self.dataOrdineEntry.setDate(self.dataOrdine)
        else:
            self.dataLabel.set_text("")

        if self.tipo == self.UNKNOWN:
            self.tipoDocLabel.set_text("")
        else:
            self.tipoDocLabel.set_text(self.TIPO_DOC[self.tipo])

    # Importa un documento Logista (fattura o ordine) nel modello passato per parametro, usando un dizionario per riconoscere i codici articolo

    def __importDoc(self, model, tabacchiDict):
        data = None
        tipo = self.UNKNOWN
        totalKg = 0
        totalEuro = 0
        filename = self.fileChooser.get_filename()
        if not filename:
            msgDialog = Gtk.MessageDialog(parent=self.importDialog, modal=True, message_type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.CLOSE, text="Importazione non avvenuta.")
            msgDialog.format_secondary_text("Seleziona un documento pdf.")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()
        else:
            model.clear()
            row = [None] * model.get_n_columns()

            tmpFile = tempfile.NamedTemporaryFile(delete=False)
            tmpFile.close()

            command = "pdftotext -layout %s %s" % (filename, tmpFile.name)
            try:
                subprocess.check_call(command, shell=True)
            except Exception as e:
                utility.gtkErrorMsg(e, self.importDialog)
                return

            # Check per stabilire se il documento è una fattura o un ordine Logista
            fatturaPattern = re.compile(r".*- FATTURA U13 -.*")
            ordinePattern = re.compile(r".*Numero ordine\s+\d+.*")
            with open(tmpFile.name, 'r') as f:
                for line in f:
                    if fatturaPattern.match(line):
                        tipo = self.FATTURA
                        break
                    elif ordinePattern.match(line):
                        tipo = self.ORDINE
                        break

            if tipo != self.UNKNOWN:
                if tipo == self.FATTURA:
                    headerPattern = re.compile(r"\s*CODICE\s+DESCRIZIONE\s+\S+\s+PREZZO\s+IMPORTO LORDO\s*")
                    datePattern = re.compile(r".*\s+(\d\d\.\d\d\.\d\d\d\d)\s*")
                    rowPattern = re.compile(r"^\s*1000*(\d+)\d{3}\s+(.*)\s+(\d+,\d+)\s+(\d+,\d+)\s+([0-9\.\,]+)\s*$")
                    ignorePattern = re.compile(r"\s*===\s+.*")
                if tipo == self.ORDINE:
                    headerPattern = re.compile(r"\s*Riga\s+Cod\.AAMS\s+Descrizione\s+Quantità\s*")
                    datePattern = re.compile(r".*\s+(\d\d\.\d\d\.\d\d\d\d)\s*")
                    rowPattern = re.compile(r"\s*\d+\s+(\d+)\s+(.*)\s+(\d+,\d+)\s*")
                    ignorePattern = re.compile(r"\s*===\s+.*")

                body = False
                with open(tmpFile.name, 'r') as f:
                    for line in f:
                        # Se non siamo nel corpo
                        if not body:
                            m = headerPattern.match(line)
                            # Se trovo la testata, allora inizia il corpo
                            if m:
                                body = True
                            elif not data:
                                m = datePattern.match(line)
                                if m:
                                    data = datetime.datetime.strptime(m.group(1), "%d.%m.%Y")

                        # Se siamo nel corpo
                        else:
                            m = rowPattern.match(line)
                            # E' una riga standard
                            if m:
                                idCod = m.group(1).strip()
                                peso = locale.atof(m.group(3))
                                costo = 0
                                prezzo_kg = 0
                                row[self.ID_CODICE] = idCod
                                row[self.ID_PESO] = peso
                                row[self.ID_DESC] = m.group(2).strip()

                                if tipo == self.FATTURA:
                                    costo = locale.atof(m.group(5))
                                    prezzo_kg = costo / peso
                                if idCod in tabacchiDict:
                                    # row[self.ID_DESC] = tabacchiDict[idCod][0]
                                    if tipo == self.ORDINE:
                                        prezzo_kg = tabacchiDict[idCod][1]
                                        costo = prezzo_kg * peso

                                row[self.ID_COSTO] = costo
                                row[self.ID_PREZZO_KG] = round(prezzo_kg, 3)

                                totalEuro += costo
                                totalKg += peso
                                model.append(row)
                            # Se non è una riga standard e non è una riga da ignorare, allora è finito il corpo
                            elif not ignorePattern.match(line):
                                body = False

            if not data or not model.get_iter_first():
                data = None
                model.clear()
                msgDialog = Gtk.MessageDialog(parent=self.importDialog, modal=True, message_type=Gtk.MessageType.WARNING,
                                              buttons=Gtk.ButtonsType.CLOSE, text="Importazione non avvenuta.")
                msgDialog.format_secondary_text("Il documento pdf è un documento Logista?")
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()
        return data, tipo, totalEuro, totalKg

    def run(self):
        self.importDialog.run()
        return self.result

    def okClose(self, widget):
        if not self.importModel.get_iter_first():
            msgDialog = Gtk.MessageDialog(parent=self.importDialog, modal=True, message_type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.OK, text="Non è stato importato alcun ordine.")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()
        else:
            self.result = [self.dataFattura, self.importModel, self.tipo, self.dataOrdine]
            self.importDialog.destroy()

    def close(self, widget, other=None):
        self.importDialog.destroy()

# Dialog per scegliere il tipo di Ordine


class U88Dialog(utility.GladeWindow):
    def __init__(self, parent, data):
        super().__init__(parent, "u88Dialog.glade")

        self.u88Dialog = self.builder.get_object("u88Dialog")
        self.u88Dialog.set_transient_for(parent)

        self.straordinario = self.builder.get_object("straordinario_radiobutton")
        self.urgente = self.builder.get_object("urgente_radiobutton")

        self.dataEntry = utility.DataEntry(self.builder.get_object("dataButton"), data, " %d %B %Y ")
        self.result = None

        self.builder.connect_signals({"on_u88Dialog_delete_event": self.close,
                                      "on_okButton_clicked": self.okClose,
                                      "on_cancelButton_clicked": self.close
                                      })

    def run(self):
        self.u88Dialog.run()
        return self.result

    def okClose(self, widget):
        tipo = ordini.ORDINARIO
        if self.straordinario.get_active():
            tipo = ordini.STRAORDINARIO
        elif self.urgente.get_active():
            tipo = ordini.URGENTE
        self.result = (tipo, self.dataEntry.data)
        self.u88Dialog.destroy()

    def close(self, widget, other=None):
        self.u88Dialog.destroy()


class TabacchiDialog(utility.GladeWindow):
    MAGAZZINO_TAB, LISTINO_TAB = (0, 1)
    IN_MAGAZZINO, ID, DESCRIZIONE, TIPO, PREZZO_PEZZO, DECORRENZA, LIVELLO_MIN, PEZZI_UNITA_MIN, UNITA_MIN, PREZZO_KG, BARCODE, DIRTY = (
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)

    MODEL_INFO_LIST = [
        ("*Magazzino", "bool", IN_MAGAZZINO),
        (None, "str", ID),
        ("^!+Descrizione", "str", DESCRIZIONE),
        ("^Tipo", "str", TIPO),
        (None, "currency", PREZZO_PEZZO),
        (None, "float", LIVELLO_MIN),
        ("*Pezzi", "int#3,0", PEZZI_UNITA_MIN),
        ("Unità min.", "float", UNITA_MIN),
        ("Prezzo Kg", "currency", PREZZO_KG),
        (None, "str", BARCODE),
        ("^Decorrenza", "date", DECORRENZA),
        (None, "bool", DIRTY)]
    MODEL_INFO_LIST_MAGAZZINO = [
        (None, "bool", IN_MAGAZZINO),
        ("Codice", "str", ID),
        ("!+Descrizione", "str", DESCRIZIONE),
        ("Tipo", "str", TIPO),
        ("Prezzo", "currency", PREZZO_PEZZO),
        ("*Livello", "float#3,3/i8,0,12", LIVELLO_MIN),
        (None, "int", PEZZI_UNITA_MIN),
        (None, "float", UNITA_MIN),
        (None, "currency", PREZZO_KG),
        ("*Barcode", "str", BARCODE),
        (None, "date", DECORRENZA),
        (None, "bool", DIRTY)]

    def __init__(self, parent):
        super().__init__(parent, "tabacchiDialog.glade")
        self.dirtyFlag = False
        self.data = prefs.dataCatalogo
        self.readBarcodeThread = None
        self.barcodeDict = dict()
        self.deleteList = []

        self.tabacchiDialog = self.builder.get_object("tabacchiDialog")
        self.tabacchiNotebook = self.builder.get_object("tabacchiNotebook")
        self.bluetoothStatusImage = self.builder.get_object('bluetoothStatusImage')
        self.listinoTreeView = self.builder.get_object("listinoTreeView")
        self.magazzinoTreeView = self.builder.get_object("magazzinoTreeView")

        self.tabacchiDialog.set_transient_for(parent)

        listino_callbacks = {self.IN_MAGAZZINO: self.__toggledCallback, self.PEZZI_UNITA_MIN: self.onListinoCellEdited}
        magazzino_callbacks = {self.BARCODE: self.onBarcodeCellEdited, self.LIVELLO_MIN: self.onCellEdited}
        magazzino_properties = {self.ID: {"xalign": 1, "scale": utility.PANGO_SCALE_SMALL}, self.TIPO: {"xalign": 0.5, "scale": utility.PANGO_SCALE_SMALL}, self.BARCODE: {
            "xalign": 1, "style": Pango.Style.ITALIC, "scale": utility.PANGO_SCALE_SMALL}, self.DECORRENZA: {"xalign": 0.5, "scale": utility.PANGO_SCALE_SMALL}}
        listino_properties = {self.DESCRIZIONE: {"scale": utility.PANGO_SCALE_SMALL},
                              self.TIPO: {"xalign": 0.5, "scale": utility.PANGO_SCALE_SMALL},
                              self.DECORRENZA: {"xalign": 0.5, "scale": utility.PANGO_SCALE_SMALL}}

        utility.ExtTreeView(self.MODEL_INFO_LIST, self.listinoTreeView, edit_callbacks=listino_callbacks, properties=listino_properties)
        self.listinoModel = self.listinoTreeView.get_model()
        self.magazzinoModel = self.listinoModel.filter_new()
        self.magazzinoModel.set_visible_column(0)

        utility.ExtTreeView(self.MODEL_INFO_LIST_MAGAZZINO, self.magazzinoTreeView, modelCallback=self.modelCallback,
                            edit_callbacks=magazzino_callbacks, properties=magazzino_properties)

        # Prima di caricare i dati nel modello, lo "sgancio" dalla Treeview
        self.listinoTreeView.set_model(None)
        self.magazzinoTreeView.set_model(None)

        # Legge da db i dati nel modello
        self.loadListino(self.listinoModel)

        # Finita la lettura da DB lo "riaggancio" alla Treeview
        self.magazzinoTreeView.set_model(self.magazzinoModel)
        self.listinoTreeView.set_model(self.listinoModel)

        self.updateTitle()

        self.listinoTreeView.connect("row-activated", self.changeView)
        self.magazzinoTreeView.connect("row-activated", self.changeView)

        self.builder.connect_signals({"on_tabacchiDialog_delete_event": self.forcedClose,
                                      "on_updateButton_clicked": self.updateCatalogo,
                                      "on_okButton_clicked": self.close,
                                      "on_barcodeToolbutton_clicked": self.enableBarcode,
                                      "on_cancelButton_clicked": self.forcedClose})

        self.bluetoothStatusImage.hide()

    def modelCallback(self, model):
        return self.magazzinoModel

    def onBarcodeCellEdited(self, widget, path, value, model, col_id):
        self.__changeBarcode(value, path, model)

    def onCellEdited(self, widget, path, value, model, col_id):
        parent_path = model.convert_path_to_child_path(Gtk.TreePath.new_from_string(path))
        parent_model = model.get_model()
        parent_model[parent_path][col_id] = value
        parent_model[parent_path][self.DIRTY] = True
        self.dirtyFlag = True

    def onListinoCellEdited(self, widget, path, value, model, col_id):
        model[path][col_id] = value
        model[path][self.DIRTY] = True
        self.dirtyFlag = True

    def __changeBarcode(self, barcode, path, model):
        barcode.strip()
        if (barcode in self.barcodeDict):
            msgDialog = Gtk.MessageDialog(parent=self.tabacchiDialog, modal=True, message_type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.OK, text="Codice a barre già associato:")
            msgDialog.format_secondary_text("%s" % self.barcodeDict[barcode])
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()
        else:
            parent_path = model.convert_path_to_child_path(Gtk.TreePath.new_from_string(path))
            parent_model = model.get_model()
            old_barcode = parent_model[parent_path][self.BARCODE]
            if len(old_barcode) > 0:
                del self.barcodeDict[old_barcode]
            if len(barcode) > 0:
                self.barcodeDict[barcode] = parent_model[parent_path][self.DESCRIZIONE]
            parent_model[parent_path][self.BARCODE] = barcode
            parent_model[parent_path][self.DIRTY] = True
            self.dirtyFlag = True

    def __toggledCallback(self, widget, path, model, col_id):
        new_value = model[path][col_id] = not model[path][col_id]
        model[path][self.DIRTY] = True
        self.dirtyFlag = True
        if new_value:
            child_path = self.magazzinoModel.convert_child_path_to_path(Gtk.TreePath.new_from_string(path))
            self.tabacchiNotebook.set_current_page(self.MAGAZZINO_TAB)
            self.magazzinoTreeView.set_cursor(child_path)

    def enableBarcode(self, widget):
        if self.readBarcodeThread:
            self.readBarcodeThread.stop()
            self.readBarcodeThread = None
            self.bluetoothStatusImage.hide()
        else:
            defBarcode = prefs.barcodeList[prefs.defaultBarcode]
            connectBarcodeThread = preferencesTabacchi.ConnectBarcodeThread(defBarcode[1], defBarcode[2])
            progressDialog = utility.ProgressDialog(self.tabacchiDialog, "Connecting to device %s.." %
                                                    defBarcode[0], "", "Bluetooth Barcode reader", connectBarcodeThread)
            progressDialog.setResponseCallback(self.responseCallback)
            progressDialog.setErrorCallback(self.bluetoothStatusImage.hide)
            progressDialog.startPulse()

    def responseCallback(self, sock):
        if sock:
            self.bluetoothStatusImage.show()
            self.readBarcodeThread = preferencesTabacchi.ReadBarcodeThread(self.updateCallback, self.errorCallback, sock)
            self.readBarcodeThread.start()
        else:
            self.bluetoothStatusImage.hide()

    def errorCallback(self, e):
        self.bluetoothStatusImage.hide()
        utility.gtkErrorMsg(e, self.tabacchiDialog)

    # Metodo che viene invocato dal thread di comunicazione bluetooth ogni volta che si legge un codice
    def updateCallback(self, data):
        selection = self.magazzinoTreeView.get_selection()
        if selection:
            model, iterator = selection.get_selected()
            if iterator:
                self.__changeBarcode(data, model.get_string_from_iter(iterator), model)
            else:
                msgDialog = Gtk.MessageDialog(parent=self.tabacchiDialog, modal=True, message_type=Gtk.MessageType.WARNING,
                                              buttons=Gtk.ButtonsType.OK, text="E' necessario selezionare un articolo.")
                msgDialog.format_secondary_text("Codice a barre non memorizzato.")
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()

    def changeView(self, widget, path, arg1=None):
        if widget == self.magazzinoTreeView:
            parent_path = self.magazzinoModel.convert_path_to_child_path(path)
            self.tabacchiNotebook.set_current_page(self.LISTINO_TAB)
            self.listinoTreeView.set_cursor(parent_path)
        else:
            child_path = self.magazzinoModel.convert_child_path_to_path(path)
            if child_path:
                self.tabacchiNotebook.set_current_page(self.MAGAZZINO_TAB)
                self.magazzinoTreeView.set_cursor(child_path)

    def loadListino(self, model):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute(
                "SELECT ID, Descrizione, UnitaMin, PrezzoKg, Tipo, InMagazzino, LivelloMin, Decorrenza as 'Decorrenza [date]', PezziUnitaMin, Barcode FROM tabacchi order by Tipo desc, Descrizione")
            result_set = cursor.fetchall()

            model.clear()
            self.barcodeDict.clear()

            for row in result_set:
                tipo = row["Tipo"].strip()
                iterator = model.append()
                model.set_value(iterator, self.ID, row["ID"])
                desc = row["Descrizione"].strip()

                model.set_value(iterator, self.DESCRIZIONE, desc)
                model.set_value(iterator, self.TIPO, tipo)

                model.set_value(iterator, self.LIVELLO_MIN, row["LivelloMin"])
                model.set_value(iterator, self.IN_MAGAZZINO, row["InMagazzino"])
                barcode = row["Barcode"]
                if barcode and len(barcode) > 0:
                    self.barcodeDict[barcode] = desc
                    model.set_value(iterator, self.BARCODE, barcode)
                else:
                    model.set_value(iterator, self.BARCODE, '')
                decorrenza = row["Decorrenza"]

                if not decorrenza:
                    decorrenza = datetime.datetime(1970, 1, 1, 0, 0)

                model.set_value(iterator, self.DECORRENZA, decorrenza)
                prezzoKg = row["PrezzoKg"]
                unitaMin = row["UnitaMin"]
                pezziUnitaMin = row["PezziUnitaMin"]
                if pezziUnitaMin > 0:
                    prezzoPezzo = (prezzoKg * unitaMin) / pezziUnitaMin
                else:
                    prezzoPezzo = 0
                model.set_value(iterator, self.PREZZO_KG, prezzoKg)
                model.set_value(iterator, self.UNITA_MIN, unitaMin)
                model.set_value(iterator, self.PEZZI_UNITA_MIN, pezziUnitaMin)
                model.set_value(iterator, self.PREZZO_PEZZO, prezzoPezzo)
                model.set_value(iterator, self.DIRTY, False)
        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self.tabacchiDialog)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def __saveModelToDB(self):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            for _id in self.deleteList:
                cursor.execute("delete from tabacchi where ID = ?", (_id,))
            for row in self.listinoModel:
                if row[self.DIRTY]:
                    inMagazzino = row[self.IN_MAGAZZINO]
                    codiceAAMS = row[self.ID]
                    descrizione = row[self.DESCRIZIONE]
                    tipo = row[self.TIPO]
                    decorrenza = row[self.DECORRENZA]
                    livelloMin = row[self.LIVELLO_MIN]
                    pezziUnitaMin = row[self.PEZZI_UNITA_MIN]
                    unitaMin = row[self.UNITA_MIN]
                    prezzoKg = row[self.PREZZO_KG]
                    barcode = row[self.BARCODE]

                    cursor.execute(
                        "UPDATE tabacchi SET Descrizione=?, UnitaMin=?, PrezzoKg=?, Tipo=?, InMagazzino=?, LivelloMin=?, Decorrenza=?, PezziUnitaMin=?, Barcode=?  WHERE ID = ?",
                        (descrizione, unitaMin, prezzoKg, tipo, inMagazzino, livelloMin, decorrenza, pezziUnitaMin, barcode, codiceAAMS))
                    if cursor.rowcount == 0:
                        cursor.execute(
                            "INSERT INTO tabacchi(Descrizione, UnitaMin, PrezzoKg, Tipo, InMagazzino, LivelloMin, Decorrenza, PezziUnitaMin, Barcode, ID) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            (descrizione, unitaMin, prezzoKg, tipo, inMagazzino, livelloMin, decorrenza, pezziUnitaMin, barcode, codiceAAMS))

            conn.commit()
        except sqlite3.Error as e:
            if conn:
                conn.rollback()
            utility.gtkErrorMsg(e, self.tabacchiDialog)
        else:
            del self.deleteList[:]
            if self.dirtyFlag:
                prefs.setDBDirty()
            # Aggiornamento data catalogo
            prefs.dataCatalogo = self.data
            prefs.save()
            self.dirtyFlag = False
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def updateTitle(self):
        self.tabacchiDialog.set_title("Catalogo Logista - %s" % self.data.strftime("%d %B %Y"))

    def run(self):
        self.result = self.tabacchiDialog.run()
        return self.result

    def forcedClose(self, widget, event=None):
        self.close(widget, True)

    def close(self, widget, forced=False):
        if self.dirtyFlag:
            if forced:
                msgDialog = Gtk.MessageDialog(parent=self.tabacchiDialog, modal=True, message_type=Gtk.MessageType.QUESTION,
                                              buttons=Gtk.ButtonsType.YES_NO, text="Ci sono modifiche non salvate.")
                msgDialog.format_secondary_text("Vuoi salvarle?")
                response = msgDialog.run()
                msgDialog.destroy()
                if response == Gtk.ResponseType.YES:
                    self.__saveModelToDB()
            else:
                self.__saveModelToDB()
        if self.readBarcodeThread:
            self.readBarcodeThread.stop()
        self.tabacchiDialog.destroy()

    # Logica per interpretare il n. di pezzi in una confezione (pezziUnitaMin)
    def __fuzzyCount(self, unitaMin, descrizione):
        value = 0
        if ("*CART20" in descrizione) or ("*AST20" in descrizione):
            value = 10
        elif ("*AST10" in descrizione):
            value = 20
        else:
            match = re.search(r"\*(\d{1,3})GR", descrizione)
            if match:
                value = int(unitaMin / (float(match.group(1)) / 1000))
        return value

    # Aggiornamento DB Tabacchi tramite portale Logista (scaricando listino su file Excel)
    def updateCatalogo(self, widget=None):
        url = prefs.catalogoUrl

        # Download dal portale Logista del file con il catalogo aggionato
        if url is not None and (len(url) > 0):
            username = prefs.tabacchiUser
            password = keyring.get_password(prefs.TABACCHI_STR, username)
            filename = f"{tempfile.mkdtemp()}/catalogo.xls"
            print(f"{username=} {password=} {filename=}")
            downloadThread = utility.DownloadThread(url, filename, username, password)
            progressDialog = utility.ProgressDialog(self.tabacchiDialog, "Download catalogo Tabacchi in corso..", "Dal sito www.logista.it", "Aggiornamento catalogo Logista", downloadThread)
            progressDialog.setResponseCallback(self.__updateCatalogoCallback)
            progressDialog.start()
        else:
            msgDialog = Gtk.MessageDialog(parent=self.tabacchiDialog, modal=True, message_type=Gtk.MessageType.ERROR, buttons=Gtk.ButtonsType.CANCEL, text="URL portale Logista mancante.")
            msgDialog.format_secondary_text("Aggiornare le preferenze.")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()

    #
    def __updateCatalogoCallback(self, filename):
        # Apre il file Excel
        book = xlrd.open_workbook(filename)
        sheet = book.sheet_by_index(0)

        tabacchiDict = dict()
        i = 0
        for row in self.listinoModel:
            codiceAAMS = row[self.ID]
            tabacchiDict[codiceAAMS] = i
            i += 1

        for row in range(1, sheet.nrows):
            codiceAAMS = sheet.cell_value(row, 0).strip()
            descrizione = sheet.cell_value(row, 2).strip()

            prezzoKg = float(sheet.cell_value(row, 5))
            unitaMin = float(sheet.cell_value(row, 4))
            tipo = sheet.cell_value(row, 3).strip()
            data = sheet.cell_value(row, 7).strip()
            try:
                data = datetime.datetime.strptime(data, "%d/%m/%Y")
                decorrenza = datetime.date(data.year, data.month, data.day)
            except ValueError:
                decorrenza = datetime.date(1970, 1, 1)

            pezziUnitaMin = self.__fuzzyCount(unitaMin, descrizione)

            # Se è un nuovo articolo
            if codiceAAMS not in tabacchiDict:
                iterator = self.listinoModel.append()
                self.listinoModel.set_value(iterator, self.IN_MAGAZZINO, False)
                self.listinoModel.set_value(iterator, self.ID, codiceAAMS)
                self.listinoModel.set_value(iterator, self.DESCRIZIONE, descrizione)
                self.listinoModel.set_value(iterator, self.LIVELLO_MIN, float(0))
                self.listinoModel.set_value(iterator, self.TIPO, tipo)
                self.listinoModel.set_value(iterator, self.UNITA_MIN, unitaMin)
                self.listinoModel.set_value(iterator, self.PREZZO_KG, prezzoKg)
                self.listinoModel.set_value(iterator, self.DECORRENZA, decorrenza)
                self.listinoModel.set_value(iterator, self.PEZZI_UNITA_MIN, pezziUnitaMin)
                self.listinoModel.set_value(iterator, self.DIRTY, True)
                self.listinoModel.set_value(iterator, self.BARCODE, '')
                self.dirtyFlag = True
            else:  # Se già esiste..
                path = tabacchiDict[codiceAAMS]
                row = self.listinoModel[path]

                if (row[self.PREZZO_KG] != prezzoKg) or (row[self.DESCRIZIONE] != descrizione) or (row[self.UNITA_MIN] != unitaMin) or (row[self.TIPO] != tipo) or (row[self.DECORRENZA] != decorrenza) or ((pezziUnitaMin > 0) and (row[self.PEZZI_UNITA_MIN] != pezziUnitaMin)):
                    row[self.PREZZO_KG] = prezzoKg
                    row[self.DECORRENZA] = decorrenza
                    row[self.DESCRIZIONE] = descrizione
                    row[self.UNITA_MIN] = unitaMin
                    row[self.TIPO] = tipo
                    row[self.DIRTY] = True
                    if (pezziUnitaMin > 0) and (row[self.PEZZI_UNITA_MIN] != pezziUnitaMin):
                        row[self.PEZZI_UNITA_MIN] = pezziUnitaMin
                    else:
                        pezziUnitaMin = row[self.PEZZI_UNITA_MIN]
                    row[self.PREZZO_PEZZO] = (prezzoKg * unitaMin) / pezziUnitaMin if pezziUnitaMin > 0 else 0
                    self.dirtyFlag = True

                # Elimina gli articoli che sono sia in memoria che nel file excel
                del tabacchiDict[codiceAAMS]

        self.data = datetime.datetime.now().replace(microsecond=0)

        # Se sono rimasti articoli, questi sono da cancellare..
        if len(tabacchiDict) > 0:
            del self.deleteList[:]
            deleteDesc = []
            for key_value in tabacchiDict.keys():
                row = self.listinoModel[tabacchiDict[key_value]]
                deleteDesc.append([row[self.ID], row[self.DESCRIZIONE], row[self.IN_MAGAZZINO]])
                self.deleteList.append(key_value)
            modelInfo = [("Codice", "str"), ("+Descrizione", "str"), ("^Magazzino", "bool")]
            extMsgDialog = utility.ExtMsgDialog(
                self.tabacchiDialog, modelInfo, "Nel listino Logista i seguenti articoli non esistono più.", "Attenzione", "dialog-warning-symbolic",
                buttons=utility.ExtMsgDialog.YES_NO)
            extMsgDialog.setSecondaryLabel("Vuoi cancellarli?")
            extMsgDialog.setData(deleteDesc)
            response = extMsgDialog.run()

            if response == Gtk.ResponseType.YES:
                iterator = self.listinoModel.get_iter_first()
                result = True
                self.dirtyFlag = True
                while iterator and result:
                    codiceAAMS = self.listinoModel.get_value(iterator, self.ID)
                    if codiceAAMS in tabacchiDict:
                        result = self.listinoModel.remove(iterator)
                    else:
                        iterator = self.listinoModel.iter_next(iterator)
        self.updateTitle()

# Thread dedicato a ricalcolare i consumi


class RicalcolaConsumiThread(utility.WorkerThread):
    def __init__(self):
        super().__init__()

    def run(self):
        conn = None
        cursor = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute("SELECT r.ID, r.Ordine, r.Giacenza, r.ID_Ordine FROM rigaOrdineTabacchi r, ordineTabacchi o where r.ID_Ordine = o.ID order by r.ID, r.ID_Ordine")
            resultList = cursor.fetchall()
            old_id = -1
            self.progressDialog.setSteps(len(resultList))

            for row in resultList:
                peso = round(row["Ordine"], 3)
                quantita = round(row["Giacenza"], 3)
                idCod = row["ID"]
                if (idCod != old_id):
                    old_quantita = 0
                    old_peso = 0
                consumo = round((old_quantita + old_peso) - quantita, 3)
                cursor.execute("update rigaOrdineTabacchi set Consumo = ? where ID = ? and ID_Ordine = ?", (consumo, idCod, row["ID_Ordine"]))
                old_peso = peso
                old_quantita = quantita
                old_id = idCod
                self.update()
            conn.commit()
        except StopIteration:
            if conn:
                conn.rollback()
        except sqlite3.Error as e:
            if conn:
                conn.rollback()
            self.setError(e)
        else:
            self.status = self.DONE
            prefs.setDBDirty()
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
            GLib.idle_add(self.progressDialog.close)

        return False

# Finestra principale


class MainWindow(Gtk.ApplicationWindow, utility.GladeWindow):
    ID, DATA, PESO, IMPORTO, STATO, DATA_SUPPLETIVO, CONSEGNA, SUPPLETIVO = (0, 1, 2, 3, 4, 5, 6, 7)
    modelInfoList = [
        (None, "int", ID),
        ("^Data ordine", "date", DATA),
        ("^Peso", "float", PESO),
        ("^Importo", "currency", IMPORTO),
        ("+Stato", ordini.STATO_ORDINE, STATO),
        ("Consegna", "date", CONSEGNA),
        ("Suppletivo", ordini.TIPO_ORDINE, SUPPLETIVO),
        (None, "date", DATA_SUPPLETIVO)]

    def __init__(self, *args, **kwargs):
        Gtk.ApplicationWindow.__init__(self, *args, **kwargs)
        utility.GladeWindow.__init__(self, self, "mainWindowContent.glade")

        self.add(self.builder.get_object("tabacchiVBox"))
        self.logo = GdkPixbuf.Pixbuf.new_from_file(str(config.RESOURCE_PATH / 'cigarette.png'))
        self.set_icon(self.logo)

        #           CREA LA HEADER BAR              #

        header_bar = Gtk.HeaderBar()
        header_bar.set_show_close_button(True)
        header_bar.set_title(self.get_title())

        btn = Gtk.MenuButton()
        icon = Gio.ThemedIcon(name="open-menu-symbolic")
        image = Gtk.Image.new_from_gicon(icon, Gtk.IconSize.BUTTON)
        btn.add(image)
        menuBuilder = Gtk.Builder.new_from_string(MENU_XML, -1)
        btn.set_popover(Gtk.Popover.new_from_model(btn, menuBuilder.get_object("app-menu")))

        header_bar.pack_end(btn)
        self.set_titlebar(header_bar)

        #                                          #

        if not prefs.checkDB():
            msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.OK_CANCEL, text="Inizializzo un nuovo database?")
            msgDialog.format_secondary_text(f"Non trovo {prefs.DB_PATHNAME}")
            msgDialog.set_title("Attenzione")
            response = msgDialog.run()
            msgDialog.destroy()
            if (response == Gtk.ResponseType.OK):
                conn = None
                try:
                    conn = sqlite3.connect(prefs.DB_PATHNAME)
                    conn.execute("pragma foreign_keys=off;")

                    # Table: ordineTabacchi
                    conn.execute("CREATE TABLE ordineTabacchi (ID INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, Data DATETIME NOT NULL, LastPos INTEGER NOT NULL DEFAULT (0), DataSuppletivo DATE DEFAULT NULL, Stato INTEGER NOT NULL DEFAULT (0), Levata DATE DEFAULT NULL, Suppletivo BOOLEAN NOT NULL DEFAULT (0), UNIQUE (Data), UNIQUE (Levata));")

                    # Table: rigaOrdineSuppletivo
                    conn.execute("CREATE TABLE rigaOrdineSuppletivo (ID varchar (8) NOT NULL, Descrizione varchar (50) NOT NULL, ID_Ordine integer NOT NULL, Ordine float NOT NULL DEFAULT '0', PRIMARY KEY (ID, ID_Ordine), CONSTRAINT fk_riga_suppletivo FOREIGN KEY (ID_Ordine) REFERENCES ordineTabacchi (ID));")

                    # Table: rigaOrdineTabacchi
                    conn.execute("CREATE TABLE rigaOrdineTabacchi (ID TEXT (8) NOT NULL, Descrizione TEXT (50) NOT NULL, ID_Ordine INTEGER NOT NULL, Ordine REAL NOT NULL DEFAULT (0), Prezzo REAL NOT NULL DEFAULT (0), Giacenza REAL NOT NULL DEFAULT (0), Consumo REAL NOT NULL DEFAULT (0), PRIMARY KEY (ID, ID_Ordine), CONSTRAINT fk_riga_ordine FOREIGN KEY (ID_Ordine) REFERENCES ordineTabacchi (ID));")

                    # Table: tabacchi
                    conn.execute("CREATE TABLE tabacchi (ID TEXT (8) NOT NULL, Descrizione TEXT (50) NOT NULL DEFAULT NULL, UnitaMin REAL NOT NULL DEFAULT (0), PrezzoKg REAL NOT NULL DEFAULT (0), Tipo TEXT (50) NOT NULL, InMagazzino BOOLEAN NOT NULL DEFAULT (0), LivelloMin REAL NOT NULL DEFAULT (0), Decorrenza DATE NOT NULL DEFAULT ('0000-00-00'), PezziUnitaMin INTEGER NOT NULL DEFAULT (0), Barcode TEXT (20) NOT NULL, PRIMARY KEY (ID));")

                    # Table: verificaOrdine
                    conn.execute("CREATE TABLE verificaOrdine (ID VARCHAR (8) NOT NULL, ID_ordine INTEGER NOT NULL, Carico REAL NOT NULL DEFAULT (0), Peso REAL NOT NULL DEFAULT (0), Eliminato BOOLEAN NOT NULL DEFAULT (0), PRIMARY KEY (ID, ID_ordine), CONSTRAINT fk_verificaOrdine_OrdineTabacchi FOREIGN KEY (ID_ordine) REFERENCES ordineTabacchi (ID) ON DELETE NO ACTION ON UPDATE NO ACTION);")

                    # Index: idx_rigaOrdineSuppletivo_fk_riga_suppletivo
                    conn.execute("CREATE INDEX idx_rigaOrdineSuppletivo_fk_riga_suppletivo ON rigaOrdineSuppletivo (ID_Ordine);")

                    # Index: idx_rigaOrdineTabacchi_fk_riga_ordine
                    conn.execute("CREATE INDEX idx_rigaOrdineTabacchi_fk_riga_ordine ON rigaOrdineTabacchi (ID_Ordine);")

                    # Index: idx_tabacchi_k_inmagazzino
                    conn.execute("CREATE INDEX idx_tabacchi_k_inmagazzino ON tabacchi (InMagazzino);")

                    # Index: idx_tabacchi_k_tipo
                    conn.execute("CREATE INDEX idx_tabacchi_k_tipo ON tabacchi (Tipo);")

                    # Index: idx_verificaOrdine_fk_verificaOrdine_OrdineTabacchi
                    conn.execute("CREATE INDEX idx_verificaOrdine_fk_verificaOrdine_OrdineTabacchi ON verificaOrdine (ID_ordine);")

                    conn.commit()
                except sqlite3.Error as e:
                    if conn:
                        conn.rollback()
                    utility.gtkErrorMsg(e, self)
                    sys.exit(1)
                finally:
                    if conn:
                        conn.close()
            else:
                sys.exit(1)

        self.ordiniPopupMenu = self.builder.get_object("ordiniPopupMenu")
        self.modificaOrdineMenuItemActivate = self.builder.get_object("on_modificaOrdineMenuItem_activate")
        self.suppletivoMenuItem = self.builder.get_object("suppletivoMenuItem")
        self.delSuppletivoMenuItem = self.builder.get_object("delSuppletivoMenuItem")
        self.exportExcelMenuItem = self.builder.get_object("exportExcelMenuItem")
        self.u88FaxMenuItem = self.builder.get_object("u88FaxMenuItem")
        self.inviaOrdineMenuItem = self.builder.get_object("inviaOrdineMenuItem")
        self.u88FaxSuppletivoMenuItem = self.builder.get_object("u88FaxSuppletivoMenuItem")
        self.filtroCombobox = self.builder.get_object("filtroCombobox")

        self.inviaOrdineToolbutton = self.builder.get_object("inviaOrdineToolbutton")
        self.ricezioneOrdineToolbutton = self.builder.get_object("ricezioneOrdineToolbutton")
        self.pianoLevataToolbutton = self.builder.get_object("pianoLevataToolbutton")
        self.ordiniTreeview = self.builder.get_object('ordiniTreeView')
        self.ricezioneOrdineMenuItem = self.builder.get_object("ricezioneOrdineMenuItem")
        self.ordiniTreeview = self.builder.get_object("ordiniTreeview")

        self.pianoLevataToolbutton.set_sensitive(prefs.pianoConsegneDaSito)

        if prefs.pianoConsegneDaSito:
            self.popover = Gtk.Popover()
            self.levataTreeview = Gtk.TreeView()
            levata_formats = {0: "%a %-d %b - %H:%M", 1: "%a %-d %b %Y"}
            levataModelInfoList = [("Data limite", "date"), ("Consegna", "date"), ("", "str"), ("", "str")]
            utility.ExtTreeView(levataModelInfoList, self.levataTreeview, formats=levata_formats)
            self.levataTreeview.get_selection().set_mode(Gtk.SelectionMode.NONE)
            self.updatePianoLevatePopover()
            self.popover.add(self.levataTreeview)
            self.popover.set_position(Gtk.PositionType.BOTTOM)
            self.pianoLevataToolbutton.connect("clicked", self.__showLevataPopover)

        ordini_formats = {self.PESO: "%.3f kg", self.DATA: "%a %-d %b - %H:%M", self.CONSEGNA: "%a %-d %b %Y", self.SUPPLETIVO: "%a %-d %b %Y"}
        utility.ExtTreeView(
            self.modelInfoList, self.ordiniTreeview, formats=ordini_formats,
            properties={self.DATA: {"xalign": 0.5, "scale": utility.PANGO_SCALE_SMALL},
                        self.CONSEGNA: {"xalign": 0.5},
                        self.STATO: {"xalign": 0.5},
                        self.SUPPLETIVO: {"xalign": 0.5}})
        self.ordiniModel = self.ordiniTreeview.get_model()
        self.__buildCombo()

        self.loadOrders()

        selection = self.ordiniTreeview.get_selection()
        selection.connect('changed', self.selectionChanged)
        self.ordiniTreeview.connect("button-press-event", self.showPopup, self.ordiniPopupMenu, self.modificaOrdineMenuItemActivate)
        self.ordiniTreeview.connect("row-activated", self.editOrder)

        self.filtroCombobox.connect("changed", self.loadOrders)

        self.builder.connect_signals({
            "on_logistaMenuitem_activate": self.runTabacchiDialog,
            "on_inventarioMenuItem_activate": self.printInventario,
            "on_editSuppletivoMenuItem_activate": self.ordineSuppletivo,
            "on_delSuppletivoMenuItem_activate": self.deleteSuppletivo,
            "on_u88FaxSuppletivoMenuItem_activate": self.printU88FaxSuppletivo,
            "on_u88FaxMenuItem_activate": self.printU88Fax,
            "on_inviaOrdineMenuItem_activate": self.sendOrder,
            "on_modificaOrdineMenuItem_activate": self.editOrder,
            "on_visualizzaOrdineMenuItem_activate": self.viewOrder,
            "on_eliminaOrdineMenuItem_activate": self.deleteOrdine,
            "on_ricezioneOrdineMenuItem_activate": self.ricezioneOrdine,
            "on_nuovoOrdineMenuItem_activate": self.newOrder
        })

    def __showLevataPopover(self, widget):
        self.popover.set_relative_to(widget)
        self.popover.show_all()
        self.popover.popup()

    # Aggiorna il Popover con la treeview del Piano Levate
    def updatePianoLevatePopover(self, widget=None):
        model = self.levataTreeview.get_model()
        self.levataTreeview.set_model(None)
        model.clear()
        for row in prefs.pianoConsegneList:
            iterator = model.append()
            model.set_value(iterator, 0, row[1])
            model.set_value(iterator, 1, row[0])
            model.set_value(iterator, 2, row[3])
            model.set_value(iterator, 3, row[5])

        self.levataTreeview.set_model(model)

    def __buildCombo(self):
        oggi = datetime.datetime.now()
        anno1 = (oggi - datetime.timedelta(days=365)).strftime('%Y-%m-%d')
        anno2 = (oggi - datetime.timedelta(days=2 * 365)).strftime('%Y-%m-%d')
        anno3 = (oggi - datetime.timedelta(days=3 * 365)).strftime('%Y-%m-%d')

        filtro_store = Gtk.ListStore(str, str)
        cell = Gtk.CellRendererText()
        self.filtroCombobox.pack_start(cell, True)
        self.filtroCombobox.add_attribute(cell, 'text', 1)
        filtro_store.append([anno1, " 1 anno "])
        filtro_store.append([anno2, " 2 anni "])
        filtro_store.append([anno3, " 3 anni "])
        filtro_store.append(["1970-01-01", " Tutto "])

        self.filtroCombobox.set_model(filtro_store)
        self.filtroCombobox.set_active(0)

    # Aggiorna popup menu e toolbar in base all'elemento selezionato
    def __refreshMenus(self, model, iterator):
        stato = model.get_value(iterator, self.STATO)
        incorso = (stato == ordini.IN_CORSO)
        inviato = (stato == ordini.INVIATO)
        suppletivo = model.get_value(iterator, self.SUPPLETIVO) != ordini.ORDINARIO

        self.u88FaxMenuItem.set_sensitive(incorso)
        self.inviaOrdineMenuItem.set_sensitive(incorso)
        self.inviaOrdineToolbutton.set_sensitive(incorso)
        self.ricezioneOrdineMenuItem.set_sensitive(inviato)
        self.ricezioneOrdineToolbutton.set_sensitive(inviato)

        self.suppletivoMenuItem.set_sensitive(True)
        self.delSuppletivoMenuItem.set_sensitive(suppletivo)
        self.u88FaxSuppletivoMenuItem.set_sensitive(suppletivo)

    def selectionChanged(self, selection):
        model, iterator = selection.get_selected()
        if iterator:
            self.__refreshMenus(model, iterator)

    # Mostra menu popup per la gestione ordini
    def showPopup(self, treeview, event, popupMenu, menuItem):
        if event.button == 3:
            x = int(event.x)
            y = int(event.y)
            pthinfo = treeview.get_path_at_pos(x, y)
            if pthinfo is not None:
                path, col, _, _ = pthinfo
                selection = treeview.get_selection()
                if not selection.path_is_selected(path):
                    treeview.set_cursor(path, col, 0)
                treeview.grab_focus()
                popupMenu.popup(None, None, None, None, event.button, event.time)
            return True

    #
    def ricezioneOrdine(self, widget):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            idOrdine = model.get_value(iterator, self.ID)
            inviato = (model.get_value(iterator, self.STATO) == ordini.INVIATO)
            data = model.get_value(iterator, self.DATA)
            if not inviato:
                msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.ERROR,
                                              buttons=Gtk.ButtonsType.CANCEL, text="Per verificare un ordine, deve essere inviato.")
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()
            else:
                if prefs.defaultBarcode >= 0:
                    ricezioneOrdineDialog = ordini.RicezioneOrdineDialog(self, idOrdine, data)
                    ricezioneOrdineDialog.run()
                    self.refreshOrder(idOrdine, model, iterator)
                else:
                    msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.ERROR,
                                                  buttons=Gtk.ButtonsType.CANCEL, text="Non è stato impostato un lettore di codici a barre.")
                    msgDialog.set_title("Attenzione")
                    msgDialog.run()
                    msgDialog.destroy()

    # Aggiorna stato di un ordine
    def __updateOrder(self, idOrdine, idStato):
        result = False
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute("update ordineTabacchi set Stato = ? where ID = ?", (idStato, idOrdine))
            conn.commit()
        except sqlite3.Error as e:
            if conn:
                conn.rollback()
            utility.gtkErrorMsg(e, self)
        else:
            result = True
            prefs.setDBDirty()
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        return result

    # Carica lista ordini
    def loadOrders(self, widget=None):
        model = self.ordiniModel
        self.ordiniTreeview.set_model(None)
        cursor = None
        conn = None
        iterator = self.filtroCombobox.get_active_iter()
        if iterator:
            dataFiltro = self.filtroCombobox.get_model()[iterator][0]
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute(
                'SELECT O.ID as ID, O.Data as "Data [timestamp]", O.Levata as "Levata [date]", sum(R.Ordine*R.Prezzo) as Costo, sum(R.Ordine) as Ordine, O.Stato, O.Suppletivo, O.DataSuppletivo as "DataSuppletivo [date]" FROM ordineTabacchi O left join rigaOrdineTabacchi R on O.ID = R.ID_Ordine where O.Data > ? group by O.ID order by O.Data desc',
                (dataFiltro,))
            result_set = cursor.fetchall()
            model.clear()
            for row in result_set:
                idOrdine = row["ID"]
                costo = row["Costo"]
                ordine = row["Ordine"]
                stato = row["Stato"]
                data = row["Data"]
                levata = row["Levata"]
                if not costo:
                    costo = 0
                if not ordine:
                    ordine = 0
                iterator = model.append()
                model.set_value(iterator, self.ID, idOrdine)
                model.set_value(iterator, self.DATA, data)
                model.set_value(iterator, self.PESO, ordine)
                model.set_value(iterator, self.IMPORTO, costo)
                model.set_value(iterator, self.STATO, stato)
                model.set_value(iterator, self.CONSEGNA, levata)
                model.set_value(iterator, self.SUPPLETIVO, row["Suppletivo"])
                model.set_value(iterator, self.DATA_SUPPLETIVO, row["DataSuppletivo"])
            conn.commit()
        except sqlite3.Error as e:
            if conn:
                conn.rollback()
            utility.gtkErrorMsg(e, self)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        self.ordiniTreeview.set_model(model)

    # Aggiorna informazioni su ordine
    def refreshOrder(self, idOrdine, model, iterator=None):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute(
                "SELECT O.Data as 'Data [timestamp]', O.Levata as 'Levata [date]', sum(R.Ordine*R.Prezzo) as Costo, sum(R.Ordine) as Ordine, O.Stato, O.Suppletivo, O.DataSuppletivo as 'DataSuppletivo [date]' FROM ordineTabacchi O left join rigaOrdineTabacchi R on O.ID = R.ID_Ordine where O.ID = ?",
                (idOrdine,))
            result_set = cursor.fetchall()
            if len(result_set) > 0:
                row = result_set[0]
                costo = row["Costo"]
                ordine = row["Ordine"]
                if not costo:
                    costo = 0
                if not ordine:
                    ordine = 0
                if not iterator:
                    iterator = model.insert(0)
                model.set_value(iterator, self.ID, idOrdine)
                model.set_value(iterator, self.DATA, row["Data"])
                model.set_value(iterator, self.PESO, ordine)
                model.set_value(iterator, self.IMPORTO, costo)
                model.set_value(iterator, self.STATO, row["Stato"])
                model.set_value(iterator, self.CONSEGNA, row["Levata"])
                model.set_value(iterator, self.SUPPLETIVO, row["Suppletivo"])
                model.set_value(iterator, self.DATA_SUPPLETIVO, row["DataSuppletivo"])
            self.__refreshMenus(model, iterator)
        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    # Estrae dalla clipboard informazioni sullo stato dell'ultimo ordine e sulle prossime levate
    # e aggiorna la lista del piano Consegne nelle preferenze
    def __parsePianoLevateFromClipboard(self):
        clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
        resultStr = clipboard.wait_for_text()
        if resultStr:
            log.debug("[CLIPBOARD RESULT] '%s'" % resultStr)
            clipboard.set_text("", -1)
            resultList = resultStr.splitlines()
            if len(resultList) > 0:
                prefs.pianoConsegneList[:] = []
                for res in resultList:
                    resList = res.split('\t')
                    if len(resList) == 6:
                        consegna = resList[0]
                        dataLimite = resList[1]
                        ordine = resList[2]
                        stato = resList[3]
                        canale = resList[4]
                        tipo = resList[5]
                        try:
                            consegna = datetime.datetime.strptime(consegna, "%d/%m/%Y").date()
                            dataLimite = datetime.datetime.strptime(dataLimite, "%d/%m/%Y - %H:%M")
                        except ValueError:
                            pass
                        else:
                            prefs.pianoConsegneList.append([consegna, dataLimite, ordine, stato, canale, tipo])

                        log.debug("Consegna='%s'  DataLimite='%s' ordine='%s' stato='%s' canale='%s' tipo='%s'" %
                                  (consegna, dataLimite, ordine, stato, canale, tipo))
                self.updatePianoLevatePopover()

    # Callback usata per estrarre informazioni sullo stato dell'ultimo ordine e sulle prossime levate

    def __pianoLevateCallback(self):
        self.__parsePianoLevateFromClipboard()
        ordiniModel, ordiniIter = self.ordiniTreeview.get_selection().get_selected()
        if ordiniIter:
            consegnaToCheck = ordiniModel.get_value(ordiniIter, self.CONSEGNA)
            model = self.levataTreeview.get_model()
            iterator = model.get_iter_first()
            consegna = None
            while (iterator is not None) and (consegna != consegnaToCheck):
                consegna = model.get_value(iterator, 1)
                stato = model.get_value(iterator, 2)
                tipo = model.get_value(iterator, 3)
                log.debug("consegnaToCheck=%s consegna=%s stato=%s, tipo=%s" % (consegnaToCheck, consegna, stato, tipo))
                iterator = model.iter_next(iterator)

            # Se l'ordine è stato realmente inviato aggiorna il DB e la lista degli ordini
            if (consegna == consegnaToCheck) and (stato in ordini.STATI_VALIDI):
                log.debug("Lo stato '%s' è valido" % (stato))
                if self.__updateOrder(ordiniModel.get_value(ordiniIter, self.ID), ordini.INVIATO):
                    ordiniModel.set_value(ordiniIter, self.STATO, ordini.INVIATO)
            else:
                log.debug("Lo stato '%s' NON è valido" % (stato))

    # Script JQuery per il login su sito Logista
    LOGIN_SCRIPT = "$('#txtLoginUsername').val('%s'); $('#txtLoginUsername').change(); $('#txtLoginPassword').val('%s'); $('#txtLoginPassword').change(); document.getElementsByClassName('button')[0].click();"

    #
    PIANO_LEVATA_SCRIPT = "document.getElementsByClassName('btn-white')[1].click();"

    # Restituisce le informazioni ottenute dal piano di levata
    # inviandole alla clipboard
    GET_PIANO_LEVATA_SCRIPT = """
        elem = document.getElementById('elencoPianoLevata');
        range = document.createRange();
        sel = window.getSelection();
        sel.removeAllRanges();
        range.selectNodeContents(elem);
        sel.addRange(range);
        document.execCommand("copy");
        """

    GET_PIANO_LEVATA_TABLE_SCRIPT = browserWebkit2.createAsyncScript("$('#panelPianoLevata').is(':visible')", GET_PIANO_LEVATA_SCRIPT)

    #
    def updatePianoLevate(self):
        browserWindow = Browser(self)
        # Script da eseguire all'apertura dell'url
        scripts = [
            ("Recupero informazioni stato ordine..", None, self.GET_PIANO_LEVATA_TABLE_SCRIPT, True, self.__parsePianoLevateFromClipboard),
            ("Mostra elenco degli ultimi ordini..", "https://webt.logistaitalia.it/Home/Index", self.PIANO_LEVATA_SCRIPT, False, None),
            ("Logging into webt.logistaitalia.it", prefs.loginUrl, self.LOGIN_SCRIPT % (prefs.tabacchiUser, prefs.tabacchiPwd), False, None)
        ]
        browserWindow.open(prefs.loginUrl, scripts)

    # Invia l'ordine in modo semi-automatico tramite browser

    def sendOrder(self, widget):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            orderPathname = self.makeOrder(model, iterator)
            if orderPathname:
                base64ExcelData = ''
                # Codifica il file Excel in base64
                with open(orderPathname, "rb") as excel_file:
                    base64ExcelData = base64.b64encode(excel_file.read())

                browserWindow = Browser(self)

                # Script per l'upload dell'ordine in formato Excel
                __ADD_EXCEL_ORDER_SCRIPT = """$('a[data-bind*="CaricaExcel.bind($data, \\\'ORD"]').click();"""

                # Genera il file Excel con l'ordine ed invoca l'evento 'drop'
                __DROP_FILE_SCRIPT = """
                const base64str = '%s';
                const raw = atob(base64str);
                const len = raw.length;
                var b64data = new Uint8Array(len);
                for (var i = 0; i < len; i++) { b64data[i] = raw.charCodeAt(i); }
                var dragElem = document.getElementsByClassName('filedrag')[0];
                var dropEvent = new Event('drop');
                const excelFile = new File([b64data], '%s', {type: 'application/vnd.ms-excel'});
                dropEvent.dataTransfer = {files : [excelFile]};
                dragElem.dispatchEvent(dropEvent);
                """

                ADD_EXCEL_ORDER_SCRIPT = browserWebkit2.createAsyncScript("!$('.modal.fade').is(':visible')", __ADD_EXCEL_ORDER_SCRIPT)

                DROP_FILE_SCRIPT = browserWebkit2.createAsyncScript("$('#panelCaricaExcel').is(':visible')", __DROP_FILE_SCRIPT)

                __SEND_FILE_SCRIPT = """setTimeout(function(){
                $('button:contains("Carica")').click();
                }, 1000);"""

                SEND_FILE_SCRIPT = browserWebkit2.createAsyncScript("$('span.custom-file-input-clear-button').length == 1", __SEND_FILE_SCRIPT, 100)

                # Script da eseguire all'apertura dell'url
                scripts = [
                    ("", None, SEND_FILE_SCRIPT, True, None),
                    ("Drop Excel file..", "https://webt.logistaitalia.it/Ordine", DROP_FILE_SCRIPT % (base64ExcelData, os.path.basename(orderPathname)), True, None),
                    ("Place an order..", "https://webt.logistaitalia.it/Home/Index", ADD_EXCEL_ORDER_SCRIPT, True, None),
                    ("Logging into webt.logistaitalia.it..", prefs.loginUrl, self.LOGIN_SCRIPT % (prefs.tabacchiUser, prefs.tabacchiPwd), False, None)
                ]

                # Script da eseguire alla chiusura del browser
                closeScripts = [
                    ("Recupero informazioni stato ordine..", None, self.GET_PIANO_LEVATA_TABLE_SCRIPT, True, self.__pianoLevateCallback),
                    ("Mostra elenco degli ultimi ordini..", None, self.PIANO_LEVATA_SCRIPT, False, None)
                ]

                browserWindow.open(prefs.loginUrl, scripts, closeScripts)

        else:
            msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.OK, text="Devi selezionare un ordine.")
            msgDialog.format_secondary_text("Non è stato selezionato alcun ordine")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()

    # Effettua i controlli e produce l'ordine in formato Excel
    def makeOrder(self, model, iterator):
        orderPathname = None
        ordineList = None
        errorList = None
        idOrdine = model.get_value(iterator, self.ID)
        data = model.get_value(iterator, self.DATA)
        levata = model.get_value(iterator, self.CONSEGNA)
        stato = model.get_value(iterator, self.STATO)

        if stato == ordini.IN_CORSO:
            (dataLimite, levata) = preferencesTabacchi.dataLimiteOrdine(data)
            log.debug("DATA %s DATALIMITE: %s " % (data, dataLimite))
            if self.__checkDataLevata(model, iterator, dataLimite, levata, idOrdine):
                cursor = None
                conn = None
                try:
                    conn = prefs.getConn()
                    cursor = prefs.getCursor(conn)

                    # Legge la lista del tabacco presente in magazzino (quello ordinato è il sottoinsieme con peso > 0)
                    cursor.execute("SELECT ID, Descrizione, Ordine, Giacenza FROM rigaOrdineTabacchi where ID_Ordine = ? and Ordine > 0 order by Descrizione", (idOrdine,))
                    ordineList = cursor.fetchall()

                    # Legge la lista dei record dei tabacchi
                    cursor.execute("SELECT ID, Descrizione, UnitaMin, PrezzoKg, LivelloMin, Tipo FROM tabacchi where InMagazzino order by Tipo desc, Descrizione")
                    tabaccoList = cursor.fetchall()
                except sqlite3.Error as e:
                    utility.gtkErrorMsg(e, self)
                finally:
                    if cursor:
                        cursor.close()
                    if conn:
                        conn.close()

                ordineDict = dict()
                for ordine in ordineList:
                    ordineDict[ordine["ID"]] = [ordine["Giacenza"], ordine["Ordine"]]

                errorList = []

                # Controllo Sigarette non ordinate per errore
                for tabacco in tabaccoList:
                    idt = tabacco["ID"]
                    if (tabacco["LivelloMin"] > 0):
                        if (idt not in ordineDict):
                            errorList.append([tabacco["Descrizione"]])
                        elif (ordineDict[idt][0] == 0) and (ordineDict[idt][1] == 0):
                            errorList.append([tabacco["Descrizione"]])

                response = Gtk.ResponseType.NO

                if len(errorList) > 0:
                    modelInfo = [("Descrizione", "str")]
                    extMsgDialog = utility.ExtMsgDialog(
                        self, modelInfo, "Alcuni articoli con livello maggiore di 0 non sono stati ordinati.", "Attenzione", "dialog-warning-symbolic",
                        buttons=utility.ExtMsgDialog.YES_NO)
                    extMsgDialog.setSecondaryLabel("Vuoi generare ugualmente l'ordine?")
                    extMsgDialog.setData(errorList)
                    response = extMsgDialog.run()

                # Si procede con la generazione dell'ordine
                if (len(errorList) == 0) or (response == Gtk.ResponseType.YES):
                    try:
                        conn = prefs.getConn()
                        cursor = prefs.getCursor(conn)

                        # Legge la lista del tabacco presente in magazzino (quello ordinato è il sottoinsieme con peso > 0)
                        cursor.execute("SELECT ID, Descrizione, Ordine FROM rigaOrdineTabacchi where ID_Ordine = ? and Ordine > 0 order by Descrizione", (idOrdine,))
                        ordineList = cursor.fetchall()
                        book = xlwt.Workbook()
                        sheet1 = book.add_sheet('Sheet 1')
                        sheet1.write(0, 0, "Codice AAMS")
                        sheet1.write(0, 1, "Peso")
                        sheet1.write(0, 2, "Descrizione")

                        c = 1
                        for ordine in ordineList:
                            sheet1.write(c, 0, ordine["ID"])
                            #  Per evitare problemi di rappresentazione di numeri vicini allo zero
                            sheet1.write(c, 1, round(ordine["Ordine"], 3))
                            sheet1.write(c, 2, ordine["Descrizione"])
                            c += 1
                        orderPathname = "%s/OrdineTabacchi_%s.xls" % (tempfile.gettempdir(), data.strftime("%Y%m%d_%H%M"))
                        book.save(orderPathname)
                    except sqlite3.Error as e:
                        utility.gtkErrorMsg(e, self)
                        if conn:
                            conn.rollback()
                    except Exception as e:
                        utility.gtkErrorMsg(e, self)
                    finally:
                        if cursor:
                            cursor.close()
                        if conn:
                            conn.close()
        else:
            if stato == ordini.RICEVUTO:
                errMsg = "L'ordine è già stato ricevuto."
            elif stato == ordini.INVIATO:
                errMsg = "L'ordine è già stato inviato."

            msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.OK, text="Non è stato generato alcun ordine.")
            msgDialog.format_secondary_text(errMsg)
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()

        return orderPathname

    # Cancella ordine tabacchi
    def deleteOrdine(self, widget):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            data = model.get_value(iterator, self.DATA)
            in_corso = (model.get_value(iterator, self.STATO) == ordini.IN_CORSO)
            response = Gtk.ResponseType.NO
            if not in_corso:
                msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING,
                                              buttons=Gtk.ButtonsType.YES_NO, text="Questo ordine è già stato inviato.")
                msgDialog.format_secondary_text("Vuoi veramente eliminarlo?")
                msgDialog.set_title("Attenzione")
                response = msgDialog.run()
                msgDialog.destroy()

            if in_corso or (response == Gtk.ResponseType.YES):
                msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.QUESTION,
                                              buttons=Gtk.ButtonsType.YES_NO, text="Stai per eliminare l'ordine di %s" % data.strftime("%a %d %b %Y alle %H:%M"))
                suppletivo = model.get_value(iterator, self.SUPPLETIVO)
                if (suppletivo != ordini.ORDINARIO):
                    secondary_text = "Sarà cancellato anche l'ordine {}. Sei sicuro?".format(ordini.TIPO_ORDINE[suppletivo])
                else:
                    secondary_text = "Sei sicuro?"
                msgDialog.format_secondary_text(secondary_text)
                msgDialog.set_title("Elimina ordine")
                response = msgDialog.run()
                msgDialog.destroy()
                if response == Gtk.ResponseType.YES:
                    _id = model.get_value(iterator, self.ID)
                    cursor = None
                    conn = None
                    try:
                        conn = prefs.getConn()
                        cursor = prefs.getCursor(conn)
                        cursor.execute("delete from rigaOrdineSuppletivo where ID_Ordine = ?", (_id,))
                        cursor.execute("delete from verificaOrdine where ID_ordine = ?", (_id,))
                        cursor.execute("delete from rigaOrdineTabacchi where ID_Ordine = ?", (_id,))
                        cursor.execute("delete from ordineTabacchi where ID = ?", (_id,))
                        conn.commit()
                        self.ordiniModel.remove(iterator)
                        prefs.setDBDirty()
                    except sqlite3.Error as e:
                        utility.gtkErrorMsg(e, self)
                        if conn:
                            conn.rollback()
                    finally:
                        if cursor:
                            cursor.close()
                        if conn:
                            conn.close()

    # Mostra la dialog con il catologo Logista
    def runTabacchiDialog(self, other):
        tabacchiDialog = TabacchiDialog(self)
        tabacchiDialog.run()

    # Crea o modifica ordine suppletivo
    def ordineSuppletivo(self, widget):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            stato = model.get_value(iterator, self.STATO)
            if (stato == ordini.IN_CORSO):
                msgDialog = Gtk.MessageDialog(
                    parent=self, modal=True, message_type=Gtk.MessageType.ERROR, buttons=Gtk.ButtonsType.CANCEL,
                    text="Non è possibile aggiungere un ordine suppletivo ad un ordine in corso.")
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()
            else:
                idOrdine = model.get_value(iterator, self.ID)
                tipo = model.get_value(iterator, self.SUPPLETIVO)
                data = model.get_value(iterator, self.DATA_SUPPLETIVO)
                suppletivoDialog = ordini.SuppletivoDialog(self, idOrdine, tipo, data)
                suppletivoDialog.run()
                self.refreshOrder(idOrdine, model, iterator)

    # Cancella ordine suppletivo
    def deleteSuppletivo(self, widget):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            tipo = model.get_value(iterator, self.SUPPLETIVO)
            if tipo == ordini.ORDINARIO:
                msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.ERROR,
                                              buttons=Gtk.ButtonsType.CANCEL, text="Non esiste un ordine suppletivo da cancellare.")
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()
            else:
                msgDialog = Gtk.MessageDialog(
                    parent=self, modal=True, message_type=Gtk.MessageType.QUESTION, buttons=Gtk.ButtonsType.YES_NO,
                    text="Stai per cancellare un ordine suppletivo %s" % ordini.TIPO_ORDINE[tipo])
                msgDialog.format_secondary_text("Sei sicuro?")
                msgDialog.set_title("Cancella ordine suppletivo")
                response = msgDialog.run()
                msgDialog.destroy()
                if response == Gtk.ResponseType.YES:
                    _id = model.get_value(iterator, self.ID)
                    cursor = None
                    conn = None
                    try:
                        conn = prefs.getConn()
                        cursor = prefs.getCursor(conn)
                        cursor.execute("delete from rigaOrdineSuppletivo where ID_Ordine = ?", (_id,))
                        cursor.execute("Update ordineTabacchi set DataSuppletivo = Null, Suppletivo = ? where ID = ?", (ordini.ORDINARIO, _id))
                        conn.commit()
                    except sqlite3.Error as e:
                        utility.gtkErrorMsg(e, self)
                        if conn:
                            conn.rollback()
                    else:
                        self.refreshOrder(_id, model, iterator)
                        prefs.setDBDirty()
                    finally:
                        if cursor:
                            cursor.close()
                        if conn:
                            conn.close()

    # Nuovo Ordine
    def newOrder(self, widget):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)

            # Cerco un ordine non ancora ricevuto
            cursor.execute("select ID from ordineTabacchi where Stato != ?", (ordini.RICEVUTO,))
            rows = cursor.fetchall()

            # Se non esiste proseguo...
            if not rows:
                # Cerco l'ultimo ordine effettuato
                cursor.execute("select ID, Data as 'Data [timestamp]' from ordineTabacchi order by Data desc")
                rows = cursor.fetchall()

                if rows and (datetime.datetime.now() < rows[0]["Data"]):
                    msgDialog = Gtk.MessageDialog(
                        parent=self, modal=True, message_type=Gtk.MessageType.WARNING, buttons=Gtk.ButtonsType.OK,
                        text="Ci sono ordini con date successive. Un nuovo ordine deve essere il più recente.")
                    msgDialog.set_title("Attenzione")
                    msgDialog.run()
                    msgDialog.destroy()
                else:
                    ordineDialog = ordini.OrdineDialog(self)
                    if (not ordineDialog.error):
                        ordineDialog.run()
                        self.refreshOrder(ordineDialog.ordineID, self.ordiniModel)
                        self.ordiniTreeview.set_cursor(0)
            else:
                msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING,
                                              buttons=Gtk.ButtonsType.OK, text="Non è possibile creare un nuovo ordine.")
                msgDialog.format_secondary_text("Esiste un ordine in corso o non ancora ricevuto.")
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()

        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    # Controlla se la lista del piano consegne è vuota e chiede di aggiornarla
    def checkUpdatePianoLevata(self):
        if prefs.pianoConsegneDaSito and len(prefs.pianoConsegneList) == 0:
            msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.QUESTION,
                                          buttons=Gtk.ButtonsType.OK_CANCEL, text="Vuoi aggiornare il piano levate dal sito Logista?")
            msgDialog.set_title("Aggiorna piano levate")
            response = msgDialog.run()
            msgDialog.destroy()
            if (response == Gtk.ResponseType.OK):
                self.updatePianoLevate()

    # Controlla la data dell'ordine e il giorno di levata se sono coerenti,
    # altrimenti chiede di modificarli con i primi disponibili
    def __checkDataLevata(self, model, iterator, dataLimite, levata, idOrdine):
        result = True
        newData = datetime.datetime.now().replace(microsecond=0)
        if dataLimite < newData:
            msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING, buttons=Gtk.ButtonsType.OK_CANCEL,
                                          text="L'ordine andava inviato %s. Vuoi spostarlo?" % dataLimite.strftime("%A %d/%m entro le %H:%M"))
            (dataLimite, levata) = preferencesTabacchi.dataLimiteOrdine(newData)
            msgDialog.format_secondary_text("Nuovo invio %s e levata %s." % (dataLimite.strftime("%A %d/%m"), levata.strftime("%A %d/%m")))
            msgDialog.set_title("Attenzione")
            response = msgDialog.run()
            msgDialog.destroy()
            if response != Gtk.ResponseType.OK:
                result = False
            else:
                cursor = None
                conn = None
                try:
                    conn = prefs.getConn()
                    cursor = prefs.getCursor(conn)
                    cursor.execute("update ordineTabacchi set Data=?, Levata=? where ID=?", (newData, levata, idOrdine))
                    conn.commit()
                except sqlite3.Error as e:
                    result = False
                    utility.gtkErrorMsg(e, self)
                    if conn:
                        conn.rollback()
                else:
                    model.set_value(iterator, self.DATA, newData)
                    model.set_value(iterator, self.CONSEGNA, levata)
                    prefs.setDBDirty()
                finally:
                    if cursor:
                        cursor.close()
                    if conn:
                        conn.close()

        return result

    # Modifica Ordine
    def editOrder(self, widget, arg1=None, arg2=None):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            idOrdine = model.get_value(iterator, self.ID)
            stato = model.get_value(iterator, self.STATO)
            data = model.get_value(iterator, self.DATA)
            (dataLimite, levata) = preferencesTabacchi.dataLimiteOrdine(data)

            if stato in [ordini.IN_CORSO]:
                mode = ordini.EDIT_MODE
                if not self.__checkDataLevata(model, iterator, dataLimite, levata, idOrdine):
                    return
            else:
                mode = ordini.REVIEW_MODE
                msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING,
                                              buttons=Gtk.ButtonsType.OK, text="Questo ordine non è più modificabile.")
                msgDialog.format_secondary_text("Poteva esserlo entro le %s" % dataLimite.strftime("%H:%M di %A %d %B %Y"))
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()

            ordineDialog = ordini.OrdineDialog(self, idOrdine, mode)
            ordineDialog.run()
            self.refreshOrder(idOrdine, model, iterator)

    # Visualizza Ordine
    def viewOrder(self, widget, arg1=None, arg2=None):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            idOrdine = model.get_value(iterator, self.ID)
            ordineDialog = ordini.OrdineDialog(self, idOrdine, ordini.VIEW_MODE)
            ordineDialog.run()

    # Stampa ordine suppletivo su modulo U88-Fax (urgente o straordinario)
    def printU88FaxSuppletivo(self, widget):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            tipo = model.get_value(iterator, self.SUPPLETIVO)
            if (tipo != ordini.ORDINARIO):
                data = model.get_value(iterator, self.DATA_SUPPLETIVO)
                _id = model.get_value(iterator, self.ID)
                try:
                    conn = prefs.getConn()
                    cursor = prefs.getCursor(conn)
                    cursor.execute("select r.ID, r.Ordine from rigaOrdineSuppletivo r where ID_Ordine =  ? and r.Ordine > 0", (_id,))
                    orderList = cursor.fetchall()
                    stampe.printU88Fax(self, orderList, tipo, data)
                except sqlite3.Error as e:
                    utility.gtkErrorMsg(e, self)
                finally:
                    if cursor:
                        cursor.close()
                    if conn:
                        conn.close()

    # Stampa ordine su modulo U88-Fax
    def printU88Fax(self, widget):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            _id = model.get_value(iterator, self.ID)
            stato = model.get_value(iterator, self.STATO)
            if (stato == ordini.IN_CORSO):
                try:
                    conn = prefs.getConn()
                    cursor = prefs.getCursor(conn)
                    cursor.execute("select r.ID, r.Ordine from rigaOrdineTabacchi r where ID_Ordine =  ? and r.Ordine > 0", (_id,))
                    orderList = cursor.fetchall()
                    u88Dialog = U88Dialog(self, model.get_value(iterator, self.CONSEGNA))
                    result = u88Dialog.run()
                    if result:
                        stampe.printU88Fax(self, orderList, result[0], result[1])
                except sqlite3.Error as e:
                    utility.gtkErrorMsg(e, self)
                finally:
                    if cursor:
                        cursor.close()
                    if conn:
                        conn.close()

        else:
            msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.OK, text="Devi selezionare un ordine.")
            msgDialog.format_secondary_text("Non è stato stampato il modulo U88-Fax.")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()

    # Stampa Inventario e valorizzazione Articoli
    def printInventario(self, widget):
        model, iterator = self.ordiniTreeview.get_selection().get_selected()
        if iterator:
            idOrdine = model.get_value(iterator, self.ID)
            data = model.get_value(iterator, self.DATA)

            inventarioReport = stampe.InventarioReport(self, idOrdine, data)
            fileObj = inventarioReport.build()
            if fileObj:
                inventarioReport.show(fileObj)
            else:
                msgDialog = Gtk.MessageDialog(parent=self, modal=True, message_type=Gtk.MessageType.WARNING,
                                              buttons=Gtk.ButtonsType.OK, text="Non ci sono articoli da mostrare")
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()

# Applicazione principale


class MainApplication(Gtk.Application):
    # Main initialization routine
    def __init__(self, application_id, flags):
        super().__init__(application_id=application_id, flags=flags)
        self.mainWindow = None

    def do_startup(self):
        Gtk.Application.do_startup(self)

        # Legge le preferenze
        try:
            prefs.load()
        except Exception as e:
            utility.gtkErrorMsg(e, None)

        action = Gio.SimpleAction.new("about", None)
        action.connect("activate", self.on_about)
        self.add_action(action)

        action = Gio.SimpleAction.new("preferences", None)
        action.connect("activate", self.gestionePreferenze)
        self.add_action(action)

        action = Gio.SimpleAction.new("importa", None)
        action.connect("activate", self.importLogistaDoc)
        self.add_action(action)

        action = Gio.SimpleAction.new("ricalcola", None)
        action.connect("activate", self.ricalcolaConsumi)
        self.add_action(action)

        action = Gio.SimpleAction.new("statistiche", None)
        action.connect("activate", self.globalStats)
        self.add_action(action)

        action = Gio.SimpleAction.new("etichette", None)
        action.connect("activate", self.printLabels)
        self.add_action(action)

        action = Gio.SimpleAction.new("articoli", None)
        action.connect("activate", self.printElencoArticoli)
        self.add_action(action)

        action = Gio.SimpleAction.new("excel", None)
        action.connect("activate", self.generaExcel)
        self.add_action(action)

        action = Gio.SimpleAction.new("ordine_excel", None)
        action.connect("activate", self.generaOrdineExcel)
        self.add_action(action)

        action = Gio.SimpleAction.new("quit", None)
        action.connect("activate", self.on_quit)
        self.add_action(action)

    def do_activate(self):
        if not self.mainWindow:
            self.mainWindow = MainWindow(application=self, title=config.__desc__)
            self.mainWindow.connect("delete-event", self.on_quit)
            self.mainWindow.set_size_request(600, 500)
            self.mainWindow.set_position(Gtk.WindowPosition.CENTER)

        self.mainWindow.show_all()
        self.mainWindow.checkUpdatePianoLevata()

    def on_about(self, action, param):
        aboutDialog = Gtk.AboutDialog(transient_for=self.mainWindow, modal=True)
        aboutDialog.set_program_name(config.__desc__)
        aboutDialog.set_version(config.__version__)
        aboutDialog.set_copyright(config.__copyright__)
        aboutDialog.set_logo(self.mainWindow.logo)
        aboutDialog.present()

    # Impostazioni preferenze
    def gestionePreferenze(self, action, param):
        preferencesDialog = preferencesTabacchi.PreferencesDialog(self.mainWindow)
        response = preferencesDialog.run()
        if response == Gtk.ResponseType.OK:
            self.mainWindow.pianoLevataToolbutton.set_sensitive(prefs.pianoConsegneDaSito)

    # Ricalcola i consumi
    def ricalcolaConsumi(self, action, param):
        thread = RicalcolaConsumiThread()
        progressDialog = utility.ProgressDialog(self.mainWindow, "Ricalcolo vendite in corso..",
                                                "In base agli acquisti e alle rimanenze.", "Ricalcolo vendite", thread)
        progressDialog.start()

    # Statistiche generali
    def globalStats(self, action, param):
        statsDialog = stats.GlobalStatsDialog(self.mainWindow)
        statsDialog.run()

    # Genera file Excel per inventario e valorizzazione Magazzino tabacchi
    def generaExcel(self, action, param):
        HEADER_STYLE = 'font: height 200, bold on; align: vert centre, horiz center; borders: bottom thin;'
        CURRENCY_STYLE = xlwt.easyxf('font: height 200; align: vert centre, horiz right;', num_format_str=u'[$\u20ac-1] #,##0.00')
        CURRENCY_STYLE_BOLD = xlwt.easyxf('font: height 260, bold on; align: vert centre, horiz right;', num_format_str=u'[$\u20ac-1] #,##0.00')
        NUM_STYLE_ED = xlwt.easyxf('protection: cell_locked false; font: height 260; align: vert centre, horiz right;', num_format_str='0')
        NUM_STYLE = xlwt.easyxf('font: height 260; align: vert centre, horiz right;', num_format_str='0')
        STR_STYLE = xlwt.easyxf('font: height 200; align: vert centre, horiz left;')
        STR_STYLE_BOLD_SMALL = xlwt.easyxf('font: height 200, bold on, underline single; align: vert centre, horiz centre;')

        QUERY = "SELECT ID, Descrizione, unitaMin, pezziUnitaMin, prezzoKG, tipo FROM tabacchi WHERE InMagazzino ORDER BY Tipo desc,Descrizione"
        SHEET_NAME = "tabacchi"

        write_book = xlwt.Workbook(encoding='UTF-8')
        cur_sheet = write_book.add_sheet(SHEET_NAME)

        cur_sheet.protect = True  # Il foglio e' tutto protetto per default
        cur_sheet.password = SHEET_NAME

        cur_sheet.col(2).set_style(CURRENCY_STYLE)
        cur_sheet.col(3).set_style(CURRENCY_STYLE)
        cur_sheet.col(6).set_style(CURRENCY_STYLE)

        cur_sheet.col(0).width = 256 * 8
        cur_sheet.col(1).width = 256 * 52
        cur_sheet.col(2).width = 256 * 18
        cur_sheet.col(3).width = 256 * 16
        cur_sheet.col(4).width = 256 * 16
        cur_sheet.col(5).width = 256 * 16
        cur_sheet.col(6).width = 256 * 16
        cur_sheet.col(7).width = 256 * 19

        cur_sheet.write(0, 0, "Codice", xlwt.easyxf(HEADER_STYLE))
        cur_sheet.write(0, 1, "Descrizione", xlwt.easyxf(HEADER_STYLE))
        cur_sheet.write(0, 2, "Pacchetti x conf.", xlwt.easyxf(HEADER_STYLE))
        cur_sheet.write(0, 3, "Prezzo pacchetto", xlwt.easyxf(HEADER_STYLE))
        cur_sheet.write(0, 4, "Confezioni", xlwt.easyxf(HEADER_STYLE))
        cur_sheet.write(0, 5, "Pacchetti", xlwt.easyxf(HEADER_STYLE))
        cur_sheet.write(0, 6, "Valore", xlwt.easyxf(HEADER_STYLE))

        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = conn.cursor()
            cursor.execute(QUERY)
            result_set = cursor.fetchall()
            y = 0

            for row in result_set:
                y += 1
                cur_sheet.row(y).height = 320
                cur_sheet.write(y, 0, row["ID"], STR_STYLE)
                cur_sheet.write(y, 1, row["Descrizione"], STR_STYLE)
                pzUnitaMin = row["pezziUnitaMin"]
                cur_sheet.write(y, 2, pzUnitaMin, NUM_STYLE)
                prezzo = row["prezzoKG"] * row["unitaMin"]

                if pzUnitaMin > 0:
                    cur_sheet.write(y, 3, prezzo / pzUnitaMin, CURRENCY_STYLE)
                else:
                    cur_sheet.write(y, 3, 0, CURRENCY_STYLE)
                cur_sheet.write(y, 4, 0, NUM_STYLE_ED)
                cur_sheet.write(y, 5, 0, NUM_STYLE_ED)

                cur_sheet.write(y, 6, xlwt.Formula('C{0}*E{0}*D{0}+F{0}*D{0}'.format(y + 1)), CURRENCY_STYLE)

        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        if y > 0:
            y += 1
            cur_sheet.row(y + 2).height = 320
            cur_sheet.row(y + 3).height = 320
            cur_sheet.write(y + 2, 6, " TOTALE ", STR_STYLE_BOLD_SMALL)
            cur_sheet.write(y + 3, 6, xlwt.Formula('SUM(G2:G{0})'.format(y)), CURRENCY_STYLE_BOLD)
            cur_sheet.write(y + 5, 6, " Aggio (10%) ", STR_STYLE_BOLD_SMALL)
            cur_sheet.write(y + 6, 6, xlwt.Formula('G{0}*0.10'.format(y + 4)), CURRENCY_STYLE_BOLD)
            cur_sheet.write(y + 8, 6, " TOTALE NETTO ", STR_STYLE_BOLD_SMALL)
            cur_sheet.write(y + 9, 6, xlwt.Formula('G{0}*0.90'.format(y + 4)), CURRENCY_STYLE_BOLD)

        tmpFile = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
        tmpFile.close()

        write_book.save(tmpFile.name)

        # Libera le risorse
        del write_book

        # Apre il file
        Gio.Subprocess.new(["gio", "open", tmpFile.name], Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_MERGE)

    # Genera file Excel come modello ordine
    def generaOrdineExcel(self, action, param):
        QUERY = "SELECT ID, Descrizione, unitaMin, tipo FROM tabacchi WHERE InMagazzino ORDER BY Tipo desc,Descrizione"
        SHEET_NAME = "tabacchi"

        write_book = xlwt.Workbook(encoding='UTF-8')
        cur_sheet = write_book.add_sheet(SHEET_NAME)

        cur_sheet.protect = True  # Il foglio e' tutto protetto per default
        cur_sheet.password = SHEET_NAME

        cur_sheet.write(0, 0, "Codice AAMS")
        cur_sheet.write(0, 1, "Peso")
        cur_sheet.write(0, 2, "Descrizione")

        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = conn.cursor()
            cursor.execute(QUERY)
            result_set = cursor.fetchall()
            y = 0

            for row in result_set:
                y += 1

                cur_sheet.write(y, 0, row["ID"])
                cur_sheet.write(y, 1, row["unitaMin"])
                cur_sheet.write(y, 2, row["Descrizione"])
        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        tmpFile = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
        tmpFile.close()

        write_book.save(tmpFile.name)

        # Libera le risorse
        del write_book

        # Apre il file
        Gio.Subprocess.new(["gio", "open", tmpFile.name], Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_MERGE)

    # Stampa etichette prezzi

    def printLabels(self, action, param):
        labelsDialog = stampe.LabelsDialog(self.mainWindow)
        result = labelsDialog.run()
        if result:
            dataStampa = result[0]

            row = result[1]  # Riga iniziale
            col = result[2]  # Colonna iniziale
            try:
                conn = prefs.getConn()
                cursor = prefs.getCursor(conn)
                cursor.execute("select Descrizione, UnitaMin, PrezzoKG, PezziUnitaMin from tabacchi where InMagazzino and Decorrenza > ?", (dataStampa,))
                labelList = cursor.fetchall()
                stampe.printLabels(self.mainWindow, labelList, dataStampa, row, col)
            except sqlite3.Error as e:
                utility.gtkErrorMsg(e, self.mainWindow)
            finally:
                if cursor:
                    cursor.close()
                if conn:
                    conn.close()

    # Stampa Elenco Articoli
    def printElencoArticoli(self, action, param):
        articoliReport = stampe.ArticoliReport(self.mainWindow)
        fileObj = articoliReport.build()
        articoliReport.show(fileObj)

    # Importa un documento (ordine o fattura) dal portale logista
    def importLogistaDoc(self, action, param):
        # Legge il file
        importDialog = ImportDialog(self.mainWindow)
        result = importDialog.run()

        if (result):
            dataFattura = result[0].date()
            model = result[1]
            tipo = result[2]
            dataOrdine = result[3]
            try:
                conn = prefs.getConn()
                cursor = prefs.getCursor(conn)
                listPrec = None
                listCur = None
                listSucc = None

                # "estrapolo" i livelli dall'ordine precedente
                cursor.execute(
                    "SELECT ID, (Ordine + Giacenza) Livello FROM rigaOrdineTabacchi where ID_ordine = (select ID from ordineTabacchi where Data = (select max(Data) from ordineTabacchi where Data < ? order by Data desc))",
                    (dataOrdine,))
                listPrec = cursor.fetchall()

                # "estrapolo" i livelli dall'ordine successivo
                cursor.execute(
                    "SELECT r.ID, (r.Ordine + r.Giacenza) Livello FROM rigaOrdineTabacchi r where r.ID_ordine = (select ID from ordineTabacchi where Data = (select min(Data) from ordineTabacchi where Data > ? order by Data))",
                    (dataOrdine,))
                listSucc = cursor.fetchall()

                # Nel caso in cui non ci siano info uso i livelli correnti
                cursor.execute("SELECT ID, LivelloMin FROM tabacchi")
                listCur = cursor.fetchall()

                stato = ordini.RICEVUTO if (tipo == importDialog.FATTURA) else ordini.INVIATO

                response = Gtk.ResponseType.OK
                log.debug("IMPORT: dataFattura = %s  dataOrdine = %s" % (dataFattura, dataOrdine))

                # Controllo se esiste già un ordine con la stessa data di levata.. (deve essere unico)
                cursor.execute("SELECT ID from ordineTabacchi where Levata = ?", (dataFattura,))

                listOrders = cursor.fetchall()
                if listOrders and (len(listOrders) > 0):
                    msgDialog = Gtk.MessageDialog(
                        parent=self.mainWindow, modal=True, message_type=Gtk.MessageType.WARNING, buttons=Gtk.ButtonsType.OK_CANCEL,
                        text="Esiste già un ordine con levata %s" % datetime.datetime.strftime(dataFattura, "%A %d %B %Y"))
                    msgDialog.format_secondary_text("Lo sovrascrivo?")
                    msgDialog.set_title("Attenzione")
                    response = msgDialog.run()
                    msgDialog.destroy()
                    if response == Gtk.ResponseType.OK:
                        firstOrder = listOrders[0]
                        idOrdine = firstOrder["ID"]
                        cursor.execute("delete from rigaOrdineTabacchi where ID_Ordine = ?", (idOrdine,))
                        cursor.execute("update ordineTabacchi set Data = ?, Stato = ?, Levata = ? where ID = ?", (dataOrdine, stato, dataFattura, idOrdine))
                else:
                    cursor.execute("UPDATE ordineTabacchi SET Data=?, Stato=? where Levata=?", (dataOrdine, stato, dataFattura))
                    if cursor.rowcount == 0:
                        cursor.execute("INSERT INTO ordineTabacchi(Data, Stato, Levata) VALUES(?, ?, ?)", (dataOrdine, stato, dataFattura))
                    idOrdine = cursor.lastrowid

                if response != Gtk.ResponseType.CANCEL:
                    # Se è possibile fare una stima delle quantità
                    if listPrec and listSucc:
                        precDict = dict()
                        succDict = dict()
                        curDict = dict()

                        for row in listPrec:
                            precDict[row["ID"]] = row["Livello"]
                        for row in listSucc:
                            succDict[row["ID"]] = row["Livello"]
                        for row in listCur:
                            curDict[row["ID"]] = row["LivelloMin"]

                    iterator = model.get_iter_first()
                    while iterator:
                        id_cod = model.get_value(iterator, importDialog.ID_CODICE)
                        descrizione = model.get_value(iterator, importDialog.ID_DESC)
                        peso = model.get_value(iterator, importDialog.ID_PESO)
                        costo = model.get_value(iterator, importDialog.ID_PREZZO_KG)

                        # Se è possibile fare una stima delle quantità
                        if listPrec and listSucc:
                            if id_cod in precDict and id_cod in succDict:
                                livello = min(precDict[id_cod], succDict[id_cod])
                                # log.debug("liv prec:%.2f succ:%.2f " % (precDict[id_cod], succDict[id_cod]))
                            else:
                                livello = curDict[id_cod]
                                # log.debug("liv prec: ? succ: ?  -  cur: %f" % curDict[id_cod])
                            quantita = round(livello, 3) - peso
                            if (quantita < 0):
                                quantita = 0
                        else:
                            quantita = 0
                        log.debug("Ordine:%.2f Prezzo:%.2f " % (peso, costo))
                        cursor.execute(
                            "INSERT INTO rigaOrdineTabacchi(ID, Descrizione, ID_Ordine, Ordine, Prezzo, Giacenza, Consumo) VALUES(?, ?, ?, ?, ?, ?, ?)",
                            (id_cod, descrizione, idOrdine, peso, costo, quantita, 0))
                        iterator = model.iter_next(iterator)
                    conn.commit()
            except sqlite3.Error as e:
                utility.gtkErrorMsg(e, self.mainWindow)
                if conn:
                    conn.rollback()
            else:
                prefs.setDBDirty()
            finally:
                if cursor:
                    cursor.close()
                if conn:
                    conn.close()
            self.mainWindow.loadOrders()

    # Override the default handler for the delete-event signal
    def on_quit(self, event=None, data=None):
        log.debug("Salva le preferenze su file..")
        prefs.save()
        if prefs.checkBackup(self.mainWindow):
            thread = preferencesTabacchi.BackupThread(prefs)
            progressDialog = utility.ProgressDialog(self.mainWindow, "Backup in corso..", "", "Backup", thread)
            progressDialog.setResponseCallback(self.__backupCallback)
            progressDialog.setErrorCallback(self.quit)
            progressDialog.setStopCallback(self.quit)
            progressDialog.start()
            return True  # Evita che i segnali destroy o delete si propaghino prima che il backup sia finito..
        log.debug("Quit !")
        self.quit()

    # Chiude l'applicazione in modo asincrono
    def __backupCallback(self):
        prefs.saveModified(False)
        log.debug("Quit !")
        self.quit()


# Launcher
def start():
    # Inizializza l'applicazione GTK
    application = MainApplication(f"net.guarnie.{config.PACKAGE_NAME}", Gio.ApplicationFlags.FLAGS_NONE)

    # Fa partire la GUI
    application.run(sys.argv)


# Esecuzione da riga di comando
if __name__ == "__main__":
    start()
