#!/usr/bin/python
# coding: UTF-8
#
# Copyright (C) Francesco Guarnieri 2020 <francesco@guarnie.net>
#

import math
import os
import datetime
import tempfile
import locale
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Table, TableStyle, Image
import gi

import utility
import ordini
from preferencesTabacchi import prefs

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk, GLib, Gio  # noqa: E402

A4_BORDERLESS = (21.55 * cm, 30.2 * cm)


# Report Articoli
class ArticoliReport(utility.Report):
    def __init__(self, parent):
        NUM_ROWS = 46
        ROW_HEIGHT = 0.6 * cm
        COL_WIDTHS = (1.6 * cm, 10 * cm, 1.6 * cm, 2 * cm, 2 * cm, 2 * cm)
        COL_DESC = ("Codice", "Descrizione", "Unita min.", "Livello", "Magazzino", "Ordine")
        QUERY = "select ID, Descrizione, LivelloMin, UnitaMin from tabacchi where InMagazzino order by Tipo desc, Descrizione"
        utility.Report.__init__(self, parent, prefs, QUERY, "Elenco articoli", NUM_ROWS, COL_DESC, COL_WIDTHS, ROW_HEIGHT)

    def initialize(self):
        self.setTableHeaderStyle([('VALIGN', 'BOTTOM'), ('ALIGN', 'CENTER')])
        self.setTableStyle([('GRID', 0.5, colors.black), ('VALIGN', 'MIDDLE'), ('ALIGN', 'RIGHT'), ('FONT', 'Helvetica', 10)])
        self.setTableColumnStyle(1, [], [('ALIGN', 'LEFT')])
        self.setTableColumnStyle(2, [], [('FONT', 'Helvetica', 8)])
        self.setTableColumnStyle(3, [], [('FONT', 'Helvetica-Bold', 10)])

    def writeTableRow(self):
        self.writeFieldToCell("ID", 0)
        self.writeFieldToCell("Descrizione", 1)
        livello = self.getField("LivelloMin")
        unita = self.getField("UnitaMin")
        self.writeCell(locale.format_string("%.3f kg", unita), 2)
        self.writeCell(locale.format_string("%.3f kg", livello), 3)


# Report inventario
class InventarioReport(utility.Report):
    def __init__(self, parent, _id, data):
        NUM_ROWS = 60
        ROW_HEIGHT = 0.45 * cm
        COL_WIDTHS = (1.5 * cm, 11 * cm, 2 * cm, 2 * cm, 2.5 * cm)
        COL_DESC = ("ID", "Descrizione", "Magazzino", "Scaffale", "Valore")
        QUERY = "select r.ID, r.Descrizione, r.Giacenza, t.unitaMin, ((r.Giacenza + t.unitaMin) * t.prezzoKG) Valore from tabacchi t, rigaOrdineTabacchi r, ordineTabacchi o where t.ID = r.ID and r.ID_Ordine = o.ID and o.ID = %s and t.InMagazzino order by Valore desc" % _id
        utility.Report.__init__(self, parent, prefs, QUERY, "Magazzino Tabacchi valorizzato a %s" %
                                data.strftime("%A %d %B %Y"), NUM_ROWS, COL_DESC, COL_WIDTHS, ROW_HEIGHT)
        self.totale = 0

    def initialize(self):
        self.setTableHeaderStyle([('VALIGN', 'BOTTOM'), ('ALIGN', 'CENTER'), ('FONT', 'Helvetica', 9)])
        self.setTableStyle([('GRID', 0.5, colors.black), ('VALIGN', 'MIDDLE'), ('ALIGN', 'RIGHT'), ('FONT', 'Helvetica', 9)])
        self.setFooterStyle([('SPAN',)])
        self.setReportFooterStyle([('SPAN',), ('ALIGN', 'RIGHT'), ('VALIGN', 'BOTTOM'), ('FONT', 'Helvetica-Bold', 11)])
        self.setTableColumnStyle(1, [], [('ALIGN', 'LEFT')])
        # self.setTableColumnStyle(2, [], [('FONT', 'Helvetica-Bold', 10)])

    def writeTableRow(self):
        self.writeFieldToCell("ID", 0)
        self.writeFieldToCell("Descrizione", 1)
        qta = self.getField("Giacenza")
        self.writeCell(locale.format_string("%.3f kg", qta), 2)
        unitaMin = self.getField("unitaMin")
        self.writeCell(locale.format_string("%.3f kg", unitaMin), 3)
        valore = self.getField("Valore")
        self.writeCell(locale.currency(valore, True, True), 4)
        self.totale += valore

    def writeReportFooter(self):
        self.writeCell("Totale: %s" % locale.currency(self.totale, True, True), 0)


# Dialog per parametri Stampa delle etichette
class LabelsDialog(utility.GladeWindow):
    def __init__(self, parent):
        super().__init__(parent, "labelsDialog.glade")

        self.labelsDialog = self.builder.get_object("labelsDialog")
        self.labelsDialog.set_transient_for(parent)

        self.prezziCalendar = self.builder.get_object("prezziCalendar")
        self.rigaSpinbutton = self.builder.get_object("rigaSpinbutton")
        self.colonnaSpinbutton = self.builder.get_object("colonnaSpinbutton")

        adjustmentRiga = Gtk.Adjustment(value=1, lower=1, upper=20, step_incr=1)
        self.rigaSpinbutton.set_adjustment(adjustmentRiga)
        adjustmentColonna = Gtk.Adjustment(value=1, lower=1, upper=5, step_incr=1)
        self.colonnaSpinbutton.set_adjustment(adjustmentColonna)

        today = datetime.date.today()
        self.prezziCalendar.select_month(today.month - 1, today.year)
        self.prezziCalendar.select_day(today.day)

        self.builder.connect_signals({"on_labelsDialog_delete_event": self.close,
                                      "on_okButton_clicked": self.okClose,
                                      "on_cancelButton_clicked": self.close,
                                      })

        self.result = Gtk.ResponseType.CANCEL

    def run(self):
        self.labelsDialog.run()
        return self.result

    def okClose(self, widget, other=None):
        year, month, day = self.prezziCalendar.get_date()
        self.result = [datetime.date(year, month + 1, day), int(self.rigaSpinbutton.get_value()) - 1, int(self.colonnaSpinbutton.get_value()) - 1]
        self.labelsDialog.destroy()

    def close(self, widget, other=None):
        self.result = None
        self.labelsDialog.destroy()


# Stampa etichette prezzi
def printLabels(parent, labelList, data, row, col):
    listSize = len(labelList)

    if listSize > 0:
        # Foglio A4 100 etichette 37x14 (LP4W-3714) "Tico Copy Laser Premium"
        NUM_COLS = 5
        NUM_ROWS = 20
        # Per avere le colonne centrate (anche se si usa l'A4 borderless, ci sono sempre bordi)
        COL_WIDTHS = (4.0 * cm, 4.0 * cm, 4.0 * cm, 4.0 * cm, 4.0 * cm)
        ROW_SIZE = (1.4 + 0.01) * cm
        # L'origine degli assi è l'angolo in basso a sinistra del foglio
        COL_OFFSET = 0.70 * cm
        ROW_OFFSET = 1.00 * cm

        tmpFile = tempfile.NamedTemporaryFile(delete=False)
        tmpFile.close()

        c = Canvas(tmpFile.name, pagesize=A4_BORDERLESS)
        width, height = A4_BORDERLESS
        c.setAuthor("Gestione Tabacchi")
        c.setTitle("Etichette prezzi modificati da: %s" % data.strftime("%d %B %Y"))

        tableData = [[""] * NUM_COLS for i in range(NUM_ROWS)]

        count = 0

        while (count < listSize):
            while (count < listSize) and (col < NUM_COLS):
                pezzi = labelList[count]["PezziUnitaMin"]
                if pezzi == 0:
                    msgDialog = Gtk.MessageDialog(parent, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT,
                                                  type=Gtk.MessageType.WARNING, buttons=Gtk.ButtonsType.OK, message_format="%s" % labelList[count]["Descrizione"])
                    msgDialog.format_secondary_text("Non ha il numero di pezzi per unità minima, non verrà stampata l'etichetta")
                    msgDialog.set_title("Attenzione")
                    msgDialog.run()
                    msgDialog.destroy()
                else:
                    tableData[row][col] = "€ %.2f" % ((labelList[count]["UnitaMin"] * labelList[count]["PrezzoKG"]) / pezzi)
                    col += 1
                count += 1
            row += 1
            if (row == NUM_ROWS) or (count == listSize):
                t = Table(tableData, COL_WIDTHS, ROW_SIZE)
                # GRIGLIA per debug
                # t.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('FONT',(0,0),(-1,-1), 'Helvetica-Bold', 24), ('BOX', (0,0), (-1,-1), 0.25, colors.black), ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black) ]))
                t.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONT', (0, 0), (-1, -1), 'Helvetica-Bold', 24)]))
                t.wrapOn(c, width, height)
                t.drawOn(c, COL_OFFSET, ROW_OFFSET)
                c.showPage()
                tableData = [[""] * NUM_COLS for i in range(NUM_ROWS)]
                row = 0
            col = 0
        c.save()
        Gio.Subprocess.new(["gio", "open", tmpFile.name], Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_MERGE)
    else:
        msgDialog = Gtk.MessageDialog(parent, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT, type=Gtk.MessageType.WARNING,
                                      buttons=Gtk.ButtonsType.OK, message_format="Non ci sono articoli con variazioni nel periodo indicato.")
        msgDialog.set_title("Attenzione")
        msgDialog.run()
        msgDialog.destroy()

# Stampa un ordine con una lista di articoli sul modello U88-Fax


def printU88Fax(parent, lista, tipo, data=None):
    listaSize = len(lista)

    if listaSize > 0:
        cognome = prefs.cognome
        nome = prefs.nome
        numRivendita = prefs.numRivendita
        cittaRivendita = prefs.citta
        telefono = prefs.telefono
        codCliente = prefs.codCliente

        urgente = (tipo == ordini.URGENTE)

        if not cognome or not nome or not numRivendita or not cittaRivendita or not telefono or not codCliente or not prefs.timbro or not prefs.firma or not prefs.u88:
            msgDialog = Gtk.MessageDialog(parent, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT, type=Gtk.MessageType.WARNING,
                                          buttons=Gtk.ButtonsType.OK, message_format="Non sono state impostate tutte le preferenze dei tabacchi")
            msgDialog.set_title("Attenzione")
            msgDialog.run()
            msgDialog.destroy()
            return

        for path in [prefs.timbro, prefs.firma, prefs.u88, prefs.u88urg]:
            if not os.path.isfile(path):
                msgDialog = Gtk.MessageDialog(parent, flags=Gtk.DialogFlags.MODAL & Gtk.DialogFlags.DESTROY_WITH_PARENT,
                                              type=Gtk.MessageType.WARNING, buttons=Gtk.ButtonsType.OK, message_format="Non è stato trovato il file:")
                msgDialog.format_secondary_text("%s" % path)
                msgDialog.set_title("Attenzione")
                msgDialog.run()
                msgDialog.destroy()
                return

        timbro = Image(prefs.timbro, prefs.timbroW * cm, prefs.timbroH * cm)
        firma = Image(prefs.firma, prefs.firmaW * cm, prefs.firmaH * cm)

        NUM_COLS = 2
        NUM_ROWS = 24
        CIFRA = 0.65 * cm
        ROW_SIZE = 0.815 * cm
        if urgente:
            DIV1 = 0.35 * cm
            DIV2 = 0.45 * cm
            DIV3 = 1.60 * cm
            DIV4 = 0.35 * cm
            COL_OFFSET = 2.00 * cm
            ROW_OFFSET = 4.40 * cm
        else:
            DIV1 = 0.5 * cm
            DIV2 = 0.5 * cm
            DIV3 = 2.0 * cm
            DIV4 = 0.3 * cm
            COL_OFFSET = 1.40 * cm
            ROW_OFFSET = 4.30 * cm
        COL_WIDTHS = (CIFRA, CIFRA, CIFRA, CIFRA, CIFRA, DIV1, CIFRA, CIFRA, CIFRA, DIV2, CIFRA, CIFRA, CIFRA,
                      DIV3, CIFRA, CIFRA, CIFRA, CIFRA, CIFRA, DIV4, CIFRA, CIFRA, CIFRA, DIV2, CIFRA, CIFRA, CIFRA)
        pageTot = math.ceil(listaSize / float(NUM_ROWS * NUM_COLS))
        tmpFile = tempfile.NamedTemporaryFile(delete=False)
        tmpFile.close()

        c = Canvas(tmpFile.name, pagesize=A4)
        width, height = A4

        tableData = [[""] * len(COL_WIDTHS) for i in range(NUM_ROWS)]

        count = 0
        pageNum = 1
        row = 0
        col = 0
        totale = 0
        newPage = False

        while (count < listaSize):
            cod = lista[count]["ID"].rjust(5)
            totale += lista[count]["Ordine"]
            peso = "% 7.3f" % lista[count]["Ordine"]
            tableData[row][(col * 14)] = cod[0]
            tableData[row][(col * 14) + 1] = cod[1]
            tableData[row][(col * 14) + 2] = cod[2]
            tableData[row][(col * 14) + 3] = cod[3]
            tableData[row][(col * 14) + 4] = cod[4]

            tableData[row][(col * 14) + 6] = peso[0]
            tableData[row][(col * 14) + 7] = peso[1]
            tableData[row][(col * 14) + 8] = peso[2]
            tableData[row][(col * 14) + 10] = peso[4]
            tableData[row][(col * 14) + 11] = peso[5]
            tableData[row][(col * 14) + 12] = peso[6]

            count += 1
            row += 1

            if (row >= NUM_ROWS):
                row = 0
                col += 1
                if col >= NUM_COLS:
                    col = 0
                    newPage = True
            if newPage or (count == listaSize):
                if urgente:
                    c.setFont('Helvetica-Bold', 11)
                    c.drawString(4.00 * cm, height - 2.2 * cm, codCliente)
                    c.drawString(4.00 * cm, height - 2.85 * cm, cognome)
                    c.drawString(4.00 * cm, height - 3.5 * cm, nome)
                    c.drawString(4.00 * cm, height - 4.2 * cm, numRivendita)
                    c.drawString(7.2 * cm, height - 4.2 * cm, cittaRivendita)
                    c.drawString(14.7 * cm, height - 2.2 * cm, telefono)
                    c.setFont('Helvetica', 9)
                    c.drawString(18.40 * cm, height - 3.20 * cm, "%i" % pageNum)
                    c.drawString(19.40 * cm, height - 3.20 * cm, "%1.0f" % pageTot)

                    # Timbro e firma
                    timbro.drawOn(c, 4.0 * cm, 2.0 * cm)
                    firma.drawOn(c, 13.0 * cm, 1.5 * cm)

                    if data:
                        dataStr = data.strftime("%d%m%y")
                        c.setFont('Helvetica-Bold', 14)
                        c.drawString(14.5 * cm, height - 4.2 * cm, dataStr[0])
                        c.drawString(15.2 * cm, height - 4.2 * cm, dataStr[1])
                        c.drawString(16.2 * cm, height - 4.2 * cm, dataStr[2])
                        c.drawString(16.85 * cm, height - 4.2 * cm, dataStr[3])
                        c.drawString(17.8 * cm, height - 4.2 * cm, dataStr[4])
                        c.drawString(18.5 * cm, height - 4.2 * cm, dataStr[5])
                else:
                    c.setFont('Helvetica-Bold', 11)
                    c.drawString(3.30 * cm, height - 2.2 * cm, codCliente)
                    c.drawString(3.30 * cm, height - 2.85 * cm, cognome)
                    c.drawString(3.30 * cm, height - 3.5 * cm, nome)
                    c.drawString(3.30 * cm, height - 4.20 * cm, numRivendita)
                    c.drawString(6.5 * cm, height - 4.20 * cm, cittaRivendita)
                    c.drawString(15.1 * cm, height - 2.2 * cm, telefono)
                    c.setFont('Helvetica', 9)
                    c.drawString(18.85 * cm, height - 3.20 * cm, "%i" % pageNum)
                    c.drawString(19.90 * cm, height - 3.20 * cm, "%1.0f" % pageTot)

                    c.setFont('Helvetica-Bold', 14)
                    x = 17
                    if tipo == ordini.ORDINARIO:
                        x = 14.4 * cm
                    elif tipo == ordini.STRAORDINARIO:
                        x = 15.0 * cm
                    c.drawString(x, height - 3.35 * cm, "X")

                    # Timbro e firma
                    timbro.drawOn(c, 4.0 * cm, 2.3 * cm)
                    firma.drawOn(c, 13.0 * cm, 2.0 * cm)

                # Totale
                if (count == listaSize):
                    totaleStr = "% 10.3f" % totale
                    x = 12.8 * cm
                    y = 3.6 * cm
                    c.drawString(x, y, totaleStr[0])
                    c.drawString(x + CIFRA, y, totaleStr[1])
                    c.drawString(x + CIFRA * 2, y, totaleStr[2])
                    x = 15.1 * cm
                    c.drawString(x, y, totaleStr[3])
                    c.drawString(x + CIFRA, y, totaleStr[4])
                    c.drawString(x + CIFRA * 2, y, totaleStr[5])
                    x = 17.7 * cm
                    c.drawString(x, y, totaleStr[7])
                    c.drawString(x + CIFRA, y, totaleStr[8])
                    c.drawString(x + CIFRA * 2, y, totaleStr[9])

                t = Table(tableData, COL_WIDTHS, ROW_SIZE)

                t.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONT', (0, 0), (-1, -1), 'Helvetica-Bold', 12)]))
                t.wrapOn(c, width, height)
                t.drawOn(c, COL_OFFSET, ROW_OFFSET)

                c.showPage()
                if (count < listaSize):
                    tableData = [[""] * len(COL_WIDTHS) for i in range(NUM_ROWS)]
            if newPage:
                newPage = False
                pageNum += 1

        c.save()
        thread = WatermarkThread(tmpFile.name, prefs.u88urg if urgente else prefs.u88, pageTot)
        progressDialog = utility.ProgressDialog(parent, "Generazione ordine su modello U88 Fax.", "Attendere prego..", "Generazione U88 Fax", thread)
        progressDialog.setResponseCallback(__u88FaxCallback)
        progressDialog.start()


def __u88FaxCallback(thread):
    pathname = thread.fileOutputName
    Gio.Subprocess.new(["gio", "open", pathname], Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_MERGE)

# Thread per l'applicazione del watermark


class WatermarkThread(utility.WorkerThread):
    def __init__(self, watermarkFilename, u88FaxFilename, pageTot):
        super(WatermarkThread, self).__init__()
        self.watermarkFilename = watermarkFilename
        self.u88FaxFilename = u88FaxFilename
        self.pageTot = pageTot
        self.fileOutputName = None

    def run(self):
        pageNum = 0
        fileWatermark = None
        fileU88Fax = None
        fileOutput = None
        try:
            self.progressDialog.setSteps(self.pageTot + 2)

            fileWatermark = open(self.watermarkFilename, "rb")
            fileU88Fax = open(self.u88FaxFilename, "rb")
            self.update()
            watermark = PdfFileReader(fileWatermark)
            output = PdfFileWriter()
            fileOutput = tempfile.NamedTemporaryFile(delete=False)
            self.fileOutputName = fileOutput.name
            while (pageNum < self.pageTot):
                self.update()
                # Rilegge pagina per l'ordine ogni volta.. altrimenti viene "sporcata" dal merge
                page = PdfFileReader(fileU88Fax).getPage(0)
                page.mergePage(watermark.getPage(pageNum))
                pageNum += 1
                output.addPage(page)
            output.write(fileOutput)
        except StopIteration:
            pass
        except BaseException as e:
            self.setError(e)
        else:
            self.status = self.DONE
        finally:
            if fileOutput:
                fileOutput.close()
            if fileWatermark:
                fileWatermark.close()
            if fileU88Fax:
                fileU88Fax.close()
            GLib.idle_add(self.progressDialog.close, self)

        return False
