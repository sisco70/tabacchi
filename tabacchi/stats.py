#!/usr/bin/python
# coding: UTF-8
#
# Copyright (C) Francesco Guarnieri 2020 <francesco@guarnie.net>
#

import datetime
import sqlite3

from matplotlib import dates
from matplotlib import ticker
from matplotlib.backends.backend_gtk3cairo import FigureCanvasGTK3Cairo as FigureCanvas
from matplotlib.dates import MonthLocator, DateFormatter
from matplotlib.figure import Figure
import gi

from .preferencesTabacchi import prefs
from . import utility

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk  # noqa: E402


# Classe minimale per la visualizzazione delle statistiche
class SimpleStatsDialog(utility.GladeWindow):
    def __init__(self, parent):
        super().__init__(parent, "statsDialog.glade")

        self.statsDialog = self.builder.get_object("statsDialog")
        self.statsNotebook = self.builder.get_object("statsNotebook")
        self.consumiScrolledWindow = self.builder.get_object("consumiScrolledWindow")
        self.pesiScrolledWindow = self.builder.get_object("pesiScrolledWindow")
        self.quantitaScrolledWindow = self.builder.get_object("quantitaScrolledWindow")
        self.piuAcquistatiTreeview = self.builder.get_object("piuAcquistatiTreeview")

        self.statsDialog.set_transient_for(parent)

        self.builder.connect_signals({"on_statsDialog_delete_event": self.close,
                                      "on_statsDialog_destroy_event": self.close,
                                      "on_closeButton_clicked": self.close
                                      })
        now = datetime.datetime.now().replace(microsecond=0)
        self.dataInizioEntry = utility.DataEntry(
            self.builder.get_object("dataInizioButton"),
            now - datetime.timedelta(days=365 * 2),
            " %d %B %Y ", self.refreshData)
        self.dataFineEntry = utility.DataEntry(self.builder.get_object("dataFineButton"), now, " %d %B %Y ", self.refreshData)

        self.figureCanvas = None

    def refreshData(self):
        self.loadData()
        self.showData()

    def initData(self):
        pass

    def loadData(self):
        pass

    def showData(self):
        pass

    def showPlot(self, scrolledWindow, date, valori, unitaMin=None):
        figure = Figure(figsize=None, dpi=None)
        ax = figure.add_subplot(111)

        ax.plot_date(dates.date2num(date), valori, 'o-')

        ax.xaxis.set_major_locator(MonthLocator())
        ax.xaxis.set_major_formatter(DateFormatter('%b %y'))
        if unitaMin:
            ax.yaxis.set_major_locator(ticker.MultipleLocator(1))
            ax.yaxis.set_minor_locator(ticker.MultipleLocator(unitaMin))
        figure.autofmt_xdate()
        ax.autoscale_view()
        ax.grid(True)

        figureCanvas = FigureCanvas(figure)

        oldCanvas = scrolledWindow.get_child()
        if oldCanvas:
            scrolledWindow.remove(oldCanvas)

        scrolledWindow.add_with_viewport(figureCanvas)
        figureCanvas.show()

    def run(self):
        self.result = self.statsDialog.run()
        return self.result

    def close(self, widget, other=None):
        self.result = Gtk.ResponseType.CANCEL
        self.statsDialog.destroy()


class StatsDialog(SimpleStatsDialog):
    def __init__(self, parent, idPar):
        SimpleStatsDialog.__init__(self, parent)
        self.statsNotebook.remove_page(3)
        self.id = idPar
        self.initData()
        self.loadData()

        self.statsDialog.set_title(self.descrizione)
        self.showData()

    # Inizializzazione
    def initData(self):
        self.date = []
        self.consumi = []
        self.giacenze = []
        self.ordini = []
        self.unitaMin = None
        self.descrizione = None

    # Legge i dati
    def loadData(self):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            cursor.execute("SELECT t.Descrizione, t.UnitaMin FROM tabacchi t WHERE t.ID = ?", (self.id,))
            row = cursor.fetchone()
            self.descrizione = row['Descrizione']
            self.unitaMin = row['UnitaMin']
            dataFine = self.dataFineEntry.data.strftime("%Y-%m-%d")
            dataInizio = self.dataInizioEntry.data.strftime("%Y-%m-%d")
            cursor.execute(
                "SELECT r.Consumo, r.Giacenza, r.Ordine, o.Data as 'Data [timestamp]' FROM rigaOrdineTabacchi r, ordineTabacchi o where (r.ID_Ordine = o.ID) and (r.ID = ?) and (o.Data < ?) and (o.Data > ?) order by o.Data",
                (self.id, dataFine, dataInizio))
            result_set = cursor.fetchall()
            self.date[:] = []
            self.consumi[:] = []
            self.giacenze[:] = []
            self.ordini[:] = []
            oldOrdine = 0
            for row in result_set:
                data = row['Data']
                consumo = round(row['Consumo'], 3)
                ordine = round(row['Ordine'], 3)
                giacenza = round(row['Giacenza'], 3) - oldOrdine
                self.date.append(data)
                self.consumi.append(consumo if (consumo > 0) else 0)
                self.ordini.append(ordine if (ordine > 0) else 0)
                self.giacenze.append(giacenza if (giacenza > 0) else 0)
                oldOrdine = ordine
        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self.statsDialog)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def showData(self):
        self.showPlot(self.consumiScrolledWindow, self.date, self.consumi, self.unitaMin)
        self.showPlot(self.pesiScrolledWindow, self.date, self.ordini, self.unitaMin)
        self.showPlot(self.quantitaScrolledWindow, self.date, self.giacenze, self.unitaMin)


class GlobalStatsDialog(SimpleStatsDialog):
    modelInfoList = [("+Descrizione", "str"), ("^Totale acquisti", "float"), ("^Data ultimo acquisto", "date")]

    def __init__(self, parent):
        SimpleStatsDialog.__init__(self, parent)

        utility.ExtTreeView(self.modelInfoList, self.piuAcquistatiTreeview, formats={1: "%.3f kg"})
        self.piuAcquistatiModel = self.piuAcquistatiTreeview.get_model()
        self.piuAcquistatiTreeview.set_model(None)
        self.initData()
        self.loadData()
        self.piuAcquistatiTreeview.set_model(self.piuAcquistatiModel)

        self.statsDialog.set_title("Statistiche generali")

        self.showData()

    # Inizializzazione
    def initData(self):
        self.date = []
        self.consumi = []
        self.giacenze = []
        self.ordini = []

    # Legge i dati

    def loadData(self):
        cursor = None
        conn = None
        try:
            conn = prefs.getConn()
            cursor = prefs.getCursor(conn)
            dataFine = self.dataFineEntry.data.strftime("%Y-%m-%d")
            dataInizio = self.dataInizioEntry.data.strftime("%Y-%m-%d")
            cursor.execute(
                "SELECT t.Descrizione, SUM(r.Ordine) totPeso, MAX(o.Data) as 'maxData [timestamp]' FROM rigaOrdineTabacchi r, ordineTabacchi o, tabacchi t where (r.ID = t.ID) and t.InMagazzino and (r.ID_Ordine = o.ID) and (r.Ordine > 0) and (o.Data < ?) and (o.Data > ?) group by r.ID order by totPeso desc",
                (dataFine, dataInizio))
            result_set = cursor.fetchall()
            self.piuAcquistatiModel.clear()
            for row in result_set:
                self.piuAcquistatiModel.append([row["Descrizione"], row["totPeso"], row["maxData"]])
            cursor.execute(
                "SELECT SUM(r.Consumo) as Consumo, SUM(r.Giacenza) as Giacenza, SUM(r.Ordine) as Ordine, o.Data as 'Data [timestamp]' FROM rigaOrdineTabacchi r, ordineTabacchi o where (r.ID_Ordine = o.ID) and ((r.Consumo > 0) or (r.Giacenza > 0) or (r.Ordine > 0)) and (o.Data < ?) and (o.Data > ?) group by o.ID order by o.Data",
                (dataFine, dataInizio))
            result_set = cursor.fetchall()

            self.date[:] = []
            self.consumi[:] = []
            self.giacenze[:] = []
            self.ordini[:] = []
            # oldOrdine = 0
            for row in result_set:
                data = row['Data']
                consumo = round(row['Consumo'], 3)
                ordine = round(row['Ordine'], 3)
                esistenza = round(row['Giacenza'], 3)
                giacenza = esistenza  # - oldOrdine

                self.date.append(data)
                self.consumi.append(consumo if (consumo > 0) else 0)
                self.ordini.append(ordine if (ordine > 0) else 0)
                self.giacenze.append(giacenza if (giacenza > 0) else 0)
                # oldOrdine = ordine
        except sqlite3.Error as e:
            utility.gtkErrorMsg(e, self.statsDialog)
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def showData(self):
        self.showPlot(self.consumiScrolledWindow, self.date, self.consumi)
        self.showPlot(self.pesiScrolledWindow, self.date, self.ordini)
        self.showPlot(self.quantitaScrolledWindow, self.date, self.giacenze)
