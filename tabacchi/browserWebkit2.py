#
# Copyright (C) Francesco Guarnieri 2020 <francesco@guarnie.net>
#
import gi
from .config import log

gi.require_version('WebKit2', '4.0')
gi.require_version('Gtk', '3.0')

from gi.repository import Gtk, Gio, WebKit2, Pango  # noqa: E402

MSG, WAIT_URL, SCRIPT, ASYNC, CALLBACK = (0, 1, 2, 3, 4)
JS_ASYNC_MSG_DONE = "JS_PYMSG_CHANNEL_DONE"
JS_ASYNC_MSG_OUT_TIME = "JS_PYMSG_CHANNEL_OUT_TIME"


# Genera uno script Javascript che rimane in attesa finche una condizione si avvera, quindi esegue lo script passato
# per parametro ed infine emette un messaggio.
# Se si superano un numero massimo di iterazioni senza che la condizione si verifichi, esce lo stesso
def createAsyncScript(condition, script, max_iter=50, interval=100):
    waitScript = """var iter_num = 0; var interval = setTimeout(timeoutFunc, """ + str(interval) + """);
        function timeoutFunc() {
            iter_num ++;
            if (""" + condition + """) { clearTimeout(interval);
                """ + script + """
                window.webkit.messageHandlers.""" + JS_ASYNC_MSG_DONE + """.postMessage('""" + JS_ASYNC_MSG_DONE + """');
            }
            else if (iter_num > """ + str(max_iter) + """) { clearTimeout(interval);
                window.webkit.messageHandlers.""" + JS_ASYNC_MSG_OUT_TIME + """.postMessage('""" + JS_ASYNC_MSG_OUT_TIME + """');
            }
            else interval = setTimeout(timeoutFunc, """ + str(interval) + """);
        }
        """
    return waitScript


class Browser(Gtk.Window):
    def __init__(self, parent, userAgent=None):
        Gtk.Window.__init__(self)
        self.set_transient_for(parent)
        self.set_position(Gtk.WindowPosition.CENTER)
        self.set_default_size(1000, 600)

        self.okButton = Gtk.Button()
        icon = Gio.ThemedIcon(name="emblem-ok-symbolic")
        image = Gtk.Image.new_from_gicon(icon, Gtk.IconSize.SMALL_TOOLBAR)
        self.okButton.add(image)
        self.okButton.connect("clicked", self.__close)
        self.cancelButton = Gtk.Button()
        icon = Gio.ThemedIcon(name="window-close-symbolic")
        image = Gtk.Image.new_from_gicon(icon, Gtk.IconSize.SMALL_TOOLBAR)
        self.cancelButton.add(image)
        self.cancelButton.connect("clicked", self.__defClose)

        self.headerBar = Gtk.HeaderBar()
        self.headerBar.set_show_close_button(False)

        boxTitle = Gtk.Box(spacing=6)
        # Gtk.StyleContext.add_class(boxTitle.get_style_context(), "linked")
        self.spinner = Gtk.Spinner()
        self.labelTitle = Gtk.Label()
        labelFont = self.labelTitle.get_pango_context().get_font_description()
        labelFont.set_weight(Pango.Weight.BOLD)
        self.labelTitle.modify_font(labelFont)
        boxTitle.add(self.labelTitle)
        boxTitle.add(self.spinner)
        self.headerBar.set_custom_title(boxTitle)

        self.set_titlebar(self.headerBar)

        self.stopButton = Gtk.Button()
        icon = Gio.ThemedIcon(name="media-playback-stop-symbolic")
        image = Gtk.Image.new_from_gicon(icon, Gtk.IconSize.SMALL_TOOLBAR)
        self.stopButton.add(image)
        # self.stopButton.set_relief(Gtk.ReliefStyle.NONE)
        self.headerBar.pack_end(self.stopButton)

        box = Gtk.Box()
        Gtk.StyleContext.add_class(box.get_style_context(), "linked")
        box.add(self.okButton)
        box.add(self.cancelButton)
        self.headerBar.pack_start(box)

        browserBox = Gtk.Box()
        browserBox.set_orientation(Gtk.Orientation.VERTICAL)

        self.contentManager = WebKit2.UserContentManager()
        self.contentManager.connect(f"script-message-received::{JS_ASYNC_MSG_DONE}", self.__handleScriptMessageDone)
        if not self.contentManager.register_script_message_handler(JS_ASYNC_MSG_DONE):
            log.debug(f"Error registering script message handler: '{JS_ASYNC_MSG_DONE}'")
        self.contentManager.connect(f"script-message-received::{JS_ASYNC_MSG_OUT_TIME}", self.__handleScriptMessageOutOfTIme)
        if not self.contentManager.register_script_message_handler(JS_ASYNC_MSG_OUT_TIME):
            log.debug(f"Error registering script message handler: '{JS_ASYNC_MSG_OUT_TIME}'")

        # Inizializza Webview
        self.web_view = WebKit2.WebView.new_with_user_content_manager(self.contentManager)

        # Imposta i settings
        settings = WebKit2.Settings()
        if userAgent is not None:
            settings.set_property('user-agent', userAgent)
        settings.set_property('javascript-can-access-clipboard', True)
        self.web_view.set_settings(settings)

        browserBox.pack_start(self.web_view, True, True, 0)
        self.add(browserBox)

        self.data = None
        self.uploadFilePath = None
        self.scriptList = None
        self.closeScriptList = None
        self.stopButton.connect("clicked", self.__on_stop_click)
        self.web_view.connect("load-changed", self.__loadFinishedCallback)

    # Gestisce i messaggi DONE ricevuti da Javascript
    def __handleScriptMessageDone(self, contentManager, js_result):
        if (self.data is not None):
            if (self.data[ASYNC]):
                log.debug("JAVASCRIPT ASYNC RETURNED OK!")
                callback = self.data[CALLBACK]
                if (callback is not None):
                    callback()
                self.__popScript()

    # Gestisce i messaggi OUT OF TIME ricevuti da Javascript
    def __handleScriptMessageOutOfTIme(self, contentManager, js_result):
        if (self.data is not None):
            if (self.data[ASYNC]):
                log.debug("JAVASCRIPT ASYNC RETURNED: OUT OF TIME! No callback called..")
                self.__popScript()

    #
    def __loadFinishedCallback(self, web_element, load_event):
        if load_event == WebKit2.LoadEvent.FINISHED:
            uri = web_element.get_uri()
            log.debug("Load finished - URI: %s " % uri)
            if (self.data is not None):
                waitUrl = self.data[WAIT_URL]
                if (waitUrl is not None) and (waitUrl == uri):
                    self.__execScript()

    def __on_stop_click(self, widget):
        log.debug("STOP")
        self.__waitMode(False)
        self.data = None
        self.scriptList = None
        self.web_view.stop_loading()

    # Se sono impostati degli script di chiusura li esegue (al loro termine chiudera' il browser)
    def __close(self, widget=None):
        if self.closeScriptList is not None:
            self.scriptList = self.closeScriptList
            self.closeScriptList = None
            self.okButton.set_sensitive(False)
            self.__waitMode(True)
            self.__popScript()

    #
    def __defClose(self, widget=None):
        log.debug("CLOSING definitively..")
        self.web_view.stop_loading()
        self.destroy()

    # Eegue il codice javascript o invoca la callback
    def __execScript(self):
        script = self.data[SCRIPT]
        if script:
            log.debug(f"EXEC SCRIPT: '{script}'")
            self.web_view.run_javascript(script, None, self.__javascript_finished, None)
        else:
            self.__popScript()

    # Callback per gestire la fine dell'esecuzione del codice javascript
    # In caso di script asincrono che ritorna subito, non passa al prossimo script,
    # ma si completera' la gestione alla ricezione di un messaggio
    def __javascript_finished(self, webview, task, user_data=None):
        try:
            # js_result = webview.run_javascript_finish(task)
            webview.run_javascript_finish(task)
        except Exception as e:
            log.error("JAVASCRIPT ERROR MSG: %s" % e)
        '''
        value = WebKit2.JavascriptResult.get_js_value(js_result)
        str_value = JavaScriptCore.Value.to_string(value)
        print str_value
        '''
        if (self.data is not None):
            if (not self.data[ASYNC]):
                log.debug("JAVASCRIPT FINISHED")
                self.__popScript()
            else:
                log.debug("JAVASCRIPT RETURN DELAYED.. ASYNC")

        return False

    # Imposta la modalita di attesa se attiva o meno
    def __waitMode(self, toggle):
        self.spinner.start() if toggle else self.spinner.stop()
        self.stopButton.set_visible(toggle)
        self.stopButton.set_sensitive(toggle)
        self.web_view.set_sensitive(not toggle)

    # Estrae il successivo script da eseguire dalla lista
    def __popScript(self):
        log.debug("scriptlist: %s" % self.scriptList)
        if self.scriptList is not None:
            log.debug("len(self.scriptList)=%i" % len(self.scriptList))
            if len(self.scriptList) > 0:
                data = self.scriptList.pop()
                self.labelTitle.set_text(data[MSG])
                self.data = data
                waitUrl = self.data[WAIT_URL]
                # LO SCRIPT VIENE ESEGUITO SUBITO SENZA CONDIZIONI DI ATTESA
                if (waitUrl is None) or (waitUrl == self.web_view.get_uri()):
                    self.__execScript()
            else:
                self.data = None
                self.scriptList = None
                self.__waitMode(False)
                if self.closeScriptList is None:
                    self.__defClose()

    # Apre il sito, e se impostati, esegue gli script di callback di apertura e chiusura
    def open(self, site, scriptList=None, closeScriptList=None):
        self.scriptList = scriptList
        self.closeScriptList = closeScriptList
        if not self.get_property("visible"):
            self.show_all()
        self.__waitMode(True)
        self.__popScript()
        self.web_view.load_uri(site)
