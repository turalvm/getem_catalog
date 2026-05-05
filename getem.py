import os
import sys
import wx
import pandas as pd
import webbrowser
import ctypes
import winsound
import urllib.parse
import re
import threading
import json
import time
import shutil
import subprocess
import zipfile
import urllib.request
from io import StringIO
from cryptography.fernet import Fernet
import requests
import functools
import pickle

# ============ SIFRELEME ANAHTARI ============
def get_key():
    key_file = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'GETEM', 'secret.key')
    if os.path.exists(key_file):
        with open(key_file, 'rb') as f:
            return f.read()
    else:
        key = Fernet.generate_key()
        os.makedirs(os.path.dirname(key_file), exist_ok=True)
        with open(key_file, 'wb') as f:
            f.write(key)
        return key

cipher = Fernet(get_key())
# ============================================

# ============ GOMULU VERSIYON ============
def get_bundled_version():
    try:
        path = resource_path("version.txt")
        if os.path.exists(path):
            with open(path, "r") as f:
                return f.read().strip()
    except:
        pass
    return "0.0"
# ============================================

# ============ CHROME DRIVER ============
def get_chrome_driver_path(lang_dict):
    appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
    driver_dir = os.path.join(appdata, 'GETEM', 'driver')
    driver_path = os.path.join(driver_dir, 'chromedriver.exe')
    if os.path.exists(driver_path):
        return driver_path
    if not os.path.exists(driver_dir):
        os.makedirs(driver_dir)
    try:
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon', '/v', 'version'],
                capture_output=True, text=True
            )
            full_version = result.stdout.split()[-1]
            main_version = full_version.split('.')[0]
        except:
            main_version = '134'
        import json
        api_url = "https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build-with-downloads.json"
        with urllib.request.urlopen(api_url) as resp:
            data = json.loads(resp.read())
        driver_url = None
        for build, info in data['builds'].items():
            if build.startswith(main_version):
                for item in info['downloads']['chromedriver']:
                    if item['platform'] == 'win64':
                        driver_url = item['url']
                        break
                if driver_url:
                    break
        if not driver_url:
            driver_url = "https://storage.googleapis.com/chrome-for-testing-public/134.0.6998.118/win64/chromedriver-win64.zip"
        zip_path = os.path.join(driver_dir, 'chromedriver.zip')
        urllib.request.urlretrieve(driver_url, zip_path)
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(driver_dir)
        for root, dirs, files in os.walk(driver_dir):
            if 'chromedriver.exe' in files:
                shutil.move(os.path.join(root, 'chromedriver.exe'), driver_path)
                break
        os.remove(zip_path)
        for item in os.listdir(driver_dir):
            item_path = os.path.join(driver_dir, item)
            if os.path.isdir(item_path):
                shutil.rmtree(item_path)
        return driver_path
    except Exception as e:
        error_msg = lang_dict.get("driver_download_error", "ChromeDriver download error: {}\nPlease check your internet connection.").format(str(e))
        wx.MessageBox(error_msg, lang_dict.get("error_title", "Error"), wx.OK | wx.ICON_ERROR)
        return None
# ============================================

if getattr(sys, 'frozen', False):
    sys.stdout = open(os.devnull, 'w')
    sys.stderr = open(os.devnull, 'w')

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

_dll = None
def nvda_speak(text):
    global _dll
    if _dll is None:
        for name in ["nvdaControllerClient64.dll", "nvdaControllerClient.dll"]:
            path = resource_path(name)
            if os.path.isfile(path):
                try:
                    _dll = ctypes.WinDLL(path)
                    break
                except:
                    continue
    if _dll:
        try:
            _dll.nvdaController_speakText(ctypes.c_wchar_p(str(text)))
        except:
            pass


SOUND_ENABLED = True
def play_app_sound(sound_name, loop=False):
    if not SOUND_ENABLED:
        return
    sound_path = resource_path(os.path.join("sound", sound_name))
    if os.path.exists(sound_path):
        flags = winsound.SND_ASYNC
        if loop:
            flags |= winsound.SND_LOOP
        winsound.PlaySound(sound_path, flags)

def stop_all_sounds():
    winsound.PlaySound(None, winsound.SND_PURGE)

download_stop_flag = False

# ============ FALLBACK INGILIZCE (KISA) ============
FALLBACK_LANG = {
    "app_title": "GETEM Catalog",
    "cat": "Category",
    "genre": "Genre",
    "search": "Search",
    "open": "Download ZIP (Ctrl+E)",
    "stop": "Stop (Esc)",
    "col_title": "Work Title",
    "col_extra": "Genre / Format",
    "cats": ["Books", "Movies", "Favorites"],
    "all_genres": "All",
    "fav_empty": "Favorite list is empty",
    "not_found": "No results found",
    "items_found": "{} results",
    "no_results": "No results",
    "fav_suffix": " (Favorite)",
    "speech_listing_books": "Listing books",
    "speech_listing_movies": "Listing movies",
    "speech_listing_favs": "Listing favorites",
    "sorted_default": "Sorted by original order",
    "update": "Update catalog",
    "exit": "Exit",
    "menu_file": "File",
    "menu_logout": "Logout",
    "menu_account": "Account",
    "lang": "Language",
    "menu_download_dir": "Download folder",
    "default_start": "Startup view",
    "view_books": "Books",
    "view_movies": "Movies",
    "view_favs": "Favorites",
    "disable_sound": "Disable sound",
    "clear_data": "Clear all data",
    "menu_options_label": "Settings",
    "about": "About",
    "help_doc": "Help (F1)",
    "menu_check_updates": "Check for updates",
    "menu_help": "Help",
    "about_text": "GETEM Catalog - Version: {version}",
    "old_file_msg": "Catalog files not found.",
    "old_file_title": "Catalog missing",
    "updating": "Updating...",
    "reading": "Reading...",
    "downloading_catalog": "Downloading: {}",
    "catalog_downloaded": "Downloaded: {}",
    "login_title": "Login",
    "login_user_label": "User:",
    "login_pass_label": "Pass:",
    "btn_save": "Save",
    "btn_cancel": "Cancel",
    "login_speech": "Login data",
    "account_saved": "Credentials saved",
    "msg_logged_out": "Logged out",
    "busy_msg": "Another operation in progress",
    "save_dir_prompt": "Save to:",
    "conn_error": "Link not found",
    "msg_browser_start": "Starting...",
    "status_login": "Logging in...",
    "login_failed": "Login failed",
    "msg_btn_not_found": "Download button not found",
    "msg_wait_seconds": "Please wait a few seconds",
    "download_denied": "Access denied",
    "download_started": "Downloading: {}",
    "download_complete": "Downloaded: {}",
    "status_stopping": "Stopping...",
    "status_stopped": "Stopped",
    "driver_not_found": "ChromeDriver not found",
    "driver_download_error": "ChromeDriver download error: {}\nPlease check your internet connection.",
    "error_title": "Error",
    "get_details": "Work Description",
    "no_details": "No detail information found.",
    "details_fetching": "Fetching details...",
    "details_shown": "Details shown",
    "select_eser_first": "Please select a work first",
    "sort_by": "Sort by",
    "sort_default": "Original",
    "sort_title_asc": "Title A-Z",
    "sort_title_desc": "Title Z-A",
    "sort_format_asc": "Format A-Z",
    "sort_format_desc": "Format Z-A",
    "sort_view_asc": "Views low-high",
    "sort_view_desc": "Views high-low",
    "sort_size_asc": "Size small-large",
    "sort_size_desc": "Size large-small",
    "sort_track_asc": "Tracks low-high",
    "sort_track_desc": "Tracks high-low",
    "ctx_fav": "Favorite (Ctrl+D)",
    "fav_added": "Added to favorites",
    "fav_removed": "Removed from favorites",
    "opening_browser": "Opening browser",
    "opening_help": "Opening help...",
    "help_not_found": "Help file not found!",
    "clear_data_confirm": "Delete all user data?",
    "clear_data_title": "Clear data",
    "update_check": "Checking for updates...",
    "update_question": "New version available!\nYour: {}\nLatest: {}\nDownload now?",
    "update_latest": "You are using the latest version.",
    "catalog_old_msg": "Catalog files older than 1 week.\nDo you want to update now?",
    "catalog_old_title": "Catalog outdated",
    "catalog_missing_msg": "Catalog files not found.\nDo you want to download now?",
    "catalog_missing_title": "Catalog missing",
    "list_item_format": "{title}, Yazar: {author}, Tür: {genre}, Format: {format}",
    "catalog_books": "books",
    "catalog_movies": "movies",
    "update_downloading": "Downloading update...",
    "update_downloaded": "Update downloaded, installing...",
    "update_complete": "Download complete, one hundred percent",
    "update_installing": "Installing update, please wait",
    "update_error": "Update error",
    "update_manual": "Download manually"
}
# ============================================

class LanguageDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="Dil Seçimi", size=(300, 150))
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        st = wx.StaticText(panel, label="Lütfen bir dil seçin:")
        vbox.Add(st, 0, wx.ALL, 10)
        self.lang_choice = wx.Choice(panel)
        self.lang_choice.AppendItems(["English", "Türkçe", "Azərbaycanca"])
        self.lang_choice.SetSelection(0)
        vbox.Add(self.lang_choice, 0, wx.EXPAND | wx.ALL, 10)
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        ok_btn = wx.Button(panel, label="Tamam", id=wx.ID_OK)
        btn_sizer.Add(ok_btn, 1, wx.ALL, 5)
        vbox.Add(btn_sizer, 0, wx.ALIGN_CENTER)
        panel.SetSizer(vbox)
        self.Centre()
        self.lang_choice.Bind(wx.EVT_CHAR_HOOK, self.on_choice_char)

    def on_choice_char(self, event):
        if event.GetKeyCode() == wx.WXK_RETURN:
            self.EndModal(wx.ID_OK)
        else:
            event.Skip()

class GETEMLoginDialog(wx.Dialog):
    def __init__(self, parent, L):
        super().__init__(parent, title=L.get("login_title", "Login"), size=(350, 250))
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        wx.StaticText(panel, label=L.get("login_user_label", "User:")).Wrap(300)
        self.user_txt = wx.TextCtrl(panel)
        vbox.Add(self.user_txt, 0, wx.EXPAND | wx.ALL, 10)
        wx.StaticText(panel, label=L.get("login_pass_label", "Pass:")).Wrap(300)
        self.pass_txt = wx.TextCtrl(panel, style=wx.TE_PASSWORD)
        vbox.Add(self.pass_txt, 0, wx.EXPAND | wx.ALL, 10)
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        save_btn = wx.Button(panel, label=L.get("btn_save", "Save"), id=wx.ID_OK)
        cancel_btn = wx.Button(panel, label=L.get("btn_cancel", "Cancel"), id=wx.ID_CANCEL)
        btn_sizer.Add(save_btn, 1, wx.ALL, 5)
        btn_sizer.Add(cancel_btn, 1, wx.ALL, 5)
        vbox.Add(btn_sizer, 0, wx.ALIGN_CENTER)
        panel.SetSizer(vbox)
        self.Centre()
        nvda_speak(L.get("login_speech", "Login data"))

class VirtualList(wx.ListCtrl):
    def __init__(self, parent):
        super().__init__(parent, style=wx.LC_REPORT | wx.LC_VIRTUAL | wx.LC_SINGLE_SEL)
        self.InsertColumn(0, "Eser Adı", width=400)
        self.InsertColumn(1, "Tür / Format", width=300)
        self.display_data = []
        self.SetItemCount(0)

    def apply_language(self, col_title, col_extra):
        item = wx.ListItem()
        item.SetText(col_title)
        self.SetColumn(0, item)
        item2 = wx.ListItem()
        item2.SetText(col_extra)
        self.SetColumn(1, item2)

    def OnGetItemText(self, item, col):
        if item < len(self.display_data):
            if col == 0:
                return self.display_data[item][0]
            else:
                return self.display_data[item][1]
        return ""

    def update_list(self, data_list, auto_focus=False):
        self.display_data = data_list
        self.SetItemCount(len(data_list))
        self.Refresh()
        if len(data_list) > 0:
            self.Select(0)
            self.Focus(0)
            if auto_focus:
                wx.CallAfter(self.SetFocus)
        else:
            play_app_sound("empty.wav")

class GetemApp(wx.Frame):
    def __init__(self):
        current_version = get_bundled_version()
        super().__init__(parent=None, title=f'GETEM Catalog Pro v{current_version}', size=(1000, 900))


        appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
        self.base_dir = os.path.join(appdata, 'GETEM')
        self.cache_dir = os.path.join(self.base_dir, 'data')
        self.profile_dir = os.path.join(self.base_dir, 'chrome_profile')
        self.driver_dir = os.path.join(self.base_dir, 'driver')
        for d in [self.base_dir, self.cache_dir, self.profile_dir, self.driver_dir]:
            if not os.path.exists(d):
                os.makedirs(d)

        self.lang_dir = resource_path("lang")
        self.fav_file = os.path.join(self.cache_dir, "favorites.json")
        self.settings_file = os.path.join(self.cache_dir, "settings.enc")

        self.settings = self.load_settings()
        global SOUND_ENABLED
        SOUND_ENABLED = self.settings.get("sound_enabled", True)

        if not os.path.exists(self.settings_file) or not self.settings.get("lang"):
            dlg = LanguageDialog(self)
            if dlg.ShowModal() == wx.ID_OK:
                selected = dlg.lang_choice.GetSelection()
                lang_map = {0: "EN", 1: "TR", 2: "AZ"}
                self.settings["lang"] = lang_map.get(selected, "EN")
                self.save_settings()
            else:
                self.settings["lang"] = "EN"
            dlg.Destroy()

        self.current_lang = self.settings.get("lang", "EN")
        self.L = self.load_language_files()
        self.favorites = self.load_favorites()
        self.is_logged_in = bool(self.settings.get("getem_user") and self.settings.get("getem_pass"))

        self.memory_cache = {"books": pd.DataFrame(), "movies": pd.DataFrame()}
        self.current_col = ""
        self.author_col = ""
        self.genre_col = ""
        self.format_col = ""
        self.view_col = ""
        self.size_col = ""
        self.track_col = ""
        self.is_busy = False
        self.has_focused = False
        self.status_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_status_timer, self.status_timer)

        self.column_cache = {}
        self.current_data_key = None

        default_view = self.settings.get("default_view", "books")
        view_index = {"books": 0, "movies": 1, "favs": 2}.get(default_view, 0)

        self.DATA_INFO = {
            "movies": {"url": "https://getem.boun.edu.tr/excelyap.php?q=katalog&title=&field_yazar_value=&field_seslendiren_value=&body_value=&field_yayinevi_value=&field_formati_value=&field_dil_value=&type=sesli_betimleme&field_alt_tur_kitap_value=&field_alindigikurum_value=&field_kategorisi_value=3&items_per_page=500",
                       "file": os.path.join(self.cache_dir, "movies.xls"),
                       "pkl": os.path.join(self.cache_dir, "movies.pkl")},
            "books": {"url": "https://getem.boun.edu.tr/excelyap.php?q=katalog&title=&field_yazar_value=&field_seslendiren_value=&body_value=&field_yayinevi_value=&field_formati_value=&field_dil_value=&type=kitap&field_alt_tur_kitap_value=&field_alindigikurum_value=&field_kategorisi_value=4&items_per_page=500",
                      "file": os.path.join(self.cache_dir, "books.xls"),
                      "pkl": os.path.join(self.cache_dir, "books.pkl")}
        }

        self.panel = wx.Panel(self)
        self.all_data = pd.DataFrame()
        self.filtered_df = pd.DataFrame()

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        top_row = wx.BoxSizer(wx.HORIZONTAL)
        self.cat_st = wx.StaticText(self.panel, label="")
        top_row.Add(self.cat_st, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        self.mode_choice = wx.Choice(self.panel, size=(150, -1))
        self.mode_choice.SetToolTip("Kategori seçin (Kitap, Film, Favori)")
        top_row.Add(self.mode_choice, 0, wx.ALL, 5)
        top_row.AddStretchSpacer()
        main_sizer.Add(top_row, 0, wx.EXPAND)

        search_row = wx.BoxSizer(wx.HORIZONTAL)
        self.search_st = wx.StaticText(self.panel, label="")
        search_row.Add(self.search_st, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        self.search = wx.TextCtrl(self.panel, style=wx.TE_PROCESS_ENTER, size=(400, -1))
        self.search.SetToolTip("Eser adı veya yazar adı ile ara")
        search_row.Add(self.search, 1, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(search_row, 0, wx.EXPAND)

        genre_row = wx.BoxSizer(wx.HORIZONTAL)
        self.genre_st = wx.StaticText(self.panel, label="")
        genre_row.Add(self.genre_st, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        self.genre_choice = wx.Choice(self.panel)
        self.genre_choice.SetToolTip("Tür / kategori seçin")
        genre_row.Add(self.genre_choice, 1, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(genre_row, 0, wx.EXPAND)

        self.status = wx.StaticText(self.panel, label="")
        main_sizer.Add(self.status, 0, wx.ALL, 10)
        self.progress_bar = wx.Gauge(self.panel, range=100, style=wx.GA_HORIZONTAL)
        self.progress_bar.Hide()
        main_sizer.Add(self.progress_bar, 0, wx.EXPAND | wx.ALL, 5)

        self.list = VirtualList(self.panel)
        self.list.SetToolTip("Eser listesi. Üzerine sağ tıklayın")
        main_sizer.Add(self.list, 1, wx.EXPAND | wx.ALL, 5)

        self.summary_text = wx.TextCtrl(self.panel, style=wx.TE_READONLY | wx.TE_RICH2 | wx.TE_MULTILINE, size=(-1, 150))
        self.summary_text.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        self.summary_text.Hide()
        main_sizer.Add(self.summary_text, 0, wx.EXPAND | wx.ALL, 5)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_open = wx.Button(self.panel, label="", size=(150, 35))
        self.btn_open.SetToolTip("ZIP dosyasını indir (Ctrl+E)")
        btn_row.Add(self.btn_open, 1, wx.RIGHT, 5)
        self.btn_stop = wx.Button(self.panel, label="", size=(150, 35))
        self.btn_stop.SetToolTip("İndirmeyi durdur (Esc)")
        self.btn_stop.Enable(False)
        btn_row.Add(self.btn_stop, 1, wx.LEFT, 5)
        main_sizer.Add(btn_row, 0, wx.EXPAND | wx.ALL, 10)

        self.panel.SetSizer(main_sizer)
        self.apply_language()
        self.mode_choice.SetSelection(view_index)
        self.mode_choice.Bind(wx.EVT_CHOICE, self.on_mode_change)
        self.genre_choice.Bind(wx.EVT_CHOICE, self.on_genre_filter)
        self.search.Bind(wx.EVT_TEXT, self.on_search)
        self.btn_open.Bind(wx.EVT_BUTTON, self.on_btn_open)
        self.btn_stop.Bind(wx.EVT_BUTTON, self.on_stop_download)
        self.list.Bind(wx.EVT_LIST_ITEM_SELECTED, self.on_select)
        self.list.Bind(wx.EVT_CONTEXT_MENU, self.on_right_click)
        self.list.Bind(wx.EVT_CHAR_HOOK, self.on_key)
        self.Bind(wx.EVT_CHAR_HOOK, self.on_global_key)

        self.check_for_app_update(manual=False, trigger_catalog_after=True)
        self.Show()

    def on_status_timer(self, event):
        self.status.SetLabel("")
        self.status_timer.Stop()

    def safe_status_update(self, text):
        wx.CallAfter(self._do_status, text)
        wx.SafeYield()

    def _do_status(self, text):
        self.status.SetLabel(text)
        nvda_speak(text)
        if self.status_timer.IsRunning():
            self.status_timer.Stop()
        self.status_timer.Start(3000, oneShot=True)

    def _announce_mode(self):
        mode = self.get_active_internal_key()
        if mode == "books":
            msg = self.L.get("speech_listing_books", "Listing books")
        elif mode == "movies":
            msg = self.L.get("speech_listing_movies", "Listing movies")
        else:
            msg = self.L.get("speech_listing_favs", "Listing favorites")
        nvda_speak(msg)

    def _set_focus_to_list(self):
        if self.list.GetItemCount() > 0:
            self.list.SetFocus()
            self.list.Select(0)
            self.has_focused = True

    def sort_by(self, field, ascending=True):
        if self.filtered_df.empty:
            return
        if field == 'default':
            self.filtered_df = self.all_data.copy()
            msg_key = "sorted_default"
        else:
            col_map = {
                'title': self.current_col,
                'format': self.format_col,
                'view': self.view_col,
                'size': self.size_col,
                'track': self.track_col
            }
            col = col_map.get(field)
            if not col:
                return
            df = self.filtered_df.copy()
            if field in ('view', 'track'):
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            elif field == 'size':
                df[col] = df[col].astype(str).str.extract(r'(\d+[,.]?\d*)')[0]
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            self.filtered_df = df.sort_values(by=col, ascending=ascending, na_position='last')
            msg_key = f"sorted_{field}_{'asc' if ascending else 'desc'}"
        self.update_list_from_filtered()
        self.safe_status_update(self.L.get(msg_key, "Sorted"))

    def update_list_from_filtered(self):
        if self.filtered_df.empty:
            self.list.update_list([], False)
            return

        df = self.filtered_df
        title = df[self.current_col].fillna('').astype(str)

        yazar_label = self.L.get("author_label", "Yazar:")
        tur_label = self.L.get("genre_label", "Tür:")
        format_label = self.L.get("format_label", "Format:")
    
        if self.author_col and self.author_col in df.columns:
            author = df[self.author_col].fillna('').astype(str)
            author_part = author.str.strip()
            author_part = author_part.where(author_part != '', '')
            author_part = author_part.apply(lambda x: f", {yazar_label} {x}" if x else '')
        else:
            author_part = [''] * len(df)

        if self.genre_col and self.genre_col in df.columns:
            genre = df[self.genre_col].fillna('').astype(str)
            genre_part = genre.str.strip()
            genre_part = genre_part.where(genre_part != '', '')
            genre_part = genre_part.apply(lambda x: f", {tur_label} {x}" if x else '')
        else:
            genre_part = [''] * len(df)

        if self.format_col and self.format_col in df.columns:
            frmt = df[self.format_col].fillna('').astype(str)
            format_part = frmt.str.strip()
            format_part = format_part.where(format_part != '', '')
            format_part = format_part.apply(lambda x: f", {format_label} {x}" if x else '')
        else:
            format_part = [''] * len(df)

        combined = title + author_part + genre_part + format_part
        combined = combined.str.replace(r'^,\s*', '', regex=True)
        self.list.update_list([(txt, "") for txt in combined.tolist()], False)
        if len(self.filtered_df) > 0:
            self.list.Select(0)
            self.list.Focus(0)

    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "rb") as f:
                    enc = f.read()
                    dec = cipher.decrypt(enc)
                    return json.loads(dec.decode())
            except:
                return {}
        return {"lang": "EN", "sound_enabled": True}

    def save_settings(self):
        data = json.dumps(self.settings).encode()
        enc = cipher.encrypt(data)
        with open(self.settings_file, "wb") as f:
            f.write(enc)

    def check_for_app_update(self, manual=False, trigger_catalog_after=False):
        threading.Thread(target=self._check_update_thread, args=(manual, trigger_catalog_after), daemon=True).start()



    def _check_update_thread(self, manual=False, trigger_catalog_after=False):
        try:
            url = "https://raw.githubusercontent.com/turalvm/getem_catalog/main/version.json"
            r = requests.get(url, timeout=5)
            if r.status_code != 200:
                if manual:
                    wx.CallAfter(wx.MessageBox, f"Update check failed (HTTP {r.status_code})", self.L.get("error_title", "Error"), wx.OK | wx.ICON_ERROR)
                if trigger_catalog_after:
                    wx.CallAfter(self.check_auto_update)
                return
            data = r.json()
            latest = data.get("version", "")
            durl = data.get("download_url", "")
            cur = get_bundled_version()
            if latest and durl and latest > cur:
                msg = self.L.get("update_question", "New version available!\nYour: {}\nLatest: {}\nDownload now?").format(cur, latest)
                answer = wx.MessageBox(msg, self.L.get("update", "Update"), wx.YES_NO | wx.ICON_QUESTION)
                if answer == wx.YES:
                    try:
                        wx.CallAfter(play_app_sound, "wait.wav", True)
                        wx.CallAfter(self.safe_status_update, self.L.get("update_downloading", "Güncelleme indiriliyor..."))
                        
                        current_exe = sys.executable if getattr(sys, 'frozen', False) else None
                        
                        # Kullanıcının çalıştırdığı EXE'nin adını al ve koru
                        if current_exe:
                            exe_name = os.path.basename(current_exe)
                        else:
                            exe_name = "getem_catalog.exe"
                        
                        temp_path = os.path.join(os.environ.get('TEMP', 'C:\\temp'), exe_name)
                        
                        wx.CallAfter(self.progress_bar.Show)
                        wx.CallAfter(self.progress_bar.SetValue, 0)
                        wx.CallAfter(self.Layout)
                        
                        req = urllib.request.Request(durl, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=60) as response:
                            total_size = int(response.headers.get('content-length', 0))
                            downloaded = 0
                            last_percent = -1
                            last_nvda_percent = -1
                            
                            with open(temp_path, 'wb') as out_file:
                                while True:
                                    chunk = response.read(8192)
                                    if not chunk:
                                        break
                                    out_file.write(chunk)
                                    downloaded += len(chunk)
                                    
                                    if total_size > 0:
                                        percent = int(downloaded * 100 / total_size)
                                        if percent != last_percent:
                                            last_percent = percent
                                            wx.CallAfter(self.progress_bar.SetValue, percent)
                                            
                                            nvda_percent = (percent // 10) * 10
                                            if nvda_percent != last_nvda_percent and nvda_percent > 0:
                                                last_nvda_percent = nvda_percent
                                                if nvda_percent == 100:
                                                    nvda_speak(self.L.get("update_complete", "İndirme tamamlandı, yüzde yüz"))
                                                else:
                                                    nvda_speak(f"{self.L.get('update_downloading', 'İndiriliyor')} {nvda_percent}")
                        
                        wx.CallAfter(stop_all_sounds)
                        wx.CallAfter(play_app_sound, "v8ok.wav")
                        wx.CallAfter(self.safe_status_update, self.L.get("update_downloaded", "Güncelleme indirildi, kuruluyor..."))
                        wx.CallAfter(self.progress_bar.SetValue, 100)
                        nvda_speak(self.L.get("update_installing", "Güncelleme kuruluyor, lütfen bekleyin"))
                        
                        if current_exe and os.path.exists(current_exe):
                            # Kullanıcının EXE adıyla kopyala
                            exe_dir = os.path.dirname(current_exe)
                            new_exe_path = os.path.join(exe_dir, exe_name)
                            
                            batch_path = os.path.join(os.environ.get('TEMP', 'C:\\temp'), 'update.bat')
                            with open(batch_path, 'w') as f:
                                f.write(f'''@echo off
:loop
timeout /t 1 /nobreak > nul
tasklist | find "{os.path.basename(current_exe)}" > nul
if not errorlevel 1 goto loop
timeout /t 2 /nobreak > nul
del /f /q "{current_exe}"
copy /y "{temp_path}" "{new_exe_path}"
start "" "{new_exe_path}"
del /f /q "{temp_path}"
del /f /q "{batch_path}"
''')
                            subprocess.Popen([batch_path], creationflags=subprocess.CREATE_NO_WINDOW)
                            wx.CallAfter(self.progress_bar.Hide)
                            wx.CallAfter(self.Close)
                    except Exception as e:
                        wx.CallAfter(stop_all_sounds)
                        wx.CallAfter(play_app_sound, "v8nome.wav")
                        wx.CallAfter(self.progress_bar.Hide)
                        wx.CallAfter(wx.MessageBox, f"{self.L.get('update_error', 'Güncelleme hatası')}: {str(e)}\n{self.L.get('update_manual', 'Manuel olarak indirin')}: {durl}", self.L.get("error_title", "Error"), wx.OK | wx.ICON_ERROR)
                        webbrowser.open(durl)
                if trigger_catalog_after:
                    wx.CallAfter(self.check_auto_update)
            else:
                if manual:
                    wx.CallAfter(wx.MessageBox, self.L.get("update_latest", "You are using the latest version."), self.L.get("update_check", "Check for updates"), wx.OK | wx.ICON_INFORMATION)
                if trigger_catalog_after:
                    wx.CallAfter(self.check_auto_update)
        except Exception as e:
            if manual:
                wx.CallAfter(wx.MessageBox, f"Update check failed: {str(e)}", self.L.get("error_title", "Error"), wx.OK | wx.ICON_ERROR)
            if trigger_catalog_after:
                wx.CallAfter(self.check_auto_update)




    def load_language_files(self):
        lang_code = self.current_lang.lower()
        json_path = os.path.join(self.lang_dir, f"{lang_code}.json")
        lang_dict = FALLBACK_LANG.copy()
        if os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    user_lang = json.load(f)
                    lang_dict.update(user_lang)
            except:
                pass
        return lang_dict

    def apply_language(self):
        L = self.L
        current_version = get_bundled_version()

        self.SetTitle(f"{L.get('app_title', 'GETEM Catalog ')} v{current_version}")
        self.cat_st.SetLabel(L.get("cat", "Category"))
        self.genre_st.SetLabel(L.get("genre", "Genre"))
        self.search_st.SetLabel(L.get("search", "Search"))
        self.btn_open.SetLabel(L.get("open", "Download ZIP (Ctrl+E)"))
        self.btn_stop.SetLabel(L.get("stop", "Stop (Esc)"))
        self.status.SetLabel("")
        col1 = L.get("col_title", "Work Title")
        col2 = L.get("col_extra", "Genre / Format")
        self.list.apply_language(col1, col2)
        cats = L.get("cats", ["Books", "Movies", "Favorites"])
        self.mode_choice.SetItems(cats)
        self.create_menu()

    def get_active_internal_key(self):
        return ["books", "movies", "favs"][self.mode_choice.GetSelection()]

    def update_genres_list(self):
        self.genre_choice.Clear()
        genres = [self.L.get("all_genres", "All")]
        if not self.all_data.empty and self.genre_col and self.genre_col in self.all_data.columns:
            all_genres = set()
            for g in self.all_data[self.genre_col].dropna():
                if isinstance(g, str):
                    for p in g.split(','):
                        p = p.strip()
                        if p and p.lower() != 'nan':
                            all_genres.add(p)
            genres.extend(sorted(all_genres))
        self.genre_choice.AppendItems(genres)
        self.genre_choice.SetSelection(0)

    def identify_columns(self):
        self.current_col = self.author_col = self.genre_col = self.format_col = self.view_col = self.size_col = self.track_col = ""
        for c in self.all_data.columns:
            cn = str(c).lower()
            if ("eser" in cn or "ad" in cn or "title" in cn) and "id" not in cn and "link" not in cn:
                if not self.current_col:
                    self.current_col = c
            if "yazar" in cn or "müellif" in cn or "author" in cn:
                self.author_col = c
            if any(x in cn for x in ["kategori", "tür", "konu", "janr", "genre"]):
                self.genre_col = c
            if any(x in cn for x in ["format", "biçim", "tip", "type", "eser formatı", "eser formati"]):
                self.format_col = c
            if any(x in cn for x in ["görüntülenme", "goruntulenme", "view", "hit", "tıklanma", "okunma", "izlenme"]):
                self.view_col = c
            if any(x in cn for x in ["eser boyutu", "boyut", "mb", "kb", "gb", "dosya boyutu", "file size"]):
                self.size_col = c
            if any(x in cn for x in ["parça sayısı", "parça", "track", "ayrım", "bölüm"]):
                self.track_col = c
        if not self.current_col and not self.all_data.empty:
            self.current_col = self.all_data.columns[0]

    def refresh_ui(self, auto_focus=False):
        if self.all_data.empty:
            self.filtered_df = pd.DataFrame()
            self.list.update_list([], auto_focus)
            msg = self.L.get("fav_empty", "Favorite list is empty") if self.get_active_internal_key() == "favs" else self.L.get("not_found", "No results found")
            self.safe_status_update(msg)
            return

        df = self.all_data.copy()
        genre = self.genre_choice.GetStringSelection()
        allg = self.L.get("all_genres", "All")
        if genre != allg and self.genre_col and self.genre_col in df.columns:
            df = df[df[self.genre_col].str.contains(genre, case=False, na=False, regex=False)]

        query = self.search.GetValue().lower()
        if query:
            mask = df[self.current_col].str.lower().str.contains(query, na=False)
            if self.author_col and self.author_col in df.columns:
                mask |= df[self.author_col].str.lower().str.contains(query, na=False)
            df = df[mask]

        self.filtered_df = df
        self.update_list_from_filtered()
        cnt = len(df)
        msg = self.L.get("items_found", "{} results").format(cnt) if cnt > 0 else self.L.get("no_results", "No results")
        self.safe_status_update(msg)
        if auto_focus and not self.has_focused and cnt > 0:
            self.has_focused = True
            wx.CallAfter(self.list.SetFocus)
        self.summary_text.SetValue("")
        self.summary_text.Hide()

    def apply_active_mode_data(self):
        key = self.get_active_internal_key()
        self.current_data_key = key
        if key == "favs":
            self.all_data = pd.DataFrame(self.favorites)
        else:
            self.all_data = self.memory_cache[key].copy()

        if self.all_data.empty:
            self.current_col = self.genre_col = self.format_col = ""
            self.update_list_from_filtered()
            return

        if key in self.column_cache:
            cols = self.column_cache[key]
            self.current_col = cols['current_col']
            self.author_col = cols['author_col']
            self.genre_col = cols['genre_col']
            self.format_col = cols['format_col']
            self.view_col = cols['view_col']
            self.size_col = cols['size_col']
            self.track_col = cols['track_col']
        else:
            self.identify_columns()
            self.column_cache[key] = {
                'current_col': self.current_col,
                'author_col': self.author_col,
                'genre_col': self.genre_col,
                'format_col': self.format_col,
                'view_col': self.view_col,
                'size_col': self.size_col,
                'track_col': self.track_col,
            }
            self.update_genres_list()

        if self.genre_col and self.genre_col in self.all_data.columns:
            self.update_genres_list()

    def check_auto_update(self):
        missing = any(not os.path.exists(self.DATA_INFO[k]["file"]) for k in ["books", "movies"])
        old = any((time.time() - os.path.getmtime(self.DATA_INFO[k]["file"])) > 7 * 24 * 3600 for k in ["books", "movies"] if os.path.exists(self.DATA_INFO[k]["file"]))
        
        if missing or old:
            self.on_force_update(None)
        else:
            self.start_loading_process()

    def start_loading_process(self, is_update=False, auto_focus=False):
        if self.is_busy:
            return
        self.is_busy = True
        self.list.update_list([])
        self.summary_text.SetValue("")
        self.summary_text.Hide()
        if is_update:
            play_app_sound("wait.wav", loop=True)
            self.safe_status_update(self.L.get("updating", "Updating..."))
            threading.Thread(target=self.batch_download_worker, args=(auto_focus,), daemon=True).start()
        else:
            self.safe_status_update(self.L.get("reading", "Reading..."))
            threading.Thread(target=self.preload_all_files, args=(auto_focus,), daemon=True).start()

    def preload_all_files(self, auto_focus):
        for key in ["books", "movies"]:
            info = self.DATA_INFO[key]
            pkl_path = info["pkl"]
            xls_path = info["file"]

            if os.path.exists(pkl_path) and os.path.exists(xls_path):
                if os.path.getmtime(pkl_path) >= os.path.getmtime(xls_path):
                    try:
                        df = pd.read_pickle(pkl_path)
                        self.memory_cache[key] = df
                        if key not in self.column_cache:
                            temp_all = self.all_data
                            self.all_data = df
                            self.identify_columns()
                            self.column_cache[key] = {
                                'current_col': self.current_col,
                                'author_col': self.author_col,
                                'genre_col': self.genre_col,
                                'format_col': self.format_col,
                                'view_col': self.view_col,
                                'size_col': self.size_col,
                                'track_col': self.track_col,
                            }
                            self.all_data = temp_all
                        continue
                    except:
                        pass
            if os.path.exists(xls_path):
                try:
                    with open(xls_path, 'rb') as f:
                        data = f.read()
                    txt = data.decode('utf-16') if data.startswith(b'\xff\xfe') else data.decode('utf-8', errors='ignore')
                    try:
                        df = pd.read_html(StringIO(txt))[0]
                    except:
                        df = pd.read_csv(StringIO(txt), sep='\t')
                    df.columns = [str(c).strip() for c in df.columns]
                    df = df.dropna(how='all').astype(str)
                    self.memory_cache[key] = df
                    df.to_pickle(pkl_path)
                    temp_all = self.all_data
                    self.all_data = df
                    self.identify_columns()
                    self.column_cache[key] = {
                        'current_col': self.current_col,
                        'author_col': self.author_col,
                        'genre_col': self.genre_col,
                        'format_col': self.format_col,
                        'view_col': self.view_col,
                        'size_col': self.size_col,
                        'track_col': self.track_col,
                    }
                    self.all_data = temp_all
                except Exception as e:
                    self.memory_cache[key] = pd.DataFrame()
            else:
                self.memory_cache[key] = pd.DataFrame()
        wx.CallAfter(self.on_process_complete, False, auto_focus)

    def batch_download_worker(self, auto_focus):
        import urllib.request as ureq
        for k, info in self.DATA_INFO.items():
            try:
                if k == "books":
                    cat_name = self.L.get("catalog_books", "books")
                else:
                    cat_name = self.L.get("catalog_movies", "movies")
                
                msg_start = self.L.get("downloading_catalog", "Downloading: {}").format(cat_name)
                self.safe_status_update(msg_start)
                
                req = ureq.Request(info["url"], headers={'User-Agent': 'Mozilla/5.0'})
                with ureq.urlopen(req, timeout=120) as r:
                    total = int(r.headers.get('content-length', 0))
                    downloaded = 0
                    last_percent = -1
                    with open(info["file"], 'wb') as f:
                        while True:
                            chunk = r.read(8192)
                            if not chunk:
                                break
                            f.write(chunk)
                            downloaded += len(chunk)
                            if total > 0:
                                percent = int(downloaded * 100 / total)
                                if percent != last_percent:
                                    last_percent = percent
                                    msg_percent = f"{cat_name}: %{percent}"
                                    wx.CallAfter(self._do_status, msg_percent)
                                    wx.SafeYield()
                if os.path.exists(info["pkl"]):
                    os.remove(info["pkl"])
                msg_end = self.L.get("catalog_downloaded", "Downloaded: {}").format(cat_name)
                self.safe_status_update(msg_end)
            except Exception as e:
                error_msg = str(e)
                if "10060" in error_msg or "timed out" in error_msg or "connection" in error_msg.lower():
                    error_msg = self.L.get("connection_timeout_error", "Internet connection failed or GETEM server unreachable. Please check your connection and try again.")
                wx.CallAfter(self.on_error, error_msg)
        self.preload_all_files(auto_focus)

    def on_process_complete(self, is_update, auto_focus):
        stop_all_sounds()
        self.is_busy = False
        if is_update:
            play_app_sound("v8ok.wav")
        self.column_cache.clear()
        self.apply_active_mode_data()
        self.refresh_ui(auto_focus)
        if auto_focus or not self.has_focused:
            wx.CallAfter(self._set_focus_to_list)

    def on_mode_change(self, e):
        self.apply_active_mode_data()
        self.refresh_ui(auto_focus=True)
        self._announce_mode()
        self.summary_text.SetValue("")
        self.summary_text.Hide()

    def on_genre_filter(self, e):
        self.refresh_ui()

    def on_search(self, e):
        self.refresh_ui()
        
    def on_select(self, e):
        idx = self.list.GetFirstSelected()
        if idx != -1 and not self.is_busy:
            play_app_sound("list.wav")
            self.summary_text.SetValue("")
            self.summary_text.Hide()

    def load_favorites(self):
        if os.path.exists(self.fav_file):
            try:
                with open(self.fav_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return []
        return []

    def save_favorites(self):
        with open(self.fav_file, "w", encoding="utf-8") as f:
            json.dump(self.favorites, f, ensure_ascii=False, indent=4)

    def toggle_favorite(self):
        idx = self.list.GetFirstSelected()
        if idx == -1:
            return
        row = self.filtered_df.iloc[idx].to_dict()
        if '_display' in row:
            del row['_display']
        title = row.get(self.current_col, '')
        exists = next((i for i, f in enumerate(self.favorites) if f.get(self.current_col) == title), -1)
        if exists != -1:
            self.favorites.pop(exists)
            self.safe_status_update(self.L.get("fav_removed", "Removed from favorites"))
        else:
            self.favorites.append(row)
            self.safe_status_update(self.L.get("fav_added", "Added to favorites"))
        self.save_favorites()
        if self.get_active_internal_key() == "favs":
            self.apply_active_mode_data()
            self.refresh_ui()

    def on_get_details(self, e):
        idx = self.list.GetFirstSelected()
        if idx == -1:
            nvda_speak(self.L.get("select_eser_first", "Please select a work first"))
            return
        row = self.filtered_df.iloc[idx]
        link = None
        for c in row.index:
            if any(x in str(c).lower() for x in ["link", "url", "baglanti", "baglantı", "bağlantı"]):
                v = str(row[c]).strip()
                if v.startswith("http"):
                    link = v
                    break
        if not link:
            title = str(row[self.current_col]).strip() if self.current_col in row else ""
            link = f"https://getem.boun.edu.tr/katalog?eser_adi={urllib.parse.quote(title)}"
        self.safe_status_update(self.L.get("details_fetching", "Fetching details..."))
        threading.Thread(target=self._fetch_details_selenium, args=(link,), daemon=True).start()

    def _fetch_details_selenium(self, url):
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        driver = None
        try:
            opts = Options()
            opts.add_argument("--headless")
            opts.add_argument("--disable-gpu")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--log-level=3")
            opts.add_argument(f"--user-data-dir={self.profile_dir}")
            drv_path = get_chrome_driver_path(self.L)
            if not drv_path:
                wx.CallAfter(wx.MessageBox, self.L.get("driver_not_found", "ChromeDriver not found"), self.L.get("error_title", "Error"), wx.OK | wx.ICON_ERROR)
                return
            service = Service(drv_path)
            driver = webdriver.Chrome(service=service, options=opts)
            driver.get(url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

            content = ""
            try:
                body_div = driver.find_element(By.CSS_SELECTOR, "div.field-name-body div.field-item")
                raw = body_div.text.strip()
                if raw:
                    for kesici in ["Yorumlar", "Yeni yorum ekle", "Sisteme Giriş Tarihi", "Etiketler"]:
                        idx2 = raw.lower().find(kesici.lower())
                        if idx2 != -1:
                            raw = raw[:idx2].strip()
                    if raw:
                        content = raw
            except:
                pass
            if not content:
                meta = driver.find_elements(By.XPATH, "//meta[@name='description']")
                if meta:
                    content = meta[0].get_attribute("content").strip()
            if not content:
                for sel in [".field-name-body", ".content", "#konu", ".konu", "div.konu", "span.konu"]:
                    elems = driver.find_elements(By.CSS_SELECTOR, sel)
                    if elems:
                        raw = elems[0].text.strip()
                        for kesici in ["Yorumlar", "Yeni yorum ekle", "Sisteme Giriş Tarihi", "Etiketler"]:
                            idx2 = raw.lower().find(kesici.lower())
                            if idx2 != -1:
                                raw = raw[:idx2].strip()
                        if raw:
                            content = raw
                            break
            if not content:
                page_text = driver.find_element(By.TAG_NAME, "body").text
                patterns = [
                    r"Konusu:\s*(.*?)(?:\n\n|\nYorumlar|\nYeni yorum ekle|$)",
                    r"Konu:\s*(.*?)(?:\n\n|\nYorumlar|$)",
                    r"Özet:\s*(.*?)(?:\n\n|\nYorumlar|$)",
                    r"Açıklama:\s*(.*?)(?:\n\n|\nYorumlar|$)"
                ]
                for pat in patterns:
                    match = re.search(pat, page_text, re.IGNORECASE | re.DOTALL)
                    if match:
                        raw = match.group(1).strip()
                        raw = re.sub(r'\s+', ' ', raw)
                        if len(raw) > 20:
                            content = raw
                            break
            if not content or len(content) < 10:
                content = self.L.get("no_details", "No detail information found.")
            content = re.sub(r'\s+', ' ', content).strip()

            wx.CallAfter(wx.MessageBox, content, self.L.get("show_details", "Work Description"), wx.OK | wx.ICON_INFORMATION)
            wx.CallAfter(self.safe_status_update, self.L.get("details_shown", "Details shown"))
        except Exception as e:
            wx.CallAfter(wx.MessageBox, f"Error: {str(e)}", self.L.get("error_title", "Error"), wx.OK | wx.ICON_ERROR)
        finally:
            if driver:
                driver.quit()

    def on_stop_download(self, e):
        global download_stop_flag
        download_stop_flag = True
        self.safe_status_update(self.L.get("status_stopping", "Stopping..."))
        self.btn_stop.Enable(False)

    def trigger_download(self):
        idx = self.list.GetFirstSelected()
        if idx == -1:
            return
        if self.is_busy:
            nvda_speak(self.L.get("busy_msg", "Another operation in progress"))
            return
        if not self.settings.get("getem_user") or not self.settings.get("getem_pass"):
            dlg = GETEMLoginDialog(self, self.L)
            if dlg.ShowModal() == wx.ID_OK:
                self.settings["getem_user"] = dlg.user_txt.GetValue()
                self.settings["getem_pass"] = dlg.pass_txt.GetValue()
                self.save_settings()
                self.is_logged_in = True
                self.create_menu()
            else:
                dlg.Destroy()
                return
            dlg.Destroy()
        row = self.filtered_df.iloc[idx].to_dict()
        title = str(row.get(self.current_col, '')).strip()
        save_dir = self.settings.get("download_dir", "")
        if not save_dir or not os.path.exists(save_dir):
            dlg = wx.DirDialog(self, self.L.get("save_dir_prompt", "Save to:"), style=wx.DD_DEFAULT_STYLE)
            if dlg.ShowModal() == wx.ID_CANCEL:
                return
            save_dir = dlg.GetPath()
            self.settings["download_dir"] = save_dir
            self.save_settings()
        global download_stop_flag
        download_stop_flag = False
        self.btn_stop.Enable(True)
        self.progress_bar.Show()
        self.Layout()
        threading.Thread(target=self.download_worker_thread, args=(row, save_dir, title), daemon=True).start()

    def download_worker_thread(self, row_dict, save_dir, title):
        global download_stop_flag
        self.is_busy = True
        wx.CallAfter(play_app_sound, "wait.wav", True)
        self.safe_status_update(self.L.get("msg_browser_start", "Starting..."))
        wx.CallAfter(self.progress_bar.SetValue, 0)
        driver = None
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC

            link = None
            for k, v in row_dict.items():
                if any(x in str(k).lower() for x in ["link", "url", "baglanti", "baglantı", "bağlantı"]):
                    val = str(v).strip()
                    if val.startswith("http"):
                        link = val
                        break
            if not link:
                self.safe_status_update(self.L.get("conn_error", "Link not found"))
                self.is_busy = False
                wx.CallAfter(self.btn_stop.Enable, False)
                wx.CallAfter(self.progress_bar.Hide)
                return

            opts = Options()
            opts.add_argument("--headless")
            opts.add_argument("--disable-gpu")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--log-level=3")
            opts.add_argument(f"--user-data-dir={self.profile_dir}")
            opts.page_load_strategy = 'eager'
            drv_path = get_chrome_driver_path(self.L)
            if not drv_path:
                self.safe_status_update(self.L.get("driver_not_found", "ChromeDriver not found"))
                self.is_busy = False
                wx.CallAfter(self.btn_stop.Enable, False)
                wx.CallAfter(self.progress_bar.Hide)
                return
            service = Service(drv_path)
            driver = webdriver.Chrome(service=service, options=opts)
            driver.set_page_load_timeout(30)
            driver.get(link)

            if driver.find_elements(By.NAME, "pass"):
                self.safe_status_update(self.L.get("status_login", "Logging in..."))
                driver.find_element(By.NAME, "name").send_keys(self.settings.get("getem_user", ""))
                inp = driver.find_element(By.NAME, "pass")
                inp.send_keys(self.settings.get("getem_pass", ""))
                inp.submit()
                time.sleep(3)
                if re.search(r"üzgünüz|invalid|hatalı", driver.page_source, re.I):
                    wx.CallAfter(self.on_error, self.L.get("login_failed", "Login failed"))
                    driver.quit()
                    self.is_busy = False
                    return
                driver.get(link)

            try:
                wait = WebDriverWait(driver, 10)
                dl_elem = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href,'downloadZip.php')]")))
                fin_url = dl_elem.get_attribute("href")
            except:
                elems = driver.find_elements(By.XPATH, "//a[contains(@href,'downloadZip.php')]")
                if not elems:
                    wx.CallAfter(self.on_error, self.L.get("msg_btn_not_found", "Download button not found"))
                    driver.quit()
                    self.is_busy = False
                    return
                fin_url = elems[0].get_attribute("href")

            sess = requests.Session()
            for c in driver.get_cookies():
                sess.cookies.set(c['name'], c['value'])
            driver.quit()
            driver = None

            self.safe_status_update(self.L.get("download_started", "Downloading: {}").format(title))
            r = sess.get(fin_url, stream=True, timeout=60)
            if 'text/html' in r.headers.get('Content-Type', ''):
                match = re.search(r'(\d+)\s*saniye', r.text[:500], re.I)
                if match:
                    beklenen_saniye = match.group(1)
                    wx.CallAfter(self.on_error, f"Lütfen {beklenen_saniye} saniye bekleyin")
                else:
                    wx.CallAfter(self.on_error, self.L.get("download_denied", "Erişim engellendi"))
                self.is_busy = False
                return

            fname = "".join(c for c in title if c.isalpha() or c.isdigit() or c == ' ').strip()
            path = os.path.join(save_dir, f"{fname}.zip")
            total = int(r.headers.get('content-length', 0))
            down = 0
            last_p = -1
            with open(path, 'wb') as f:
                for chunk in r.iter_content(16384):
                    if download_stop_flag:
                        f.close()
                        os.remove(path)
                        self.safe_status_update(self.L.get("status_stopped", "Stopped"))
                        wx.CallAfter(stop_all_sounds)
                        break
                    if chunk:
                        f.write(chunk)
                        down += len(chunk)
                        if total > 0:
                            p = int(down * 100 / total)
                            if p != last_p:
                                wx.CallAfter(self.progress_bar.SetValue, p)
                                last_p = p
            if not download_stop_flag:
                wx.CallAfter(stop_all_sounds)
                wx.CallAfter(play_app_sound, "v8ok.wav")
                wx.CallAfter(self.progress_bar.SetValue, 100)
                self.safe_status_update(self.L.get("download_complete", "Downloaded: {}").format(title))
        except Exception as ex:
            wx.CallAfter(self.on_error, str(ex))
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            self.is_busy = False
            wx.CallAfter(self.progress_bar.SetValue, 0)
            wx.CallAfter(stop_all_sounds)
            wx.CallAfter(self.btn_stop.Enable, False)
            wx.CallAfter(self.progress_bar.Hide)

    def on_btn_open(self, e):
        self.trigger_download()

    def create_menu(self):
        menubar = wx.MenuBar()
        L = self.L
        fileM = wx.Menu()
        if self.is_logged_in:
            acc = fileM.Append(wx.ID_ANY, L.get("menu_logout", "Logout"))
        else:
            acc = fileM.Append(wx.ID_ANY, L.get("menu_account", "Account"))
        fileM.AppendSeparator()
        updCat = fileM.Append(wx.ID_ANY, L.get("update", "Update catalog"))
        fileM.AppendSeparator()
        exitM = fileM.Append(wx.ID_EXIT, L.get("exit", "Exit"))
        menubar.Append(fileM, L.get("menu_file", "File"))

        optM = wx.Menu()
        langSub = wx.Menu()
        en = langSub.Append(wx.ID_ANY, "English", kind=wx.ITEM_RADIO)
        tr = langSub.Append(wx.ID_ANY, "Türkçe", kind=wx.ITEM_RADIO)
        az = langSub.Append(wx.ID_ANY, "Azərbaycanca", kind=wx.ITEM_RADIO)
        if self.current_lang == "EN":
            en.Check()
        elif self.current_lang == "TR":
            tr.Check()
        else:
            az.Check()
        optM.AppendSubMenu(langSub, L.get("lang", "Language"))
        optM.AppendSeparator()
        downDir = optM.Append(wx.ID_ANY, L.get("menu_download_dir", "Download folder"))
        optM.AppendSeparator()
        viewSub = wx.Menu()
        vb = viewSub.Append(wx.ID_ANY, L.get("view_books", "Books"), kind=wx.ITEM_RADIO)
        vm = viewSub.Append(wx.ID_ANY, L.get("view_movies", "Movies"), kind=wx.ITEM_RADIO)
        vf = viewSub.Append(wx.ID_ANY, L.get("view_favs", "Favorites"), kind=wx.ITEM_RADIO)
        dv = self.settings.get("default_view", "books")
        if dv == "books":
            vb.Check()
        elif dv == "movies":
            vm.Check()
        else:
            vf.Check()
        optM.AppendSubMenu(viewSub, L.get("default_start", "Startup view"))
        optM.AppendSeparator()
        snd = optM.Append(wx.ID_ANY, L.get("disable_sound", "Disable sound"), kind=wx.ITEM_CHECK)
        snd.Check(not SOUND_ENABLED)
        optM.AppendSeparator()
        clear = optM.Append(wx.ID_ANY, L.get("clear_data", "Clear all data"))
        menubar.Append(optM, L.get("menu_options_label", "Settings"))

        helpM = wx.Menu()
        about = helpM.Append(wx.ID_ABOUT, L.get("about", "About"))
        helpDoc = helpM.Append(wx.ID_ANY, L.get("help_doc", "Help (F1)"))
        chkUpd = helpM.Append(wx.ID_ANY, L.get("menu_check_updates", "Check for updates"))
        menubar.Append(helpM, L.get("menu_help", "Help"))

        self.SetMenuBar(menubar)

        self.Bind(wx.EVT_MENU, self.on_account_menu_click, acc)
        self.Bind(wx.EVT_MENU, self.on_force_update, updCat)
        self.Bind(wx.EVT_MENU, lambda e: self.Close(), exitM)
        self.Bind(wx.EVT_MENU, lambda e: self.change_lang("EN"), en)
        self.Bind(wx.EVT_MENU, lambda e: self.change_lang("TR"), tr)
        self.Bind(wx.EVT_MENU, lambda e: self.change_lang("AZ"), az)
        self.Bind(wx.EVT_MENU, self.on_set_download_dir, downDir)
        self.Bind(wx.EVT_MENU, lambda e: self.set_default_view("books"), vb)
        self.Bind(wx.EVT_MENU, lambda e: self.set_default_view("movies"), vm)
        self.Bind(wx.EVT_MENU, lambda e: self.set_default_view("favs"), vf)
        self.Bind(wx.EVT_MENU, self.on_toggle_sound, snd)
        self.Bind(wx.EVT_MENU, self.on_clear_all_data, clear)
        self.Bind(wx.EVT_MENU, self.on_about, about)
        self.Bind(wx.EVT_MENU, self.on_help_doc, helpDoc)
        self.Bind(wx.EVT_MENU, self.on_check_updates_click, chkUpd)

    def on_account_menu_click(self, e):
        if self.is_logged_in:
            self.settings["getem_user"] = self.settings["getem_pass"] = ""
            self.save_settings()
            self.is_logged_in = False
            nvda_speak(self.L.get("msg_logged_out", "Logged out"))
            self.create_menu()
        else:
            dlg = GETEMLoginDialog(self, self.L)
            dlg.user_txt.SetValue(self.settings.get("getem_user", ""))
            if dlg.ShowModal() == wx.ID_OK:
                self.settings["getem_user"] = dlg.user_txt.GetValue()
                self.settings["getem_pass"] = dlg.pass_txt.GetValue()
                self.save_settings()
                self.is_logged_in = True
                wx.MessageBox(self.L.get("account_saved", "Credentials saved"), self.L.get("info_title", "Info"), wx.OK)
                self.create_menu()
            dlg.Destroy()

    def on_set_download_dir(self, e):
        dlg = wx.DirDialog(self, self.L.get("menu_download_dir", "Select folder"), defaultPath=self.settings.get("download_dir", ""))
        if dlg.ShowModal() == wx.ID_OK:
            self.settings["download_dir"] = dlg.GetPath()
            self.save_settings()

    def change_lang(self, code):
        if code == self.current_lang:
            return
        self.current_lang = code
        self.settings["lang"] = code
        self.save_settings()
        self.L = self.load_language_files()
        self.apply_language()
        self.apply_active_mode_data()
        self.has_focused = False
        self.refresh_ui(auto_focus=True)
        self.list.SetFocus()

    def on_about(self, e):

        cur = get_bundled_version()
        about = self.L.get("about_text", "GETEM Catalog Pro - Version: {version}").format(version=cur)
        wx.MessageBox(about, self.L.get("about", "About"), wx.OK)

    def on_help_doc(self, e):
        self.open_help()

    def open_help(self):
        hf = f"{self.current_lang.lower()}.html"
        p = resource_path(os.path.join("help", hf))
        if os.path.exists(p):
            webbrowser.open(f"file://{p}")
            self.safe_status_update(self.L.get("opening_help", "Opening help..."))
        else:
            eng = resource_path(os.path.join("help", "en.html"))
            if os.path.exists(eng):
                webbrowser.open(f"file://{eng}")
            else:
                wx.MessageBox(self.L.get("help_not_found", "Help file not found!"), self.L.get("error_title", "Error"), wx.ICON_ERROR)

    def set_default_view(self, v):
        self.settings["default_view"] = v
        self.save_settings()

    def on_toggle_sound(self, e):
        global SOUND_ENABLED
        SOUND_ENABLED = not e.IsChecked()
        self.settings["sound_enabled"] = SOUND_ENABLED
        self.save_settings()

    def on_clear_all_data(self, e):
        if wx.MessageBox(self.L.get("clear_data_confirm", "Delete all user data?"), self.L.get("clear_data_title", "Clear data"), wx.YES_NO | wx.ICON_WARNING) == wx.YES:
            shutil.rmtree(self.cache_dir, ignore_errors=True)
            shutil.rmtree(self.profile_dir, ignore_errors=True)
            os.makedirs(self.cache_dir, exist_ok=True)
            os.makedirs(self.profile_dir, exist_ok=True)
            self.settings = {"lang": self.current_lang, "sound_enabled": True}
            self.save_settings()
            self.favorites = []
            self.save_favorites()
            self.memory_cache = {"books": pd.DataFrame(), "movies": pd.DataFrame()}
            self.column_cache.clear()
            self.all_data = pd.DataFrame()
            self.filtered_df = pd.DataFrame()
            self.restart_program()

    def restart_program(self):
        if getattr(sys, 'frozen', False):
            subprocess.Popen([sys.executable], shell=True)
            self.Close()
        else:
            subprocess.Popen([sys.executable, __file__])
            self.Close()

    def on_force_update(self, e):
        self.column_cache.clear()
        self.start_loading_process(is_update=True, auto_focus=True)

    def on_check_updates_click(self, e):
        self.check_for_app_update(manual=True, trigger_catalog_after=False)
        self.safe_status_update(self.L.get("update_check", "Checking for updates..."))

    def on_error(self, msg):
        stop_all_sounds()
        self.is_busy = False
        play_app_sound("v8nome.wav")
        self.safe_status_update(self.L.get("error_title", "Error"))
        wx.MessageBox(msg, self.L.get("error_title", "Error"), wx.ICON_ERROR)

    def on_right_click(self, e):
        if self.list.GetFirstSelected() == -1:
            return
        menu = wx.Menu()
        sort_sub = wx.Menu()
        sort_sub.Append(wx.ID_ANY, self.L.get("sort_default", "Original"))
        sort_sub.AppendSeparator()
        sort_sub.Append(wx.ID_ANY, self.L.get("sort_title_asc", "Title A-Z"))
        sort_sub.Append(wx.ID_ANY, self.L.get("sort_title_desc", "Title Z-A"))
        sort_sub.AppendSeparator()
        if self.format_col:
            sort_sub.Append(wx.ID_ANY, self.L.get("sort_format_asc", "Format A-Z"))
            sort_sub.Append(wx.ID_ANY, self.L.get("sort_format_desc", "Format Z-A"))
            sort_sub.AppendSeparator()
        if self.view_col:
            sort_sub.Append(wx.ID_ANY, self.L.get("sort_view_asc", "Views low-high"))
            sort_sub.Append(wx.ID_ANY, self.L.get("sort_view_desc", "Views high-low"))
            sort_sub.AppendSeparator()
        if self.size_col:
            sort_sub.Append(wx.ID_ANY, self.L.get("sort_size_asc", "Size small-large"))
            sort_sub.Append(wx.ID_ANY, self.L.get("sort_size_desc", "Size large-small"))
            sort_sub.AppendSeparator()
        if self.track_col:
            sort_sub.Append(wx.ID_ANY, self.L.get("sort_track_asc", "Tracks low-high"))
            sort_sub.Append(wx.ID_ANY, self.L.get("sort_track_desc", "Tracks high-low"))
            sort_sub.AppendSeparator()
        menu.AppendSubMenu(sort_sub, self.L.get("sort_by", "Sort by"))
        menu.AppendSeparator()
        if self.is_busy:
            stop = menu.Append(wx.ID_ANY, self.L.get("stop", "Stop"))
            self.Bind(wx.EVT_MENU, lambda ev: self.on_stop_download(ev), stop)
            menu.AppendSeparator()
        down = menu.Append(wx.ID_ANY, self.L.get("open", "Download (Ctrl+E)"))
        fav = menu.Append(wx.ID_ANY, self.L.get("ctx_fav", "Favorite (Ctrl+D)"))
        det = menu.Append(wx.ID_ANY, self.L.get("get_details", "Show Description (Ctrl+K)"))

        menu_ids = sort_sub.GetMenuItems()
        self.Bind(wx.EVT_MENU, lambda e: self.sort_by('default', True), menu_ids[0])
        self.Bind(wx.EVT_MENU, lambda e: self.sort_by('title', True), menu_ids[2])
        self.Bind(wx.EVT_MENU, lambda e: self.sort_by('title', False), menu_ids[3])
        idx = 5
        if self.format_col:
            self.Bind(wx.EVT_MENU, lambda e: self.sort_by('format', True), menu_ids[idx])
            self.Bind(wx.EVT_MENU, lambda e: self.sort_by('format', False), menu_ids[idx+1])
            idx += 3
        if self.view_col:
            self.Bind(wx.EVT_MENU, lambda e: self.sort_by('view', True), menu_ids[idx])
            self.Bind(wx.EVT_MENU, lambda e: self.sort_by('view', False), menu_ids[idx+1])
            idx += 3
        if self.size_col:
            self.Bind(wx.EVT_MENU, lambda e: self.sort_by('size', True), menu_ids[idx])
            self.Bind(wx.EVT_MENU, lambda e: self.sort_by('size', False), menu_ids[idx+1])
            idx += 3
        if self.track_col:
            self.Bind(wx.EVT_MENU, lambda e: self.sort_by('track', True), menu_ids[idx])
            self.Bind(wx.EVT_MENU, lambda e: self.sort_by('track', False), menu_ids[idx+1])

        self.Bind(wx.EVT_MENU, lambda ev: self.trigger_download(), down)
        self.Bind(wx.EVT_MENU, lambda ev: self.toggle_favorite(), fav)
        self.Bind(wx.EVT_MENU, self.on_get_details, det)
        self.PopupMenu(menu)
        menu.Destroy()

    def on_key(self, e):
        key = e.GetKeyCode()
        ctrl = e.ControlDown()
        if ctrl and key == ord('B'):
            self.mode_choice.SetSelection(0)
            self.on_mode_change(None)
            self.safe_status_update(self.L.get("view_books", "Books"))
        elif ctrl and key == ord('M'):
            self.mode_choice.SetSelection(1)
            self.on_mode_change(None)
            self.safe_status_update(self.L.get("view_movies", "Movies"))
        elif ctrl and key == ord('F'):
            self.mode_choice.SetSelection(2)
            self.on_mode_change(None)
            self.safe_status_update(self.L.get("view_favs", "Favorites"))
        elif ctrl and key == ord('K'):
            self.on_get_details(None)
        elif key == wx.WXK_RETURN:
            idx = self.list.GetFirstSelected()
            if idx != -1:
                row = self.filtered_df.iloc[idx]
                link = None
                for c in row.index:
                    if any(x in str(c).lower() for x in ["link", "url", "baglanti", "baglantı", "bağlantı"]):
                        val = str(row[c]).strip()
                        if val.startswith("http"):
                            link = val
                            break
                if link:
                    webbrowser.open(link)
                else:
                    title = str(row[self.current_col]).strip() if self.current_col in row else ""
                    webbrowser.open(f"https://getem.boun.edu.tr/katalog?eser_adi={urllib.parse.quote(title)}")
                self.safe_status_update(self.L.get("opening_browser", "Opening browser"))
        elif ctrl and key == ord('D'):
            self.toggle_favorite()
        elif ctrl and key == ord('E'):
            self.trigger_download()
        elif key == wx.WXK_ESCAPE and self.is_busy:
            self.on_stop_download(None)
        else:
            e.Skip()

    def on_global_key(self, e):
        if e.GetKeyCode() == wx.WXK_F1:
            self.open_help()
        else:
            e.Skip()

if __name__ == '__main__':
    app = wx.App()
    GetemApp()
    app.MainLoop()