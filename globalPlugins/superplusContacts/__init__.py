import os
import json
import csv
import wx
import gui
import addonHandler
import globalPluginHandler
from logHandler import log

# NVDA Çeviri Sistemini Başlat
addonHandler.initTranslation()

class ContactAddDialog(wx.Dialog):
	"""Yeni kişi eklemek için çoklu kutulu özel pencere."""
	def __init__(self, parent):
		super(ContactAddDialog, self).__init__(parent, title=_("Yeni Kişi Ekle"), size=(400, 350))
		sizer = wx.BoxSizer(wx.VERTICAL)
		
		# Giriş Alanları
		self.first_name = self._create_input(sizer, _("Ad:"))
		self.last_name = self._create_input(sizer, _("Soyad:"))
		self.phone = self._create_input(sizer, _("Telefon:"))
		self.email = self._create_input(sizer, _("E-posta:"))
		
		# Butonlar
		btn_sizer = self.CreateButtonSizer(wx.OK | wx.CANCEL)
		sizer.Add(btn_sizer, 0, wx.ALL | wx.CENTER, 10)
		
		self.SetSizer(sizer)
		self.Layout()

	def _create_input(self, sizer, label):
		lbl = wx.StaticText(self, label=label)
		txt = wx.TextCtrl(self)
		sizer.Add(lbl, 0, wx.ALL | wx.EXPAND, 5)
		sizer.Add(txt, 0, wx.ALL | wx.EXPAND, 5)
		return txt

	def get_data(self):
		return {
			"first_name": self.first_name.GetValue().strip(),
			"last_name": self.last_name.GetValue().strip(),
			"phone": self.phone.GetValue().strip(),
			"email": self.email.GetValue().strip()
		}

class ContactsDialog(wx.Dialog):
	def __init__(self, parent):
		super(ContactsDialog, self).__init__(parent, title=_("SuperPlus Rehber Yöneticisi"), size=(700, 500))
		self.contacts_file = os.path.join(os.path.dirname(__file__), "contacts.json")
		self.contacts = self.load_contacts()
		
		main_sizer = wx.BoxSizer(wx.VERTICAL)
		
		# Liste Görünümü (Ad, Soyad, Telefon, E-posta)
		self.list_ctrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.SUNKEN_BORDER)
		self.list_ctrl.InsertColumn(0, _("Ad"), width=120)
		self.list_ctrl.InsertColumn(1, _("Soyad"), width=120)
		self.list_ctrl.InsertColumn(2, _("Telefon"), width=150)
		self.list_ctrl.InsertColumn(3, _("E-posta"), width=180)
		main_sizer.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 10)
		self.refresh_list()

		# Butonlar
		btn_sizer = wx.WrapSizer(wx.HORIZONTAL)
		btns = [
			(_("Ekle"), self.on_add),
			(_("Sil"), self.on_delete),
			(_("İçe Aktar"), self.on_import),
			(_("Dışa Aktar"), self.on_export),
			(_("Kapat"), lambda e: self.Close())
		]

		for label, handler in btns:
			btn = wx.Button(self, label=label)
			btn.Bind(wx.EVT_BUTTON, handler)
			btn_sizer.Add(btn, 0, wx.ALL, 5)

		main_sizer.Add(btn_sizer, 0, wx.CENTER | wx.BOTTOM, 10)
		self.SetSizer(main_sizer)

	def load_contacts(self):
		if os.path.exists(self.contacts_file):
			try:
				with open(self.contacts_file, "r", encoding="utf-8") as f:
					return json.load(f)
			except: return []
		return []

	def save_contacts(self):
		with open(self.contacts_file, "w", encoding="utf-8") as f:
			json.dump(self.contacts, f, ensure_ascii=False, indent=4)

	def refresh_list(self):
		self.list_ctrl.DeleteAllItems()
		for c in self.contacts:
			self.list_ctrl.Append([c.get("first_name", ""), c.get("last_name", ""), c.get("phone", ""), c.get("email", "")])

	def on_add(self, evt):
		dlg = ContactAddDialog(self)
		if dlg.ShowModal() == wx.ID_OK:
			self.contacts.append(dlg.get_data())
			self.save_contacts()
			self.refresh_list()
		dlg.Destroy()

	def on_delete(self, evt):
		idx = self.list_ctrl.GetFirstSelected()
		if idx != -1:
			del self.contacts[idx]
			self.save_contacts()
			self.refresh_list()

	def on_import(self, evt):
		with wx.FileDialog(self, _("CSV Seç"), wildcard="CSV files (*.csv)|*.csv", style=wx.FD_OPEN) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				try:
					with open(dlg.GetPath(), "r", encoding="utf-8-sig") as f:
						reader = csv.DictReader(f)
						for row in reader:
							# Akıllı Eşleştirme (Google/Outlook/SuperPlus)
							fn = row.get("Ad") or row.get("First Name") or row.get("Given Name", "")
							ln = row.get("Soyad") or row.get("Last Name") or row.get("Family Name", "")
							ph = row.get("Telefon") or row.get("Mobile Phone") or row.get("Phone 1 - Value", "")
							em = row.get("E-posta") or row.get("E-mail Address") or row.get("E-mail 1 - Value", "")
							self.contacts.append({"first_name": fn, "last_name": ln, "phone": ph, "email": em})
					self.save_contacts()
					self.refresh_list()
					gui.messageBox(_("İşlem Başarılı"), _("Bilgi"))
				except Exception as e:
					gui.messageBox(str(e), _("Hata"))

	def on_export(self, evt):
		with wx.FileDialog(self, _("Rehberi Kaydet"), wildcard="CSV files (*.csv)|*.csv", style=wx.FD_SAVE) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				with open(dlg.GetPath(), "w", encoding="utf-8", newline="") as f:
					# CSV başlıkları senin istediğin gibi
					fieldnames = ["Ad", "Soyad", "Telefon", "E-posta"]
					writer = csv.DictWriter(f, fieldnames=fieldnames)
					writer.writeheader()
					for c in self.contacts:
						writer.writerow({
							"Ad": c.get("first_name", ""),
							"Soyad": c.get("last_name", ""),
							"Telefon": c.get("phone", ""),
							"E-posta": c.get("email", "")
						})
				gui.messageBox(_("Dışa aktarma tamamlandı"), _("Bilgi"))

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def __init__(self):
		super(GlobalPlugin, self).__init__()
		self.menu_item = gui.mainFrame.sysTrayIcon.toolsMenu.Append(wx.ID_ANY, _("SuperPlus Rehber..."))
		gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, lambda e: ContactsDialog(gui.mainFrame).ShowModal(), self.menu_item)