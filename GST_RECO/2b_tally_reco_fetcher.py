"""
Purchase Register Fetcher (Purchase All only)
Wrapper around 2b_tally_reco_fetcher to lock the voucher type.
"""

from __future__ import annotations

import importlib.util
import os

import customtkinter as ctk


def _load_base_module():
	base_path = os.path.join(os.path.dirname(__file__), "2b_tally_reco_fetcher.py")
	spec = importlib.util.spec_from_file_location("tally_reco_base", base_path)
	if spec is None or spec.loader is None:
		raise ImportError(f"Unable to load base module at {base_path}")
	module = importlib.util.module_from_spec(spec)
	spec.loader.exec_module(module)
	return module


_base = _load_base_module()
BaseApp = _base.PurchaseRegisterApp


class PurchaseAllApp(BaseApp):
	def __init__(self):
		super().__init__()
		self.title("Tally Purchase Register Fetcher - Purchase All")
		self.vchr_type_var.set("Purchase All")
		self._hide_voucher_type_controls()

	def _fetch_register_thread(self):
		self.vchr_type_var.set("Purchase All")
		super()._fetch_register_thread()

	def _hide_voucher_type_controls(self):
		for widget in self._walk_children(self):
			if isinstance(widget, ctk.CTkOptionMenu):
				try:
					widget.pack_forget()
				except Exception:
					pass
			elif isinstance(widget, ctk.CTkLabel):
				try:
					if widget.cget("text") == "Type":
						widget.pack_forget()
				except Exception:
					pass

	def _walk_children(self, widget):
		for child in widget.winfo_children():
			yield child
			yield from self._walk_children(child)


if __name__ == "__main__":
	app = PurchaseAllApp()
	app.mainloop()
