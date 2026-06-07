import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
import tkinter as tk
from datetime import datetime
import threading
import os

from config import load_config, save_config, COLUMNS, COL_WIDTHS
from sheets_manager import SheetsManager
from export_manager import export_to_excel, export_to_csv

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


def _safe_float(val):
    try:
        return float(str(val).replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError):
        return 0.0


class InventaireApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gestionnaire d'Inventaire")
        self.geometry("1280x720")
        self.minsize(950, 600)
        self.config_data = load_config()
        self.sheets_mgr = None
        self.all_data = []
        self.filtered_data = []
        self.row_indices = []
        self._build_ui()
        self.after(200, self._try_auto_connect)

    # ------------------------------------------------------------------ UI ---

    def _build_ui(self):
        # ── Top bar ──────────────────────────────────────────────────────────
        top = ctk.CTkFrame(self, height=58, corner_radius=0, fg_color="#1F4E79")
        top.pack(fill="x")
        top.pack_propagate(False)

        ctk.CTkLabel(top, text="📦  Gestionnaire d'Inventaire",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color="white").pack(side="left", padx=20)

        self.status_lbl = ctk.CTkLabel(top, text="Non connecté",
                                        text_color="#ADB9CA",
                                        font=ctk.CTkFont(size=11))
        self.status_lbl.pack(side="right", padx=12)

        self.sync_btn = ctk.CTkButton(top, text="🔄  Synchroniser",
                                       command=self._sync, width=140,
                                       fg_color="#2E75B6", hover_color="#144D82")
        self.sync_btn.pack(side="right", padx=6, pady=8)

        ctk.CTkButton(top, text="⚙  Paramètres",
                       command=self._open_settings, width=120,
                       fg_color="#2E75B6", hover_color="#144D82").pack(side="right", padx=4, pady=8)

        # ── Sheet selector bar ───────────────────────────────────────────────
        self.sheet_bar = ctk.CTkFrame(self, height=38, corner_radius=0, fg_color="#144D82")
        self.sheet_bar.pack(fill="x")
        self.sheet_bar.pack_propagate(False)

        ctk.CTkLabel(self.sheet_bar, text="Onglet :",
                     text_color="#ADB9CA", font=ctk.CTkFont(size=11)).pack(side="left", padx=12)

        self.sheet_var = tk.StringVar(value="—")
        self.sheet_combo = ctk.CTkComboBox(
            self.sheet_bar, variable=self.sheet_var, values=["—"],
            width=220, font=ctk.CTkFont(size=12),
            command=self._on_sheet_change
        )
        self.sheet_combo.pack(side="left", padx=4, pady=4)

        # ── Filter bar ───────────────────────────────────────────────────────
        fbar = ctk.CTkFrame(self, height=48, corner_radius=0, fg_color="#E8F0F7")
        fbar.pack(fill="x")
        fbar.pack_propagate(False)

        ctk.CTkLabel(fbar, text="Filtres :", font=ctk.CTkFont(weight="bold"),
                     text_color="#1F4E79").pack(side="left", padx=12)

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filters())
        ctk.CTkEntry(fbar, textvariable=self.search_var,
                     placeholder_text="Rechercher…", width=190).pack(side="left", padx=4, pady=6)

        ctk.CTkLabel(fbar, text="Catégorie :").pack(side="left", padx=(10, 2))
        self.cat_var = tk.StringVar(value="Toutes")
        self.cat_combo = ctk.CTkComboBox(fbar, variable=self.cat_var, values=["Toutes"],
                                          width=160, command=lambda _: self._apply_filters())
        self.cat_combo.pack(side="left", padx=4)

        ctk.CTkLabel(fbar, text="Fournisseur :").pack(side="left", padx=(10, 2))
        self.fourn_var = tk.StringVar(value="Tous")
        self.fourn_combo = ctk.CTkComboBox(fbar, variable=self.fourn_var, values=["Tous"],
                                            width=160, command=lambda _: self._apply_filters())
        self.fourn_combo.pack(side="left", padx=4)

        ctk.CTkButton(fbar, text="Réinitialiser", command=self._reset_filters,
                       width=100, fg_color="#6B7280", hover_color="#4B5563").pack(side="left", padx=10)

        # ── Table ────────────────────────────────────────────────────────────
        content = ctk.CTkFrame(self, corner_radius=0, fg_color="white")
        content.pack(fill="both", expand=True)
        self._build_table(content)

        # ── Bottom bar ───────────────────────────────────────────────────────
        bot = ctk.CTkFrame(self, height=50, corner_radius=0, fg_color="#F0F4F8")
        bot.pack(fill="x")
        bot.pack_propagate(False)

        self.stats_lbl = ctk.CTkLabel(bot, text="Articles : 0  |  Valeur totale : 0,00 €",
                                       font=ctk.CTkFont(size=11), text_color="#4B5563")
        self.stats_lbl.pack(side="left", padx=15)

        ctk.CTkButton(bot, text="📅  Clôture de mois", command=self._close_month,
                       width=160, fg_color="#059669", hover_color="#047857").pack(side="right", padx=10, pady=8)
        ctk.CTkButton(bot, text="📊  Exporter Excel", command=self._export_excel,
                       width=140, fg_color="#D97706", hover_color="#B45309").pack(side="right", padx=4, pady=8)
        ctk.CTkButton(bot, text="➕  Ajouter", command=self._add_item,
                       width=110, fg_color="#3B82F6", hover_color="#2563EB").pack(side="right", padx=4, pady=8)

    def _build_table(self, parent):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Inv.Treeview", background="white", foreground="#1F2937",
                         rowheight=26, fieldbackground="white",
                         font=("Segoe UI", 9))
        style.configure("Inv.Treeview.Heading", background="#1F4E79", foreground="white",
                         font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Inv.Treeview",
                   background=[("selected", "#DBEAFE")],
                   foreground=[("selected", "#1E3A5F")])
        style.map("Inv.Treeview.Heading", background=[("active", "#2E75B6")])

        tf = tk.Frame(parent, bg="white")
        tf.pack(fill="both", expand=True, padx=4, pady=4)

        vsb = ttk.Scrollbar(tf, orient="vertical")
        hsb = ttk.Scrollbar(tf, orient="horizontal")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(tf, columns=COLUMNS, show="headings",
                                   style="Inv.Treeview",
                                   yscrollcommand=vsb.set,
                                   xscrollcommand=hsb.set,
                                   selectmode="browse")
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        for col, w in zip(COLUMNS, COL_WIDTHS):
            anchor = "w" if col in ("PRODUIT", "FOURNISSEUR", "CATEGORIE") else "center"
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=55, anchor=anchor)

        self.tree.tag_configure("odd",  background="white")
        self.tree.tag_configure("even", background="#F0F7FF")
        self.tree.tag_configure("zero", background="#FEF3C7")  # QTE M31 = 0 → jaune

        self.tree.bind("<Double-1>", self._on_dbl_click)
        self.tree.bind("<Button-3>", self._show_ctx_menu)
        self.tree.pack(fill="both", expand=True)

        self.ctx_menu = tk.Menu(self, tearoff=0)
        self.ctx_menu.add_command(label="✏  Modifier QTE M31", command=self._edit_qty)
        self.ctx_menu.add_command(label="📝  Modifier l'article", command=self._edit_item)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="🗑  Supprimer", command=self._delete_item)

    # ---------------------------------------------------------- Data logic ---

    def _try_auto_connect(self):
        if self.config_data.get("spreadsheet_id"):
            self._sync()
        else:
            self.status_lbl.configure(text="⚠  Configurez la connexion dans Paramètres")

    def _sync(self):
        if not self.config_data.get("spreadsheet_id"):
            messagebox.showwarning("Configuration manquante",
                                   "Entrez l'ID de votre Google Sheet dans Paramètres.")
            self._open_settings()
            return
        self.sync_btn.configure(state="disabled", text="Chargement…")
        self.status_lbl.configure(text="Connexion…", text_color="#F59E0B")

        def worker():
            try:
                if not self.sheets_mgr:
                    self.sheets_mgr = SheetsManager(
                        self.config_data["spreadsheet_id"]
                    )
                self.sheets_mgr.connect()
                sheets = self.sheets_mgr.list_sheets()
                data = self.sheets_mgr.get_all_data()
                current = self.sheets_mgr.sheet_name
                self.after(0, lambda: self._on_data_loaded(data, sheets, current))
            except FileNotFoundError as e:
                self.after(0, lambda: (
                    messagebox.showerror("Fichier manquant", str(e)),
                    self.status_lbl.configure(text="⚠  credentials.json manquant",
                                               text_color="#EF4444")
                ))
            except Exception as e:
                self.after(0, lambda: (
                    messagebox.showerror("Erreur de connexion", str(e)),
                    self.status_lbl.configure(text="✗  Erreur", text_color="#EF4444")
                ))
            finally:
                self.after(0, lambda: self.sync_btn.configure(
                    state="normal", text="🔄  Synchroniser"))

        threading.Thread(target=worker, daemon=True).start()

    def _on_sheet_change(self, sheet_name):
        if not self.sheets_mgr or sheet_name == self.sheets_mgr.sheet_name:
            return
        self.sync_btn.configure(state="disabled", text="Chargement…")
        self.status_lbl.configure(text=f"Chargement {sheet_name}…", text_color="#F59E0B")

        def worker():
            try:
                self.sheets_mgr.switch_sheet(sheet_name)
                data = self.sheets_mgr.get_all_data()
                self.after(0, lambda: self._on_data_loaded(data, None, sheet_name))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
            finally:
                self.after(0, lambda: self.sync_btn.configure(
                    state="normal", text="🔄  Synchroniser"))

        threading.Thread(target=worker, daemon=True).start()

    def _on_data_loaded(self, data, sheets=None, current_sheet=None):
        self.all_data = data
        if sheets:
            self.sheet_combo.configure(values=sheets)
        if current_sheet:
            self.sheet_var.set(current_sheet)
        self._refresh_combos()
        self._apply_filters()
        self.status_lbl.configure(
            text=f"✓  Synchronisé à {datetime.now().strftime('%H:%M:%S')}",
            text_color="#22C55E")

    def _refresh_combos(self):
        cats = sorted({r[0] for r in self.all_data if r and r[0]})
        fourns = sorted({r[2] for r in self.all_data if len(r) > 2 and r[2]})
        self.cat_combo.configure(values=["Toutes"] + cats)
        self.fourn_combo.configure(values=["Tous"] + fourns)

    def _apply_filters(self):
        search = self.search_var.get().lower()
        cat = self.cat_var.get()
        fourn = self.fourn_var.get()
        self.filtered_data, self.row_indices = [], []
        for i, row in enumerate(self.all_data):
            if len(row) < 4:
                continue
            p = (row + [""] * 9)[:9]
            if cat != "Toutes" and p[0] != cat:
                continue
            if fourn != "Tous" and p[2] != fourn:
                continue
            if search and search not in p[3].lower() and search not in p[1].lower():
                continue
            self.filtered_data.append(p)
            self.row_indices.append(i)
        self._refresh_tree()

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        total = 0.0
        for i, row in enumerate(self.filtered_data):
            qty = _safe_float(row[7])
            tag = "zero" if qty == 0 else ("even" if i % 2 == 0 else "odd")
            self.tree.insert("", "end", values=row, tags=(tag,), iid=str(i))
            total += _safe_float(row[8])
        self.stats_lbl.configure(
            text=f"Articles : {len(self.filtered_data)}  |  Valeur totale : {total:,.2f} €")

    def _reset_filters(self):
        self.search_var.set("")
        self.cat_var.set("Toutes")
        self.fourn_var.set("Tous")
        self._apply_filters()

    # --------------------------------------------------------- Table events --

    def _on_dbl_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        col_id = self.tree.identify_column(event.x)
        col_idx = int(col_id[1:]) - 1
        if col_idx == 7:
            self._edit_qty()
        else:
            self._edit_item()

    def _show_ctx_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.ctx_menu.post(event.x_root, event.y_root)

    def _selected(self):
        sel = self.tree.selection()
        if not sel:
            return None, None
        iid = int(sel[0])
        return iid, self.row_indices[iid]

    # ---------------------------------------------------------------- CRUD ---

    def _edit_qty(self):
        iid, orig = self._selected()
        if iid is None:
            messagebox.showinfo("Info", "Sélectionnez un article.")
            return
        row = (self.all_data[orig] + [""] * 9)[:9]
        dlg = QtyDialog(self, row[3], row[7] or "0", row[5])
        if dlg.new_qty is None:
            return
        qty = _safe_float(dlg.new_qty)
        prix = _safe_float(row[5])
        valorise = round(qty * prix, 2)
        self.all_data[orig] = list((self.all_data[orig] + [""] * 9)[:9])
        self.all_data[orig][7] = dlg.new_qty
        self.all_data[orig][8] = str(valorise)
        self._apply_filters()
        if self.sheets_mgr:
            def w():
                try:
                    self.sheets_mgr.update_qte_m31(orig, dlg.new_qty, valorise)
                    self.after(0, lambda: self.status_lbl.configure(
                        text=f"✓  Sauvegardé à {datetime.now().strftime('%H:%M:%S')}",
                        text_color="#22C55E"))
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
            threading.Thread(target=w, daemon=True).start()

    def _edit_item(self):
        iid, orig = self._selected()
        if iid is None:
            messagebox.showinfo("Info", "Sélectionnez un article.")
            return
        dlg = ItemDialog(self, self.all_data[orig])
        if dlg.result is None:
            return
        self.all_data[orig] = dlg.result
        self._refresh_combos()
        self._apply_filters()
        if self.sheets_mgr:
            def w():
                try:
                    self.sheets_mgr.batch_update_row(orig, dlg.result)
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
            threading.Thread(target=w, daemon=True).start()

    def _add_item(self):
        dlg = ItemDialog(self, None)
        if dlg.result is None:
            return
        self.all_data.append(dlg.result)
        self._refresh_combos()
        self._apply_filters()
        if self.sheets_mgr:
            def w():
                try:
                    self.sheets_mgr.add_row(dlg.result)
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
            threading.Thread(target=w, daemon=True).start()

    def _delete_item(self):
        iid, orig = self._selected()
        if iid is None:
            return
        produit = self.all_data[orig][3] if len(self.all_data[orig]) > 3 else "cet article"
        if not messagebox.askyesno("Confirmation", f"Supprimer « {produit} » ?"):
            return
        self.all_data.pop(orig)
        self._refresh_combos()
        self._apply_filters()
        if self.sheets_mgr:
            def w():
                try:
                    self.sheets_mgr.delete_row(orig)
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
            threading.Thread(target=w, daemon=True).start()

    # ------------------------------------------------------------ Exports ---

    def _export_excel(self):
        if not self.all_data:
            messagebox.showwarning("Aucune donnée", "Synchronisez d'abord.")
            return
        label = datetime.now().strftime("%Y_%m")
        try:
            path = export_to_excel(self.all_data, self.config_data.get("exports_folder", "exports"), label)
            if messagebox.askyesno("Export réussi", f"Fichier créé :\n{path}\n\nOuvrir le dossier ?"):
                os.startfile(os.path.dirname(path))
        except Exception as e:
            messagebox.showerror("Erreur d'export", str(e))

    def _close_month(self):
        if not self.all_data:
            messagebox.showwarning("Aucune donnée", "Synchronisez d'abord.")
            return
        label = datetime.now().strftime("%Y_%m")
        msg = (f"Clôture du mois {label.replace('_', '/')} :\n\n"
               "1. Export Excel de l'inventaire actuel\n"
               "2. QTE M31 → QTE M-1\n"
               "3. QTE M31 remis à zéro\n\n"
               "⚠ Cette action est IRRÉVERSIBLE. Continuer ?")
        if not messagebox.askyesno("Clôture de mois", msg):
            return
        try:
            path = export_to_excel(self.all_data, self.config_data.get("exports_folder", "exports"), label)
        except Exception as e:
            messagebox.showerror("Erreur export", str(e))
            return
        if self.sheets_mgr:
            def w():
                try:
                    self.sheets_mgr.close_month(self.all_data)
                    self.after(0, self._sync)
                    self.after(0, lambda: messagebox.showinfo(
                        "Clôture terminée",
                        f"Mois {label.replace('_', '/')} clôturé !\n\nArchive :\n{path}"))
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
            threading.Thread(target=w, daemon=True).start()
        else:
            messagebox.showinfo("Export", f"Archive créée :\n{path}")

    # ---------------------------------------------------------- Settings ---

    def _open_settings(self):
        SettingsDialog(self, self.config_data, self._on_settings_saved)

    def _on_settings_saved(self, new_cfg):
        self.config_data = new_cfg
        save_config(new_cfg)
        self.sheets_mgr = None
        self._sync()


# ===================================================================== Dialogs

class QtyDialog(ctk.CTkToplevel):
    def __init__(self, parent, produit, current_qty, prix_ht):
        super().__init__(parent)
        self.new_qty = None
        self.title("Modifier la quantité")
        self.geometry("360x210")
        self.resizable(False, False)
        self.grab_set()
        self.focus()

        ctk.CTkLabel(self, text=f"Produit : {produit}",
                     font=ctk.CTkFont(weight="bold"), wraplength=320).pack(pady=(18, 4), padx=20)
        ctk.CTkLabel(self, text=f"Prix HT : {prix_ht} €",
                     text_color="#6B7280").pack(pady=2)
        ctk.CTkLabel(self, text="Nouvelle quantité (QTE M31) :").pack(pady=(10, 2))

        self.v = tk.StringVar(value=current_qty)
        e = ctk.CTkEntry(self, textvariable=self.v, width=160,
                          justify="center", font=ctk.CTkFont(size=16))
        e.pack(pady=6)
        e.select_range(0, "end")
        e.focus()
        e.bind("<Return>", lambda _: self._save())

        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(pady=8)
        ctk.CTkButton(bf, text="Enregistrer", command=self._save, width=120).pack(side="left", padx=6)
        ctk.CTkButton(bf, text="Annuler", command=self.destroy, width=80,
                       fg_color="#6B7280", hover_color="#4B5563").pack(side="left")
        self.wait_window()

    def _save(self):
        try:
            float(self.v.get().replace(",", "."))
            self.new_qty = self.v.get().replace(",", ".")
            self.destroy()
        except ValueError:
            messagebox.showerror("Erreur", "Valeur numérique requise.", parent=self)


class ItemDialog(ctk.CTkToplevel):
    FIELDS = ["CATEGORIE", "REFERENCE", "FOURNISSEUR", "PRODUIT",
              "CONDITIONNEMENT", "PRIX HT", "QTE M-1", "QTE M31"]

    def __init__(self, parent, row_data):
        super().__init__(parent)
        self.result = None
        self.title("Modifier l'article" if row_data else "Nouvel article")
        self.geometry("460x490")
        self.resizable(False, False)
        self.grab_set()
        self.focus()

        ctk.CTkLabel(self, text="Modifier" if row_data else "Ajouter un article",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(16, 8))

        scroll = ctk.CTkScrollableFrame(self, height=370)
        scroll.pack(fill="both", expand=True, padx=16)

        self.vars = {}
        for field in self.FIELDS:
            fr = ctk.CTkFrame(scroll, fg_color="transparent")
            fr.pack(fill="x", pady=4)
            ctk.CTkLabel(fr, text=field + " :", width=140, anchor="w").pack(side="left")
            v = tk.StringVar()
            if row_data:
                idx = self.FIELDS.index(field)
                v.set(row_data[idx] if idx < len(row_data) else "")
            ctk.CTkEntry(fr, textvariable=v, width=270).pack(side="left")
            self.vars[field] = v

        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(pady=10)
        ctk.CTkButton(bf, text="Enregistrer", command=self._save, width=130).pack(side="left", padx=6)
        ctk.CTkButton(bf, text="Annuler", command=self.destroy, width=80,
                       fg_color="#6B7280", hover_color="#4B5563").pack(side="left")
        self.wait_window()

    def _save(self):
        vals = [self.vars[f].get() for f in self.FIELDS]
        try:
            qty = float(vals[7].replace(",", ".")) if vals[7] else 0
            prix = float(vals[5].replace(",", ".")) if vals[5] else 0
            valorise = str(round(qty * prix, 2))
        except (ValueError, TypeError):
            valorise = ""
        self.result = vals + [valorise]
        self.destroy()


class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, config, on_save):
        super().__init__(parent)
        self.on_save = on_save
        self.cfg = config.copy()
        self.title("Paramètres")
        self.geometry("520x380")
        self.resizable(False, False)
        self.grab_set()
        self.focus()

        ctk.CTkLabel(self, text="Configuration Google Sheets",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(16, 10))

        fr = ctk.CTkFrame(self, fg_color="transparent")
        fr.pack(fill="both", expand=True, padx=22)

        ctk.CTkLabel(fr, text="ID du Google Sheet :", anchor="w").pack(fill="x", pady=(8, 2))
        self.id_v = tk.StringVar(value=config.get("spreadsheet_id", ""))
        ctk.CTkEntry(fr, textvariable=self.id_v, width=460).pack(fill="x")
        ctk.CTkLabel(fr,
                     text="L'ID est dans l'URL : docs.google.com/spreadsheets/d/[ID]/edit",
                     text_color="#6B7280", font=ctk.CTkFont(size=10)).pack(anchor="w")

        ctk.CTkLabel(fr, text="Nom de la feuille :", anchor="w").pack(fill="x", pady=(10, 2))
        self.name_v = tk.StringVar(value=config.get("sheet_name", "Feuil1"))
        ctk.CTkEntry(fr, textvariable=self.name_v, width=200).pack(anchor="w")

        ctk.CTkLabel(fr, text="Dossier d'exports :", anchor="w").pack(fill="x", pady=(10, 2))
        self.exp_v = tk.StringVar(value=config.get("exports_folder", "exports"))
        ef = ctk.CTkFrame(fr, fg_color="transparent")
        ef.pack(fill="x")
        ctk.CTkEntry(ef, textvariable=self.exp_v, width=380).pack(side="left")
        ctk.CTkButton(ef, text="…", width=44, command=self._browse).pack(side="left", padx=6)

        info = ctk.CTkFrame(fr, fg_color="#FEF3C7", corner_radius=8)
        info.pack(fill="x", pady=(14, 0))
        ctk.CTkLabel(info,
                     text="⚠  Placez credentials.json dans le dossier de l'application.\n"
                          "Consultez SETUP.txt pour les instructions détaillées.",
                     text_color="#92400E", font=ctk.CTkFont(size=10),
                     justify="left").pack(padx=12, pady=8)

        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(pady=14)
        ctk.CTkButton(bf, text="Enregistrer", command=self._save, width=130).pack(side="left", padx=6)
        ctk.CTkButton(bf, text="Annuler", command=self.destroy, width=80,
                       fg_color="#6B7280", hover_color="#4B5563").pack(side="left")
        self.wait_window()

    def _browse(self):
        folder = filedialog.askdirectory()
        if folder:
            self.exp_v.set(folder)

    def _save(self):
        self.cfg["spreadsheet_id"] = self.id_v.get().strip()
        self.cfg["sheet_name"] = self.name_v.get().strip()
        self.cfg["exports_folder"] = self.exp_v.get().strip()
        self.on_save(self.cfg)
        self.destroy()


if __name__ == "__main__":
    app = InventaireApp()
    app.mainloop()
