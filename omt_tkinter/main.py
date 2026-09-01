"""OMT Offline — run me:   python main.py

Load your Item Master (and Ship-To, if you want it), pick a channel, set the
margin, choose the PO file, press Generate. Download writes the sheet.

Two files in total: this one is the window, core.py is everything else.
"""
from __future__ import annotations

import json
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import core

# Sober palette — flat greys, one green accent (the same green as the sheet).
BG, CARD, LINE = '#f4f4f5', '#ffffff', '#d8d8dc'
TEXT, MUTED = '#1c1c1e', '#6c6c72'
BAD, GOOD = '#b3261e', '#1b5e20'

DATA_TYPES = [('Excel / CSV', '*.xlsx *.xls *.csv'), ('All files', '*.*')]
PO_TYPES = [('PO files', '*.csv *.xlsx *.xls *.xlsb'), ('All files', '*.*')]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('OMT Offline')
        self.geometry('1040x680')
        self.minsize(880, 580)
        self.configure(bg=BG)
        try:
            self.iconbitmap(core.resource_path('assets', 'renee.ico'))
        except Exception:                                   # noqa: BLE001 — cosmetic
            pass

        self.v_items = tk.StringVar()
        self.v_shipto = tk.StringVar()
        self.v_skumap = tk.StringVar()
        self.v_po = tk.StringVar()
        self.v_channel = tk.StringVar(value=core.CHANNELS[0])
        self.v_margin = tk.StringVar(value=str(core.DEFAULT_MARGINS[core.CHANNELS[0]]))
        self.v_status = tk.StringVar(value='Load your Item Master to begin.')
        self.items = None
        self.shipto = None
        self.skumap = None
        self.result = None

        self.v_find = tk.StringVar()
        self.v_only_issues = tk.BooleanVar(value=False)
        self._sort_key, self._sort_desc = None, False

        self._style()
        self._menu()
        self._build()
        self._restore()
        self._shortcuts()

    # ── menu ──────────────────────────────────────────────────────────────
    def _menu(self):
        m = tk.Menu(self)

        f = tk.Menu(m, tearoff=0)
        f.add_command(label='Open PO file…\tCtrl+O', command=self._browse_po)
        f.add_command(label='Generate\tF5', command=self._generate)
        f.add_command(label='Download sheet…\tCtrl+S', command=self._download)
        f.add_separator()
        f.add_command(label='Exit', command=self.destroy)
        m.add_cascade(label='File', menu=f)

        r = tk.Menu(m, tearoff=0)
        r.add_command(label='Load Item Master…', command=self._browse_items)
        r.add_command(label='Load Ship-To Mapping…', command=self._browse_shipto)
        r.add_command(label='Load Swiggy SKU Map…', command=self._browse_skumap)
        m.add_cascade(label='Reference', menu=r)

        t = tk.Menu(m, tearoff=0)
        t.add_command(label='Item Master template', command=self._tpl_items)
        t.add_command(label='Ship-To Mapping template', command=self._tpl_shipto)
        t.add_command(label='Swiggy SKU Map template', command=self._tpl_skumap)
        t.add_separator()
        for ch in core.CHANNELS:                     # standard PO file per channel
            t.add_command(label=f'{ch} PO file template',
                          command=lambda c=ch: self._tpl_po(c))
        m.add_cascade(label='Templates', menu=t)

        h = tk.Menu(m, tearoff=0)
        h.add_command(label='About', command=self._about)
        m.add_cascade(label='Help', menu=h)
        self.config(menu=m)

    def _shortcuts(self):
        self.bind('<Control-o>', lambda e: self._browse_po())
        self.bind('<Control-s>', lambda e: self._download())
        self.bind('<F5>', lambda e: self._generate())
        self.bind('<Control-f>', lambda e: self.ent_find.focus_set())

    def _about(self):
        messagebox.showinfo(
            'OMT Offline',
            'OMT Offline\n\n'
            'PO file → Item No · MRP · Landing · GST Code · Cost Price · Diffn\n'
            'Blinkit · RK · Swiggy · GT Mass\n\n'
            'Landing = MRP × margin%\n'
            'Cost Price = Landing ÷ (1 + GST)\n\n'
            'Runs entirely on this machine — no database, no internet.')

    def _style(self):
        s = ttk.Style(self)
        try:
            s.theme_use('clam')
        except tk.TclError:
            pass
        s.configure('.', background=BG, foreground=TEXT, font=('Segoe UI', 10))
        s.configure('TLabel', background=CARD, foreground=TEXT)
        s.configure('M.TLabel', background=CARD, foreground=MUTED, font=('Segoe UI', 9))
        s.configure('H.TLabel', background=BG, foreground=TEXT, font=('Segoe UI', 15, 'bold'))
        s.configure('S.TLabel', background=BG, foreground=MUTED, font=('Segoe UI', 9))
        s.configure('Sec.TLabel', background=CARD, foreground=MUTED, font=('Segoe UI', 8, 'bold'))
        s.configure('TButton', padding=(11, 6))
        s.configure('Go.TButton', padding=(18, 8), font=('Segoe UI', 10, 'bold'))
        s.configure('Treeview', rowheight=25, fieldbackground=CARD, background=CARD)
        s.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))

    def _card(self):
        f = tk.Frame(self, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        f.pack(fill='x', padx=16, pady=(0, 8))
        g = tk.Frame(f, bg=CARD)
        g.pack(fill='x', padx=14, pady=12)
        return g

    def _data_row(self, g, label, var, browse, template, r):
        """label [path.....] [Browse] [Template]  + a status line under it."""
        ttk.Label(g, text=label).grid(row=r, column=0, sticky='w', pady=4)
        tk.Entry(g, textvariable=var, relief='solid', bd=1).grid(
            row=r, column=1, sticky='we', padx=(10, 6), ipady=3)
        ttk.Button(g, text='Browse', command=browse).grid(row=r, column=2)
        ttk.Button(g, text='Template', command=template).grid(row=r, column=3, padx=(6, 0))
        lbl = ttk.Label(g, text='', style='M.TLabel')
        lbl.grid(row=r + 1, column=1, columnspan=3, sticky='w', padx=(10, 0))
        g.columnconfigure(1, weight=1)
        return lbl

    def _build(self):
        head = tk.Frame(self, bg=BG)
        head.pack(fill='x', padx=16, pady=(14, 8))
        ttk.Label(head, text='OMT Offline', style='H.TLabel').pack(anchor='w')
        ttk.Label(head, text='PO file → Item No · MRP · Landing · GST Code · Cost Price · Diffn',
                  style='S.TLabel').pack(anchor='w')

        # ── reference data ──
        g1 = self._card()
        ttk.Label(g1, text='REFERENCE DATA', style='Sec.TLabel')\
            .grid(row=0, column=0, columnspan=4, sticky='w', pady=(0, 6))
        self.lbl_items = self._data_row(g1, 'Item Master', self.v_items,
                                        self._browse_items, self._tpl_items, 1)
        self.lbl_shipto = self._data_row(g1, 'Ship-To Mapping', self.v_shipto,
                                         self._browse_shipto, self._tpl_shipto, 3)
        self.lbl_skumap = self._data_row(g1, 'Swiggy SKU Map', self.v_skumap,
                                         self._browse_skumap, self._tpl_skumap, 5)
        ttk.Label(g1, text='Loaded once and kept — only Browse again when you want to update them.',
                  style='M.TLabel').grid(row=7, column=1, columnspan=3, sticky='w',
                                         padx=(10, 0), pady=(6, 0))

        # ── run ──
        g2 = self._card()
        ttk.Label(g2, text='RUN', style='Sec.TLabel')\
            .grid(row=0, column=0, columnspan=4, sticky='w', pady=(0, 6))
        ttk.Label(g2, text='Channel').grid(row=1, column=0, sticky='w')
        cb = ttk.Combobox(g2, textvariable=self.v_channel, values=core.CHANNELS,
                          state='readonly', width=13)
        cb.grid(row=2, column=0, sticky='w', pady=(3, 0))
        cb.bind('<<ComboboxSelected>>', self._on_channel)

        ttk.Label(g2, text='Margin').grid(row=1, column=1, sticky='w', padx=(16, 0))
        mf = tk.Frame(g2, bg=CARD)
        mf.grid(row=2, column=1, sticky='w', padx=(16, 0), pady=(3, 0))
        tk.Entry(mf, textvariable=self.v_margin, width=6, justify='right',
                 relief='solid', bd=1).pack(side='left', ipady=3)
        ttk.Label(mf, text='%').pack(side='left', padx=(4, 0))

        ttk.Label(g2, text='PO file').grid(row=1, column=2, sticky='w', padx=(16, 0))
        ff = tk.Frame(g2, bg=CARD)
        ff.grid(row=2, column=2, sticky='we', padx=(16, 0), pady=(3, 0))
        tk.Entry(ff, textvariable=self.v_po, relief='solid', bd=1).pack(
            side='left', fill='x', expand=True, ipady=3)
        ttk.Button(ff, text='Browse', command=self._browse_po).pack(side='left', padx=(6, 0))
        g2.columnconfigure(2, weight=1)
        ttk.Button(g2, text='Generate', style='Go.TButton', command=self._generate)\
            .grid(row=2, column=3, sticky='e', padx=(16, 0), pady=(3, 0))

        # ── results: filter bar + table ──
        outer = tk.Frame(self, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        outer.pack(fill='both', expand=True, padx=16, pady=(0, 8))

        bar = tk.Frame(outer, bg=CARD)
        bar.pack(fill='x', padx=12, pady=(10, 6))
        ttk.Label(bar, text='Find').pack(side='left')
        self.ent_find = tk.Entry(bar, textvariable=self.v_find, relief='solid',
                                 bd=1, width=30)
        self.ent_find.pack(side='left', padx=(8, 0), ipady=2)
        self.v_find.trace_add('write', lambda *_: self._refill())
        ttk.Checkbutton(bar, text='Only flagged', variable=self.v_only_issues,
                        command=self._refill).pack(side='left', padx=(14, 0))
        self.lbl_shown = ttk.Label(bar, text='', style='M.TLabel')
        self.lbl_shown.pack(side='right')

        wrap = tk.Frame(outer, bg=CARD)
        wrap.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(wrap, columns=core._KEYS, show='headings')
        widths = {'ean': 125, 'item_no': 115, 'mrp': 80, 'lr': 95, 'gst_code': 85,
                  'cp': 90, 'diffn': 105, 'issue': 175}
        for key, heading in core.columns(70):
            # Click a heading to sort by that column.
            self.tree.heading(key, text=heading,
                              command=lambda k=key: self._sort_by(k))
            self.tree.column(key, width=widths.get(key, 100),
                             anchor='e' if key in ('mrp', 'lr', 'cp', 'diffn') else 'w')
        self.tree.tag_configure('bad', background='#fdeaea', foreground=BAD)
        vs = ttk.Scrollbar(wrap, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        # Copy the selected row (Ctrl+C or right-click).
        self.tree.bind('<Control-c>', lambda e: self._copy_row())
        self.tree.bind('<Button-3>', self._row_menu)
        self._rowmenu = tk.Menu(self, tearoff=0)
        self._rowmenu.add_command(label='Copy row', command=self._copy_row)

        # ── footer ──
        foot = tk.Frame(self, bg=BG)
        foot.pack(fill='x', padx=16, pady=(0, 14))
        self.lbl_status = tk.Label(foot, textvariable=self.v_status, bg=BG, fg=MUTED,
                                   font=('Segoe UI', 9), anchor='w')
        self.lbl_status.pack(side='left')
        self.btn_dl = ttk.Button(foot, text='Download sheet', command=self._download,
                                 state='disabled')
        self.btn_dl.pack(side='right')

    # ── remembered masters ─────────────────────────────────────────────────
    # Load once, not every run: the app keeps its own copy of both files and
    # reloads them at startup. Browse is only for updating them.
    def _restore(self):
        try:
            saved = json.loads(core.settings_file().read_text(encoding='utf-8'))
        except (OSError, ValueError):
            saved = {}
        self._cfg = saved
        for kind, var, loader in (('items', self.v_items, self._load_items),
                                  ('shipto', self.v_shipto, self._load_shipto),
                                  ('skumap', self.v_skumap, self._load_skumap)):
            e = saved.get(kind) or {}
            if isinstance(e, str):                 # older settings format
                e = {'source': e, 'cached': e}
            cached, source = e.get('cached', ''), e.get('source', '')
            if cached and Path(cached).exists():
                var.set(source or cached)
                loader(cached, quiet=True, remember=False)

    def _save(self, kind, source, cached):
        cfg = getattr(self, '_cfg', {}) or {}
        cfg[kind] = {'source': str(source), 'cached': str(cached)}
        self._cfg = cfg
        try:
            core.settings_file().write_text(json.dumps(cfg, indent=2), encoding='utf-8')
        except OSError:
            pass                                   # remembering is a nicety

    # ── the two reference files ────────────────────────────────────────────
    def _browse_items(self):
        p = filedialog.askopenfilename(title='Item Master file', filetypes=DATA_TYPES)
        if p:
            self.v_items.set(p)
            self._load_items(p)

    def _load_items(self, path, quiet=False, remember=True):
        try:
            self.items = core.load_items(path)
        except Exception as exc:                            # noqa: BLE001
            self.items = None
            self.lbl_items.configure(text='Could not read this file.', foreground=BAD)
            if not quiet:
                messagebox.showerror('Item Master', str(exc))
            return
        kept = core.cache_reference(path, 'item_master') if remember else path
        if remember:
            self._save('items', path, kept)
        stamp = core.file_stamp(kept)
        self.lbl_items.configure(
            text=f'{len(self.items):,} items' + (f'  ·  last updated {stamp}' if stamp else ''),
            foreground=GOOD)
        self.v_status.set('Ready — choose a PO file.')

    def _browse_shipto(self):
        p = filedialog.askopenfilename(title='Ship-To Mapping file', filetypes=DATA_TYPES)
        if p:
            self.v_shipto.set(p)
            self._load_shipto(p)

    def _load_shipto(self, path, quiet=False, remember=True):
        try:
            self.shipto = st = core.load_shipto(path)
        except Exception as exc:                            # noqa: BLE001
            self.shipto = None
            self.lbl_shipto.configure(text='Could not read this file.', foreground=BAD)
            if not quiet:
                messagebox.showerror('Ship-To Mapping', str(exc))
            return
        kept = core.cache_reference(path, 'ship_to') if remember else path
        if remember:
            self._save('shipto', path, kept)
        n_p = len(st.parties())
        stamp = core.file_stamp(kept)
        self.lbl_shipto.configure(
            text=f'{len(st):,} locations' + (f' · {n_p} parties' if n_p else '') +
                 (f'  ·  last updated {stamp}' if stamp else ''),
            foreground=GOOD)

    def _browse_skumap(self):
        p = filedialog.askopenfilename(title='Swiggy SKU Map file', filetypes=DATA_TYPES)
        if p:
            self.v_skumap.set(p)
            self._load_skumap(p)

    def _load_skumap(self, path, quiet=False, remember=True):
        try:
            self.skumap = core.load_sku_map(path)
        except Exception as exc:                            # noqa: BLE001
            self.skumap = None
            self.lbl_skumap.configure(text='Could not read this file.', foreground=BAD)
            if not quiet:
                messagebox.showerror('Swiggy SKU Map', str(exc))
            return
        kept = core.cache_reference(path, 'sku_map') if remember else path
        if remember:
            self._save('skumap', path, kept)
        stamp = core.file_stamp(kept)
        self.lbl_skumap.configure(
            text=f'{len(self.skumap):,} SKU codes' +
                 (f'  ·  last updated {stamp}' if stamp else ''),
            foreground=GOOD)

    def _tpl_skumap(self):
        self._save_template('Swiggy SKU Map', 'swiggy_sku_map_template.xlsx',
                            core.write_sku_map_template)

    def _tpl_items(self):
        self._save_template('Item Master', 'item_master_template.xlsx',
                            core.write_item_master_template)

    def _tpl_shipto(self):
        self._save_template('Ship-To Mapping', 'ship_to_template.xlsx',
                            core.write_shipto_template)

    def _save_template(self, label, default_name, writer):
        p = filedialog.asksaveasfilename(
            title=f'Save the {label} template', defaultextension='.xlsx',
            initialfile=default_name, filetypes=[('Excel', '*.xlsx')])
        if not p:
            return
        try:
            writer(p)
        except Exception as exc:                            # noqa: BLE001
            messagebox.showerror('Template', str(exc))
            return
        messagebox.showinfo('Template saved',
                            f'{label} template saved.\n\nFill it in, then load it '
                            f'with Browse.\n\n{p}')

    # ── the run ────────────────────────────────────────────────────────────
    def _on_channel(self, _e=None):
        self.v_margin.set(str(core.DEFAULT_MARGINS.get(self.v_channel.get(), 70)))

    def _browse_po(self):
        p = filedialog.askopenfilename(title='PO file', filetypes=PO_TYPES)
        if p:
            self.v_po.set(p)

    def _generate(self):
        if not self.items or not len(self.items):
            messagebox.showwarning('Item Master', 'Load your Item Master first.')
            return
        path = self.v_po.get().strip()
        if not path or not Path(path).exists():
            messagebox.showwarning('PO file', 'Choose a PO file first.')
            return
        try:
            margin = float(self.v_margin.get())
            if margin <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning('Margin', 'Margin must be a number above 0.')
            return
        self.v_status.set('Working…')
        self.lbl_status.configure(fg=MUTED)
        self.btn_dl.configure(state='disabled')
        self.update_idletasks()
        threading.Thread(target=self._work, args=(path, self.v_channel.get(), margin),
                         daemon=True).start()

    def _work(self, path, channel, margin):
        try:
            res = core.process(path, channel, margin, self.items, self.skumap)
        except Exception as exc:                            # noqa: BLE001
            self.after(0, self._failed, str(exc))
            return
        self.after(0, self._done, res, margin)

    def _failed(self, msg):
        self.v_status.set('Failed.')
        self.lbl_status.configure(fg=BAD)
        messagebox.showerror('Could not read the PO file', msg)

    def _done(self, res, margin):
        self.result = res
        self._margin_used = margin
        for key, heading in core.columns(margin):           # Landing (m%) follows the box
            self.tree.heading(key, text=heading,
                              command=lambda k=key: self._sort_by(k))
        self._sort_key = None
        self._refill()
        c = res['counts']
        self.v_status.set(f"{c['total']} lines · {c['priced']} clean · {c['issues']} flagged")
        self.lbl_status.configure(fg=BAD if c['issues'] else GOOD)
        self.btn_dl.configure(state='normal' if res['rows'] else 'disabled')
        if res['warnings']:
            messagebox.showinfo('Notes', '\n'.join(res['warnings']))

    # ── table: filter / sort / copy ────────────────────────────────────────
    def _visible_rows(self):
        rows = (self.result or {}).get('rows', [])
        if self.v_only_issues.get():
            rows = [r for r in rows if r.get('issue')]
        q = self.v_find.get().strip().lower()
        if q:
            rows = [r for r in rows
                    if any(q in str(r.get(k, '')).lower() for k in core._KEYS)]
        if self._sort_key:
            def key(r):
                v = r.get(self._sort_key)
                if isinstance(v, (int, float)):
                    return (0, v, '')
                return (1, 0, str(v or '').lower())         # numbers before text
            rows = sorted(rows, key=key, reverse=self._sort_desc)
        return rows

    def _refill(self):
        if not self.result:
            return
        margin = getattr(self, '_margin_used', 70)
        self.tree.delete(*self.tree.get_children())
        rows = self._visible_rows()
        for r in rows:
            vals = [('' if r.get(k) is None else
                     (f"{r[k]:,.2f}" if k in ('mrp', 'lr', 'cp', 'diffn') else r[k]))
                    for k, _ in core.columns(margin)]
            self.tree.insert('', 'end', values=vals, tags=('bad',) if r['issue'] else ())
        total = len(self.result['rows'])
        self.lbl_shown.configure(
            text=f'showing {len(rows)} of {total}' if len(rows) != total else '')

    def _sort_by(self, key):
        self._sort_desc = not self._sort_desc if self._sort_key == key else False
        self._sort_key = key
        self._refill()

    def _row_menu(self, ev):
        iid = self.tree.identify_row(ev.y)
        if iid:
            self.tree.selection_set(iid)
            self._rowmenu.tk_popup(ev.x_root, ev.y_root)

    def _copy_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        self.clipboard_clear()
        self.clipboard_append('\t'.join(
            str(v) for v in self.tree.item(sel[0])['values']))

    def _tpl_po(self, channel):
        self._save_template(f'{channel} PO file',
                            f"{channel.replace(' ', '_').lower()}_po_template.xlsx",
                            lambda p: core.write_po_template(channel, p))

    def _download(self):
        if not self.result or not self.result['rows']:
            return
        margin = float(self.v_margin.get())
        p = filedialog.asksaveasfilename(
            title='Save the sheet', defaultextension='.xlsx',
            initialfile=core.default_name(self.v_channel.get(), margin),
            filetypes=[('Excel', '*.xlsx')])
        if not p:
            return
        try:
            out = core.build_workbook(
                self.result['rows'], self.result['lines'], self.v_channel.get(),
                margin, p, shipto=self.shipto,
                warnings=self.result.get('warnings', ()))
        except Exception as exc:                            # noqa: BLE001
            messagebox.showerror('Download failed', str(exc))
            return
        self.v_status.set(f'Saved: {out.name}')
        messagebox.showinfo('Saved', f'7-sheet workbook saved.\n\n{out}')


if __name__ == '__main__':
    App().mainloop()
