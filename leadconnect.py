"""LeadConnect Pro desktop application."""
import csv, re, sys, webbrowser, traceback
from pathlib import Path
from urllib.parse import urlencode
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from contacts import Store, STATUSES, eligible, normalize, read_records, safe_csv

TEMPLATE = ('مساء الخير يا فندم، مع حضرتك [اسمك]، مؤسس International Software Studio.\n\n'
            'حابب أعرض على إدارة {business} نظام LoungeLink POS، اللي طورناه خصيصاً لإدارة محلات البلايستيشن والجيمينج. '
            'السيستم بيحسب وقت اللعب وتسعير الغرف، ويدير المخزون والمصروفات والورديات من غير اشتراك شهري إجباري.\n\n'
            'لو مناسب لحضرتك، ممكن نحدد وقت لعرض ديمو مجاني بدون التزام بالشراء؟')

class App:
    def __init__(self,root,store):
        self.root=root; self.s=store; self.selected=None; self.opened=None
        root.title('LeadConnect Pro | Local Sales Workspace'); root.geometry('1350x850'); root.minsize(1000,650); root.configure(bg='#101827')
        style=ttk.Style(root); style.theme_use('clam')
        style.configure('.',background='#101827',foreground='#e2e8f0',font=('Segoe UI',10))
        style.configure('TFrame',background='#101827')
        style.configure('TLabel',background='#101827',foreground='#e2e8f0')
        style.configure('TButton',background='#24364a',foreground='#ffffff',padding=7)
        style.map('TButton',background=[('active','#34506b'),('disabled','#192435')])
        style.configure('Accent.TButton',background='#087e78')
        style.map('Accent.TButton',background=[('active','#0b9b91'),('disabled','#192435')])
        style.configure('Treeview',background='#172338',fieldbackground='#172338',foreground='#e2e8f0',rowheight=29)
        style.configure('Treeview.Heading',background='#24364a',foreground='white')
        style.map('Treeview',background=[('selected','#126a73')])
        style.configure('TNotebook',background='#101827')
        style.configure('TNotebook.Tab',padding=(16,9),background='#24364a',foreground='white')
        style.map('TNotebook.Tab',background=[('selected','#087e78')])
        style.configure('TEntry',fieldbackground='#eaf0f6',foreground='#142337')
        style.configure('TCombobox',fieldbackground='#eaf0f6',foreground='#142337')
        style.map('TCombobox',fieldbackground=[('readonly','#eaf0f6')],foreground=[('readonly','#142337')])
        header=ttk.Frame(root,padding=18);header.pack(fill='x')
        ttk.Label(header,text='LeadConnect Pro',font=('Segoe UI',24,'bold')).pack(anchor='w')
        ttk.Label(header,text='LOCAL SALES WORKSPACE  /  EGYPT',foreground='#8fcac6').pack(anchor='w')
        self.kpi=ttk.Label(header,text='',font=('Segoe UI',12,'bold'));self.kpi.pack(anchor='w',pady=(12,0))
        tabs=ttk.Notebook(root);tabs.pack(fill='both',expand=True,padx=18,pady=(0,12))
        self.contacts=ttk.Frame(tabs,padding=12); self.replies=ttk.Frame(tabs,padding=20)
        tabs.add(self.contacts,text='01  Prospects & outreach');tabs.add(self.replies,text='02  Reply matching')
        tools=ttk.Frame(self.contacts);tools.pack(fill='x')
        ttk.Button(tools,text='Import CSV / Excel / JSON',command=self.import_file,style='Accent.TButton').pack(side='left')
        ttk.Button(tools,text='Paste phone numbers',command=self.paste).pack(side='left',padx=5)
        ttk.Button(tools,text='Export tracker CSV',command=self.export).pack(side='left')
        self.filter=tk.StringVar(value='Egypt mobile')
        ttk.Combobox(tools,textvariable=self.filter,values=['All','Egypt mobile','Landline / other Egypt','Invalid','International / review','Ready for review','Awaiting reply','Replied','Do not contact'],state='readonly',width=24).pack(side='right')
        self.search=tk.StringVar()
        ttk.Entry(tools,textvariable=self.search,width=23).pack(side='right',padx=8)
        ttk.Label(tools,text='Search').pack(side='right')
        pane=ttk.Panedwindow(self.contacts,orient='horizontal');pane.pack(fill='both',expand=True,pady=12)
        left=ttk.Frame(pane);right=ttk.Frame(pane,padding=12);pane.add(left,weight=3);pane.add(right,weight=2)
        cols=('name','phone','kind','status','verification')
        self.tree=ttk.Treeview(left,columns=cols,show='headings',selectmode='browse')
        for c,w in zip(cols,(190,125,140,105,140)):
            self.tree.heading(c,text=c.title());self.tree.column(c,width=w,minwidth=75,stretch=True)
        scroll=ttk.Scrollbar(left,orient='vertical',command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set);scroll.pack(side='right',fill='y');self.tree.pack(fill='both',expand=True)
        self.tree.bind('<<TreeviewSelect>>',self.select)
        self.title=ttk.Label(right,text='Select a prospect',font=('Segoe UI',15,'bold'),wraplength=420);self.title.pack(anchor='w')
        self.info=ttk.Label(right,text='Import MapHarvest results to begin.',wraplength=420);self.info.pack(anchor='w',pady=6)
        self.status=tk.StringVar(value='New');self.verify=tk.StringVar(value='Unknown');self.consent=tk.StringVar();self.notes=tk.StringVar()
        for label,var,values in [('Sales status',self.status,STATUSES),('WhatsApp account — manual check',self.verify,('Unknown','Confirmed manually','Unavailable'))]:
            ttk.Label(right,text=label).pack(anchor='w',pady=(6,2));ttk.Combobox(right,textvariable=var,values=values,state='readonly').pack(fill='x')
        for label,var in [('Permission evidence (how and when granted)',self.consent),('Sales notes',self.notes)]:
            ttk.Label(right,text=label).pack(anchor='w',pady=(6,2));ttk.Entry(right,textvariable=var).pack(fill='x')
        ttk.Button(right,text='Save prospect details',command=self.save).pack(anchor='w',pady=7)
        ttk.Label(right,text='Message template  •  {business} / {area}').pack(anchor='w')
        self.template=tk.Text(right,height=7,wrap='word',bg='#eaf0f6',fg='#142337',font=('Segoe UI',10));self.template.pack(fill='x');self.template.insert('1.0',TEMPLATE)
        ttk.Button(right,text='Review personalized message & open WhatsApp',command=self.open_chat,style='Accent.TButton').pack(fill='x',pady=(10,3))
        ttk.Button(right,text='Confirm I sent the message manually',command=self.mark_sent).pack(fill='x',pady=3)
        ttk.Label(right,text='WhatsApp opens one chat. You press Send yourself. Record permission first. No bulk sending, account probing or automatic delivery claims.',wraplength=420,foreground='#94a3b8').pack(anchor='w',pady=8)
        ttk.Label(self.replies,text='Match a received message to an outreach contact',font=('Segoe UI',17,'bold')).pack(anchor='w')
        ttk.Label(self.replies,text='Enter the sender number from a reply you received in WhatsApp. Only contacts with recorded outbound messages match. Unmatched conversations are not stored.',wraplength=850).pack(anchor='w',pady=10)
        self.sender=tk.StringVar();ttk.Entry(self.replies,textvariable=self.sender,width=35).pack(anchor='w')
        self.replynote=tk.Text(self.replies,height=5,width=80,wrap='word',bg='#eaf0f6',fg='#142337');self.replynote.pack(anchor='w',pady=10)
        ttk.Button(self.replies,text='Match sender & log reply',command=self.reply,style='Accent.TButton').pack(anchor='w')
        self.replyresult=ttk.Label(self.replies,text='',wraplength=850,font=('Segoe UI',12,'bold'));self.replyresult.pack(anchor='w',pady=15)
        ttk.Label(self.replies,text='Automatic inbox monitoring is not enabled. Use WhatsApp notifications; this screen records replies you explicitly enter.',wraplength=850,foreground='#94a3b8').pack(anchor='w')
        self.footer=ttk.Label(root,text='Local storage only.  |  Public listing ≠ WhatsApp permission.',foreground='#8fcac6');self.footer.pack(anchor='w',padx=20,pady=(0,10))
        self.filter.trace_add('write',lambda *_:self.refresh());self.search.trace_add('write',lambda *_:self.refresh())
        self.refresh()
    def refresh(self):
        rows=self.s.all()
        self.kpi.config(text=f"{len(rows)} imported  |  {sum(r['kind']=='Egypt mobile' for r in rows)} mobile  |  {sum(r['kind']!='Egypt mobile' for r in rows)} excluded / review  |  {sum(eligible(r) and not r['sent'] for r in rows)} ready  |  {sum(bool(r['sent']) for r in rows)} sent (recorded)  |  {sum(bool(r['replied']) for r in rows)} replies")
        self.tree.delete(*self.tree.get_children());mode=self.filter.get();q=self.search.get().strip().lower()
        for r in rows:
            if q and q not in (r['name']+' '+r['phone']+' '+r['raw']+' '+r['area']).lower():continue
            if mode=='Ready for review' and not (eligible(r) and not r['sent']):continue
            if mode=='Awaiting reply' and not (r['sent'] and not r['replied']):continue
            if mode=='Replied' and not r['replied']:continue
            if mode=='Do not contact' and r['status'] not in {'Declined','Do not contact'}:continue
            if mode in ('Egypt mobile','Landline / other Egypt','Invalid','International / review') and mode!=r['kind']:continue
            self.tree.insert('', 'end',iid=str(r['id']),values=(r['name'],r['phone'] or r['raw'],r['kind'],r['status'],r['verification']))
        if self.selected and self.tree.exists(str(self.selected)):self.tree.selection_set(str(self.selected))
        else:
            self.selected=None;self.title.config(text='Select a prospect');self.info.config(text='Select a visible contact to review its details.')
            self.status.set('New');self.verify.set('Unknown');self.consent.set('');self.notes.set('')
    def select(self,_=None):
        selection=self.tree.selection()
        if not selection:return
        self.selected=int(selection[0]);r=self.s.get(self.selected)
        self.title.config(text=r['name'])
        self.info.config(text=f"{r['phone'] or r['raw']}  •  {r['area']}\n{r['kind']} | Last sent: {r['sent'] or 'not recorded'}\nReply: {r['replied'] or 'not recorded'}")
        self.status.set(r['status']);self.verify.set(r['verification']);self.consent.set(r['consent']);self.notes.set(r['notes'])
    def save(self):
        if not self.selected:return
        self.s.update(self.selected,self.status.get(),self.verify.get(),self.consent.get(),self.notes.get());self.refresh();self.footer.config(text='Prospect saved locally.')
    def import_file(self):
        path=filedialog.askopenfilename(filetypes=[('Supported files','*.csv *.xlsx *.json *.tsv *.txt'),('All files','*.*')])
        if not path:return
        try:self.report_import(self.s.import_rows(read_records(path)))
        except Exception as e:messagebox.showerror('Import failed',str(e))
    def report_import(self,counts):
        a,d,i=counts;self.refresh()
        messagebox.showinfo('Import complete',f'{a} new records, {d} duplicates skipped, {i} invalid number entries. Existing statuses and opt-outs preserved.')
        self.footer.config(text=f'Imported {a} new records; {d} duplicates skipped.')
    def paste(self):
        win=tk.Toplevel(self.root);win.title('Paste numbers — one per line');win.geometry('500x350')
        box=tk.Text(win,wrap='word');box.pack(fill='both',expand=True,padx=10,pady=10)
        def done():
            self.report_import(self.s.import_rows([{'phone':n} for n in re.split(r'[\n;,]+',box.get('1.0','end').strip()) if n.strip()]));win.destroy()
        ttk.Button(win,text='Import numbers',command=done).pack(pady=5)
    def open_chat(self):
        if not self.selected:return
        contact_id=self.selected;self.save();r=self.s.get(contact_id)
        if not eligible(r):
            messagebox.showwarning('Review required','A valid Egyptian mobile number, recorded permission and a non-blocked status are required.');return
        if r['sent'] and not messagebox.askyesno('Previously contacted','This contact already has a recorded outbound message. Continue only if a relevant follow-up was agreed?'):return
        text=self.template.get('1.0','end').strip().replace('{business}',r['name']).replace('{area}',r['area'])
        if not text or len(text)>1500 or '[اسمك]' in text:
            messagebox.showwarning('Message','Enter your name and a non-empty message up to 1,500 characters.');return
        if messagebox.askyesno('Review recipient',f"{r['name']}\n+{r['phone']}\n\n{text}\n\nOpen this individual chat? Nothing will be sent automatically."):
            if webbrowser.open('https://wa.me/'+r['phone']+'?'+urlencode({'text':text})):
                self.opened=(r['id'],text);self.footer.config(text='Chat opened; delivery not verified. Confirm only after sending.')
            else:messagebox.showerror('Browser','Could not open browser.')
    def mark_sent(self):
        if not self.selected or not self.opened or self.opened[0]!=self.selected:
            messagebox.showwarning('No reviewed draft','Open and review this contact’s chat first.');return
        if messagebox.askyesno('Confirm sent','Did you actually press Send and see the outbound message in WhatsApp?'):
            try:
                self.s.event(self.selected,'Sent',self.opened[1]);self.opened=None;self.refresh();self.select();self.footer.config(text='Outbound recorded manually; delivery is not confirmed.')
            except ValueError as e:messagebox.showerror('Unable to record',str(e))
    def reply(self):
        r=self.s.match(self.sender.get())
        if not r:
            self.replyresult.config(text='Unmatched sender: no recorded outreach for this number. Nothing saved.');return
        if not messagebox.askyesno('Match found',f"Sender matches {r['name']} (+{r['phone']}). Is this an actual incoming reply?"):return
        self.s.event(r['id'],'Reply',self.replynote.get('1.0','end').strip());self.refresh()
        self.replyresult.config(text='Reply logged for '+r['name']+' • matched by phone, not message thread.')
        self.root.bell();messagebox.showinfo('Reply recorded',f"{r['name']} has a manually recorded reply.")
    def export(self):
        filename=filedialog.asksaveasfilename(defaultextension='.csv',initialfile='LeadConnect_tracker.csv',filetypes=[('CSV','*.csv')])
        if not filename:return
        try:
            keys=['name','raw','phone','kind','area','source','status','verification','consent','notes','sent','replied']
            with open(filename,'w',encoding='utf-8-sig',newline='') as f:
                w=csv.writer(f);w.writerow(keys);w.writerows([[safe_csv(r[k]) for k in keys] for r in self.s.all()])
            self.footer.config(text='Tracker exported. Keep the file private.')
        except OSError as e:messagebox.showerror('Export failed',str(e))

def main():
    if '--self-test' in sys.argv:
        store=Store(':memory:');root=tk.Tk();root.withdraw()
        app=App(root,store);root.update()
        assert normalize('01012345678') == ('201012345678','Egypt mobile')
        assert normalize('0221234567')[1] == 'Landline / other Egypt'
        assert store.import_rows([{'name':'Test','phone':'01012345678'},{'name':'Duplicate','phone':'+201012345678'},{'name':'Landline','phone':'0221234567'}]) == (2,1,0)
        app.refresh();root.update();assert len(app.tree.get_children()) == 1
        root.destroy();store.db.close()
        print('PASS: Tkinter, GUI initialization, SQLite, mobile filtering and deduplication.',flush=True)
        return
    data=Path.home()/'LeadConnectPro';data.mkdir(exist_ok=True)
    store=Store(data/'leadconnect.sqlite3');root=tk.Tk();App(root,store)
    def callback_error(exc_type, exc, tb):
        traceback.print_exception(exc_type,exc,tb)
        messagebox.showerror('LeadConnect error',str(exc)+'\nSee the Command Prompt for details.')
    root.report_callback_exception=callback_error
    try:root.mainloop()
    finally:store.db.close()
if __name__=='__main__':main()
