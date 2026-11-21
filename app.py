import json, time, threading, queue, pathlib, sys, traceback, datetime, re, copy
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from pynput import mouse, keyboard
import pyautogui as pg
from openpyxl import load_workbook

pg.FAILSAFE = True
pg.PAUSE = 0.05
ROOT = pathlib.Path.home() / 'Documents' / 'AutoClickRecorder'
ROOT.mkdir(parents=True, exist_ok=True)

class App:
    def __init__(self, root):
        self.root=root; root.title('Auto Click Recorder V5 - Windows'); root.geometry('1020x690')
        self.steps=[]; self.events=queue.Queue(); self.recording=False; self.running=False
        self.stop=threading.Event(); self.last=0; self.ml=None; self.kl=None; self.hotkey=None
        self.excel_path=''; self.excel_rows=[]; self.excel_headers=[]; self.schedule_job=None; self.schedule_at=tk.StringVar(value=(datetime.datetime.now()+datetime.timedelta(minutes=5)).strftime('%Y-%m-%d %H:%M:%S')); self.excel_mode=tk.BooleanVar(value=False); self.loop=tk.StringVar(value='5'); self.delay=tk.StringVar(value='0.3'); self.status=tk.StringVar(value='Sẵn sàng | F8: dừng khẩn cấp')
        self.date_start=tk.StringVar(value='2026-11-20'); self.date_advance=tk.BooleanVar(value=True); self.date_skip_weekends=tk.BooleanVar(value=False); self.date_increment=tk.StringVar(value='1')
        self.window_mode=tk.BooleanVar(value=True); self.date_end=tk.StringVar(value='2026-11-30'); self.window_ranges=tk.StringVar(value='06:00-14:30,14:30-23:30'); self.build(); self.root.after(100,self.drain)
        self.hotkey=keyboard.Listener(on_press=self.emergency); self.hotkey.start()
        root.protocol('WM_DELETE_WINDOW',self.close)
    def build(self):
        bar=ttk.Frame(self.root,padding=10); bar.pack(fill='x')
        for title,cmd in [('● Ghi',self.start_record),('■ Dừng ghi',self.stop_record),('▶ Chạy',self.play),('⏹ Dừng F8',self.halt),('💾 Lưu',self.save),('📂 Mở',self.load),('＋ Bước',self.add),('✎ Sửa bước',self.edit),('✕ Xóa',self.delete),('↑',lambda:self.move(-1)),('↓',lambda:self.move(1))]:
            ttk.Button(bar,text=title,command=cmd).pack(side='left',padx=2)
        opts=ttk.Frame(self.root,padding=(10,0,10,8)); opts.pack(fill='x')
        ttk.Label(opts,text='Số vòng (0 = liên tục):').pack(side='left'); ttk.Entry(opts,textvariable=self.loop,width=8).pack(side='left',padx=5)
        ttk.Label(opts,text='Chờ giữa vòng (giây):').pack(side='left'); ttk.Entry(opts,textvariable=self.delay,width=8).pack(side='left',padx=5)
        ttk.Button(opts,text='📊 Nạp Excel',command=self.import_excel).pack(side='left',padx=4)
        ttk.Checkbutton(opts,text='Chạy từng dòng Excel',variable=self.excel_mode).pack(side='left',padx=3)
        ttk.Label(opts,text='F8 dừng ngay; rê chuột góc trên trái để dừng dự phòng',foreground='#a33').pack(side='left',padx=12)
        sched=ttk.Frame(self.root,padding=(10,0,10,8));sched.pack(fill='x')
        ttk.Label(sched,text='Hẹn giờ (YYYY-MM-DD HH:MM:SS):').pack(side='left')
        ttk.Entry(sched,textvariable=self.schedule_at,width=22).pack(side='left',padx=5)
        ttk.Button(sched,text='⏰ Hẹn chạy',command=self.arm_schedule).pack(side='left',padx=3)
        ttk.Button(sched,text='Hủy hẹn',command=self.cancel_schedule).pack(side='left',padx=3)
        ttk.Button(sched,text='＋ Chọn ngày/giờ',command=self.add_datetime_step).pack(side='left',padx=3)
        ttk.Button(sched,text='＋ Nhập từ Excel',command=self.add_excel_step).pack(side='left',padx=3)
        dates=ttk.Frame(self.root,padding=(10,0,10,8));dates.pack(fill='x')
        ttk.Label(dates,text='Ngày bắt đầu (YYYY-MM-DD):').pack(side='left')
        ttk.Entry(dates,textvariable=self.date_start,width=13).pack(side='left',padx=5)
        ttk.Checkbutton(dates,text='Tăng ngày sau mỗi vòng hoàn tất',variable=self.date_advance).pack(side='left',padx=8)
        ttk.Label(dates,text='Số ngày tăng:').pack(side='left')
        ttk.Entry(dates,textvariable=self.date_increment,width=5).pack(side='left',padx=5)
        ttk.Checkbutton(dates,text='Bỏ qua thứ 7, CN',variable=self.date_skip_weekends).pack(side='left',padx=8)
        windowbar=ttk.Frame(self.root,padding=(10,0,10,8));windowbar.pack(fill='x')
        ttk.Checkbutton(windowbar,text='V5: Chạy từng khung giờ, hết ngày mới tăng ngày',variable=self.window_mode).pack(side='left',padx=3)
        ttk.Label(windowbar,text='Đến ngày:').pack(side='left',padx=4)
        ttk.Entry(windowbar,textvariable=self.date_end,width=12).pack(side='left')
        ttk.Label(windowbar,text='Khung giờ:').pack(side='left',padx=4)
        ttk.Entry(windowbar,textvariable=self.window_ranges,width=30).pack(side='left')
        ttk.Button(windowbar,text='＋ Mẫu 4 ô ngày/giờ + tải',command=self.add_window_template).pack(side='left',padx=4)
        columns=('kind','data','wait'); self.tree=ttk.Treeview(self.root,columns=columns,show='headings',selectmode='browse')
        for col,txt,w in [('kind','Loại bước',160),('data','Dữ liệu / tọa độ / ảnh',650),('wait','Chờ trước (s)',110)]: self.tree.heading(col,text=txt); self.tree.column(col,width=w)
        self.tree.pack(fill='both',expand=True,padx=10,pady=5); self.tree.bind('<Double-1>',lambda e:self.edit())
        ttk.Label(self.root,textvariable=self.status,padding=10).pack(fill='x')
        ttk.Label(self.root,text='Bước hỗ trợ: click, scroll, key, text, wait, wait_image, click_image. Để nhận diện ảnh: chụp vùng nút bằng công cụ Snipping Tool, lưu PNG rồi thêm bước.',padding=(10,0,10,10),wraplength=950).pack(fill='x')
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for i,s in enumerate(self.steps):
            detail={k:v for k,v in s.items() if k not in ('kind','wait')}
            self.tree.insert('', 'end',iid=str(i),values=(s['kind'],str(detail),s.get('wait',0)))
    def emergency(self,key):
        if key==keyboard.Key.f8: self.stop.set(); self.recording=False; self.events.put(('status','Đã nhấn F8: yêu cầu dừng'))
    def start_record(self):
        if self.running or self.recording:return
        self.recording=True; self.last=time.monotonic(); self.status.set('Đang ghi... F8 dừng. Không ghi mật khẩu / thông tin nhạy cảm.')
        def push(s):
            if not self.recording:return
            now=time.monotonic(); s['wait']=round(min(now-self.last,60),2); self.last=now; self.events.put(('step',s))
        def click(x,y,button,pressed):
            if pressed and button==mouse.Button.left: push({'kind':'click','x':int(x),'y':int(y)})
            elif pressed and button==mouse.Button.right: push({'kind':'right_click','x':int(x),'y':int(y)})
        def scroll(x,y,dx,dy): push({'kind':'scroll','amount':int(dy)})
        def keypress(key):
            if key==keyboard.Key.f8:return
            try: val=key.char
            except AttributeError: val=str(key).replace('Key.','')
            if val:push({'kind':'key','value':val})
        self.ml=mouse.Listener(on_click=click,on_scroll=scroll); self.kl=keyboard.Listener(on_press=keypress)
        self.ml.start();self.kl.start()
    def stop_record(self):
        self.recording=False
        if self.ml:self.ml.stop();self.ml=None
        if self.kl:self.kl.stop();self.kl=None
        self.status.set('Đã dừng ghi')
    def drain(self):
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='step':self.steps.append(value);self.refresh()
                elif kind=='status':self.status.set(value)
                elif kind=='error':messagebox.showerror('Lỗi thực thi',value)
                elif kind=='done':self.running=False;self.status.set(value)
        except queue.Empty:pass
        self.root.after(100,self.drain)
    def selected(self):
        s=self.tree.selection();return int(s[0]) if s else None
    def add(self):
        kind=simpledialog.askstring('Thêm bước','Loại: click / right_click / scroll / key / text / wait / wait_image / click_image',parent=self.root)
        if not kind:return
        kind=kind.strip().lower()
        if kind not in ('click','right_click','scroll','key','text','wait','wait_image','click_image','set_datetime','excel_text'):return messagebox.showwarning('Không hợp lệ','Loại bước chưa hỗ trợ')
        s={'kind':kind,'wait':0.2}
        if kind in ('click','right_click'):
            s['x']=pg.position().x;s['y']=pg.position().y
        elif kind=='scroll':s['amount']=simpledialog.askinteger('Cuộn','Số nấc (+ lên / - xuống)',initialvalue=-3) or -3
        elif kind in ('key','text'):
            s['value']=simpledialog.askstring('Nội dung','Nhập phím hoặc văn bản (phím đặc biệt: enter, tab, esc, ctrl+c...)') or ''
        elif kind=='wait':s['seconds']=simpledialog.askfloat('Chờ','Số giây',initialvalue=2.0,minvalue=0.0) or 0
        elif kind=='set_datetime':s.update(value='{run_date}',format='%d/%m/%Y',mode='auto',select_all='true',verify_image='',calendar_image='',day_image='',timeout=10,confidence=0.85)
        elif kind=='excel_text':s.update(column=self.excel_headers[0] if self.excel_headers else 'Mã SAP')
        else:
            p=filedialog.askopenfilename(filetypes=[('Ảnh PNG','*.png'),('Ảnh','*.png *.jpg *.jpeg')]);
            if not p:return
            s.update(image=p,timeout=20,confidence=0.85)
        self.steps.append(s);self.refresh()
    def edit(self):
        i=self.selected()
        if i is None:return
        step=copy.deepcopy(self.steps[i]); win=tk.Toplevel(self.root)
        win.title(f'Sửa bước #{i+1}');win.geometry('600x600');win.transient(self.root);win.grab_set()
        ttk.Label(win,text=f'Bước {i+1}: {step.get("kind")}',font=('Segoe UI',13,'bold')).pack(anchor='w',padx=14,pady=10)
        frame=ttk.Frame(win,padding=12);frame.pack(fill='both',expand=True)
        values={}; kinds=('click','right_click','scroll','key','text','wait','wait_image','click_image','set_datetime','excel_text')
        def add_field(name,value):
            row=ttk.Frame(frame);row.pack(fill='x',pady=4)
            ttk.Label(row,text=name,width=16).pack(side='left')
            var=tk.StringVar(value=str(value));ttk.Entry(row,textvariable=var).pack(side='left',fill='x',expand=True)
            values[name]=var
        add_field('kind',step.get('kind','click'))
        keys=['wait','x','y','amount','value','seconds','image','timeout','confidence','column','format','mode','select_all','verify_image','calendar_image','day_image','label']
        for key in keys:
            if key in step or key=='wait':add_field(key,step.get(key,0 if key=='wait' else ''))
        ttk.Label(frame,text='Biến: {run_date}, {window_start_time}, {window_start_date}, {window_end_time}, {window_end_date}, {Tên cột Excel}\nNgày/giờ: mode=auto/text/calendar/hotkey. Auto nhập trực tiếp, kiểm tra verify_image nếu có; nếu thất bại dùng calendar_image và day_image. Để trống ảnh kiểm tra: không thể xác minh nhập thành công.',wraplength=480).pack(anchor='w',pady=8)
        def commit():
            try:
                d={}
                for key,var in values.items():
                    v=var.get()
                    if key in ('x','y','amount'):v=int(v)
                    elif key in ('wait','seconds','timeout','confidence'):v=float(v)
                    d[key]=v
                if d['kind'] not in kinds:raise ValueError('Loại bước không hợp lệ')
                if float(d.get('wait',0))<0:raise ValueError('Thời gian chờ phải >= 0')
                self.steps[i]=d;self.refresh();self.tree.selection_set(str(i));win.destroy()
            except Exception as e:messagebox.showerror('Không lưu được',str(e),parent=win)
        ttk.Button(win,text='Lưu thay đổi',command=commit).pack(pady=10)

    def add_window_template(self):
        if self.steps and not messagebox.askyesno('Thêm mẫu V5','Thêm 8 bước vào cuối kịch bản hiện tại?\\nCác bước click vào ô cần được ghi hoặc chỉnh tọa độ trước khi chạy.'):
            return
        template=[]
        for label,token in [('giờ bắt đầu','{window_start_time}'),('ngày bắt đầu','{window_start_date}'),('giờ kết thúc','{window_end_time}'),('ngày kết thúc','{window_end_date}')]:
            template.append({'kind':'click','x':0,'y':0,'wait':0.2,'label':'Click ô '+label})
            template.append({'kind':'set_datetime','value':token,'format':'%d/%m/%Y','mode':'auto','select_all':'true','wait':0.2,'label':'Nhập '+label})
        template.extend([
            {'kind':'click_image','image':'','timeout':20,'confidence':0.85,'wait':0.2,'label':'Bấm nút tải danh sách'},
            {'kind':'wait_image','image':'','timeout':120,'confidence':0.85,'wait':0.2,'label':'Chờ dữ liệu sẵn sàng'},
            {'kind':'click_image','image':'','timeout':20,'confidence':0.85,'wait':0.2,'label':'Bấm nút tải file'},
        ])
        self.steps.extend(template);self.refresh()
        self.status.set('Đã thêm mẫu V5. Cần chỉnh tọa độ 4 ô và ảnh nút tải/dữ liệu sẵn sàng trước khi chạy.')

    @staticmethod
    def parse_windows(spec):
        windows=[]
        for segment in spec.split(','):
            segment=segment.strip()
            if not segment:continue
            match=re.fullmatch(r'(\\d{2}:\\d{2})\\s*-\\s*(\\d{2}:\\d{2})',segment)
            if not match:raise ValueError('Khung giờ phải theo mẫu 06:00-14:30,14:30-23:30')
            a,b=(datetime.datetime.strptime(t,'%H:%M').time() for t in match.groups())
            if a>=b:raise ValueError('Giờ bắt đầu phải nhỏ hơn giờ kết thúc: '+segment)
            windows.append((a,b))
        if not windows:raise ValueError('Cần ít nhất một khung giờ')
        return windows

    def add_datetime_step(self):
        self.steps.append({'kind':'set_datetime','value':'{run_date}','format':'%d/%m/%Y','mode':'auto','select_all':'true','verify_image':'','calendar_image':'','day_image':'','timeout':10,'confidence':0.85,'wait':0.2});self.refresh()
        self.tree.selection_set(str(len(self.steps)-1));self.edit()
    def add_excel_step(self):
        self.steps.append({'kind':'excel_text','column':self.excel_headers[0] if self.excel_headers else 'Mã SAP','wait':0.2});self.refresh()
        self.tree.selection_set(str(len(self.steps)-1));self.edit()
    def import_excel(self):
        path=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx *.xlsm')])
        if not path:return
        try:
            wb=load_workbook(path,read_only=True,data_only=True)
            ws=wb.active;iterator=ws.iter_rows(values_only=True)
            headers=[str(v).strip() if v is not None else '' for v in next(iterator)]
            if not any(headers) or len(set(h for h in headers if h))!=len([h for h in headers if h]):raise ValueError('Tiêu đề cột trống hoặc trùng')
            rows=[]
            for row in iterator:
                if not any(v is not None and str(v).strip() for v in row):continue
                rows.append({h:row[j] if j<len(row) else None for j,h in enumerate(headers) if h})
            wb.close()
            if not rows:raise ValueError('Excel không có dòng dữ liệu')
            self.excel_path=path;self.excel_headers=headers;self.excel_rows=rows;self.excel_mode.set(True)
            self.status.set(f'Excel: {len(rows)} dòng, {len(headers)} cột. Mỗi dòng chạy trọn kịch bản một lần.')
        except Exception as e:messagebox.showerror('Lỗi Excel',str(e))
    def cancel_schedule(self):
        if self.schedule_job is not None:self.root.after_cancel(self.schedule_job);self.schedule_job=None
        self.status.set('Đã hủy lịch hẹn')
    def arm_schedule(self):
        try:
            at=datetime.datetime.strptime(self.schedule_at.get().strip(),'%Y-%m-%d %H:%M:%S')
            delay=(at-datetime.datetime.now()).total_seconds()
            if delay<=0:raise ValueError('Thời điểm phải ở tương lai')
            if delay>30*86400:raise ValueError('Chỉ hỗ trợ hẹn tối đa 30 ngày')
            if self.running or self.recording:raise ValueError('Dừng ghi/chạy trước khi hẹn')
            self.cancel_schedule()
            self.schedule_job=self.root.after(int(delay*1000),self.run_schedule)
            self.status.set('Đã hẹn chạy lúc '+at.strftime('%d/%m/%Y %H:%M:%S')+'; giữ ứng dụng mở')
        except Exception as e:messagebox.showerror('Lỗi lịch hẹn',str(e))
    def run_schedule(self):
        self.schedule_job=None;self.play(confirmed=True)
    def resolve(self,value,row):
        if not isinstance(value,str):return value
        now=datetime.datetime.now()
        def replace(m):
            token=m.group(1)
            if token in ('window_start_time','window_end_time','window_start_date','window_end_date'):
                if token not in row:raise ValueError('Biến '+token+' chỉ dùng trong chế độ V5')
                return row[token]
            if token=='run_date':return row.get('__run_date',now.date()).strftime('%d/%m/%Y')
            if token=='run_time':return row.get('__run_date',now.date()).strftime('%d/%m/%Y')+' '+now.strftime('%H:%M:%S')
            if token.startswith('run_date:'):return row.get('__run_date',now.date()).strftime(token[9:])
            if token=='today':return now.strftime('%d/%m/%Y')
            if token=='now':return now.strftime('%d/%m/%Y %H:%M:%S')
            if token.startswith('date:'):return now.strftime(token[5:])
            if token in row:
                v=row[token]
                if isinstance(v,(datetime.date,datetime.datetime)):return v.strftime('%d/%m/%Y' if isinstance(v,datetime.date) and not isinstance(v,datetime.datetime) else '%d/%m/%Y %H:%M:%S')
                return '' if v is None else str(v)
            raise ValueError('Không có cột/biến: '+token)
        return re.sub(r'\{([^{}]+)\}',replace,value)
    def delete(self):
        i=self.selected()
        if i is not None:self.steps.pop(i);self.refresh()
    def move(self,d):
        i=self.selected()
        if i is None or not 0<=i+d<len(self.steps):return
        self.steps[i],self.steps[i+d]=self.steps[i+d],self.steps[i];self.refresh();self.tree.selection_set(str(i+d))
    def save(self):
        p=filedialog.asksaveasfilename(initialdir=ROOT,defaultextension='.json',filetypes=[('Kịch bản JSON','*.json')]);
        if p:pathlib.Path(p).write_text(json.dumps({'version':5,'window_mode':self.window_mode.get(),'date_end':self.date_end.get(),'window_ranges':self.window_ranges.get(),'steps':self.steps,'excel_path':self.excel_path,'excel_mode':self.excel_mode.get(),'schedule_at':self.schedule_at.get(),'date_start':self.date_start.get(),'date_advance':self.date_advance.get(),'date_increment':self.date_increment.get(),'date_skip_weekends':self.date_skip_weekends.get()},ensure_ascii=False,indent=2),encoding='utf-8');self.status.set('Đã lưu '+p)
    def load(self):
        p=filedialog.askopenfilename(initialdir=ROOT,filetypes=[('Kịch bản JSON','*.json')]);
        if not p:return
        try:
            d=json.loads(pathlib.Path(p).read_text(encoding='utf-8'));steps=d['steps']
            if not isinstance(steps,list):raise ValueError('Sai định dạng')
            self.steps=steps;self.refresh();self.window_mode.set(bool(d.get('window_mode',False)));self.date_end.set(d.get('date_end',self.date_end.get()));self.window_ranges.set(d.get('window_ranges',self.window_ranges.get()));self.schedule_at.set(d.get('schedule_at',self.schedule_at.get()));self.date_start.set(d.get('date_start',self.date_start.get()));self.date_advance.set(bool(d.get('date_advance',True)));self.date_increment.set(str(d.get('date_increment','1')));self.date_skip_weekends.set(bool(d.get('date_skip_weekends',False)));self.excel_mode.set(bool(d.get('excel_mode',False)));self.status.set('Đã mở '+p+' (Excel cần nạp lại khi mở)')
        except Exception as e:messagebox.showerror('Lỗi mở file',str(e))
    def halt(self):self.stop.set();self.stop_record();self.status.set('Đang yêu cầu dừng...')
    def sleep(self,t):
        if self.stop.wait(max(0,float(t))):raise InterruptedError('Đã dừng')
    def find_image(self,path,step,timeout=10):
        if not path or not pathlib.Path(path).is_file():raise FileNotFoundError('Chưa cấu hình ảnh hoặc không tìm thấy ảnh: '+str(path))
        deadline=time.monotonic()+max(0,float(timeout))
        while not self.stop.is_set():
            try:
                found=pg.locateCenterOnScreen(path,confidence=float(step.get('confidence',0.85)))
            except pg.ImageNotFoundException:
                found=None
            if found:return found
            if time.monotonic()>=deadline:return None
            self.sleep(min(0.35,max(0,deadline-time.monotonic())))
        raise InterruptedError('Đã dừng')

    def select_calendar_day(self,step):
        icon=str(step.get('calendar_image','')).strip()
        day=str(step.get('day_image','')).strip()
        if not icon or not day:
            raise ValueError('Chọn ngày bằng lịch cần calendar_image và day_image phù hợp với ngày hiện tại. Không thể suy ra ô ngày từ ảnh lịch chung.')
        timeout=float(step.get('timeout',10))
        pos=self.find_image(icon,step,timeout)
        if pos is None:raise TimeoutError('Không thấy biểu tượng mở lịch')
        pg.click(pos)
        self.sleep(0.4)
        pos=self.find_image(day,step,timeout)
        if pos is None:raise TimeoutError('Không thấy ô ngày trong lịch; cần ảnh day_image ứng với ngày đang chạy')
        pg.click(pos)
        verify=str(step.get('verify_image','')).strip()
        if verify and self.find_image(verify,step,timeout) is None:
            raise TimeoutError('Đã click lịch nhưng không xác minh được ngày đã chọn')

    def execute(self,s,row):
        kind=s['kind']
        if kind in ('click','right_click'):
            x,y=int(s['x']),int(s['y'])
            if x==0 and y==0:raise ValueError('Chưa cấu hình tọa độ click: '+str(s.get('label','')))
            pg.click(x,y,button='right' if kind=='right_click' else 'left')
        elif kind=='scroll':pg.scroll(int(s['amount']))
        elif kind=='key':
            v=str(self.resolve(s['value'],row))
            if '+' in v:pg.hotkey(*[x.strip().lower() for x in v.split('+')])
            else:pg.press(v.lower()) if len(v)>1 else pg.write(v)
        elif kind=='text':pg.write(str(self.resolve(s['value'],row)),interval=0.02)
        elif kind=='wait':self.sleep(s.get('seconds',1))
        elif kind=='excel_text':
            col=s['column']
            if col not in row:raise ValueError('Excel không có cột: '+col)
            value=row[col]
            if isinstance(value,(datetime.date,datetime.datetime)):value=value.strftime('%d/%m/%Y')
            pg.write('' if value is None else str(value),interval=0.02)
        elif kind=='set_datetime':
            raw=str(s.get('value','{run_date}'))
            fmt=str(s.get('format','%d/%m/%Y'))
            if raw=='{run_date}':value=row.get('__run_date',datetime.date.today()).strftime(fmt)
            elif raw=='{today}':value=datetime.datetime.now().strftime(fmt)
            else:value=str(self.resolve(raw,row))
            mode=str(s.get('mode','auto')).lower()
            if mode=='hotkey':
                pg.hotkey(*value.lower().split('+'))
            elif mode in ('text','auto'):
                if str(s.get('select_all','true')).lower() in ('true','1','yes'):
                    pg.hotkey('ctrl','a')
                pg.write(value,interval=0.02)
                if mode=='auto':
                    # Không thể xác nhận ô nhận ngày chỉ từ thao tác gõ.
                    # Chỉ fallback nếu đã cấu hình ảnh xác minh và ảnh lịch.
                    verify=str(s.get('verify_image','')).strip()
                    if verify and not self.find_image(verify,s,timeout=float(s.get('timeout',10))):
                        self.select_calendar_day(s)
            elif mode=='calendar':
                self.select_calendar_day(s)
            else:raise ValueError('Chế độ chọn ngày không hợp lệ: '+mode)
        elif kind in ('wait_image','click_image'):
            deadline=time.monotonic()+float(s.get('timeout',20));image=s['image'];found=None
            if not image or not pathlib.Path(image).is_file():raise FileNotFoundError('Thiếu ảnh mẫu: '+str(s.get('label',kind)))
            while time.monotonic()<deadline:
                if self.stop.is_set():raise InterruptedError('Đã dừng')
                try:found=pg.locateCenterOnScreen(image,confidence=float(s.get('confidence',0.85)))
                except pg.ImageNotFoundException:found=None
                if found:break
                self.sleep(0.35)
            if not found:raise TimeoutError('Không tìm thấy ảnh trong thời gian chờ: '+image)
            if kind=='click_image':pg.click(found)
        else:raise ValueError('Bước không hỗ trợ: '+str(kind))
    def play(self,confirmed=False):
        if self.running or self.recording or not self.steps:return
        try:
            loops=int(self.loop.get());delay=float(self.delay.get());assert loops>=0 and delay>=0
        except Exception:return messagebox.showerror('Lỗi','Số vòng và thời gian chờ phải >= 0')
        try:
            run_date=datetime.date.fromisoformat(self.date_start.get().strip())
            increment=int(self.date_increment.get())
            if increment<1 or increment>365:raise ValueError('Số ngày tăng phải từ 1 đến 365')
        except Exception as e:return messagebox.showerror('Ngày không hợp lệ',str(e))
        use_windows=self.window_mode.get()
        try:
            windows=self.parse_windows(self.window_ranges.get()) if use_windows else []
            end_date=datetime.date.fromisoformat(self.date_end.get().strip()) if use_windows else run_date
            if use_windows and end_date<run_date:raise ValueError('Ngày kết thúc phải >= ngày bắt đầu')
        except Exception as e:return messagebox.showerror('Cấu hình V5 không hợp lệ',str(e))
        if self.excel_mode.get() and not self.excel_rows:return messagebox.showerror('Thiếu Excel','Nạp file Excel trước khi chạy chế độ từng dòng')
        if not confirmed and not messagebox.askyesno('Xác nhận','Bắt đầu tự động điều khiển chuột và bàn phím?\nF8 để dừng. V5 chạy đến hết ngày kết thúc; V4 vòng 0 chạy liên tục.') :return
        self.date_advance_value=self.date_advance.get();self.date_increment_value=increment;self.date_skip_weekends_value=self.date_skip_weekends.get()
        self.stop.clear();self.running=True;steps=copy.deepcopy(self.steps);rows=list(self.excel_rows) if self.excel_mode.get() else [{}]
        if use_windows:
            for step in steps:
                if step.get('kind') in ('click_image','wait_image') and not step.get('image'):
                    self.running=False
                    return messagebox.showerror('Thiếu ảnh mẫu','Cấu hình ảnh cho bước '+str(step.get('label',step['kind'])))
                if step.get('kind')=='click' and (int(step.get('x',0)),int(step.get('y',0)))==(0,0):
                    self.running=False
                    return messagebox.showerror('Thiếu tọa độ','Cấu hình vị trí click: '+str(step.get('label','')))
        def work():
            n=0;current_date=run_date
            try:
                while not self.stop.is_set() and ((use_windows and current_date<=end_date) or (not use_windows and (loops==0 or n<loops))):
                    n+=1;self.events.put(('status',f'Đang chạy ngày {current_date:%d/%m/%Y}'))
                    day_windows=windows if use_windows else [(None,None)]
                    for window_index,(start_time,end_time) in enumerate(day_windows,1):
                        for rownum,source_row in enumerate(rows,1):
                            row=dict(source_row);row['__run_date']=current_date
                            if use_windows:
                                row.update(window_start_time=start_time.strftime('%H:%M'),
                                           window_end_time=end_time.strftime('%H:%M'),
                                           window_start_date=current_date.strftime('%d/%m/%Y'),
                                           window_end_date=current_date.strftime('%d/%m/%Y'))
                            if self.stop.is_set():raise InterruptedError('Đã dừng')
                            self.events.put(('status',f'Ngày {current_date:%d/%m/%Y} | khung {window_index}/{len(day_windows)} | Excel {rownum}/{len(rows)}'))
                            for i,step in enumerate(steps):
                                self.sleep(step.get('wait',0));self.execute(step,row)
                            self.sleep(delay)
                    if use_windows or self.date_advance_value:
                        current_date += datetime.timedelta(days=self.date_increment_value)
                        if self.date_skip_weekends_value:
                            while current_date.weekday()>=5:current_date+=datetime.timedelta(days=1)
                self.events.put(('done',f'Đã dừng sau {n} vòng'))
            except InterruptedError:self.events.put(('done',f'Đã dừng sau {n} vòng'))
            except Exception as e:
                self.stop.set();self.events.put(('error',f'Vòng {n}: {e}'));self.events.put(('done','Đã dừng vì lỗi'))
        threading.Thread(target=work,daemon=True).start()
    def close(self):
        self.cancel_schedule();self.halt()
        if self.hotkey:self.hotkey.stop()
        self.root.destroy()

if __name__=='__main__':
    root=tk.Tk();App(root);root.mainloop()
