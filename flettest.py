import tkinter as tk
from tkinter import ttk,  messagebox, simpledialog
from tkcalendar import Calendar, DateEntry
import time
import datetime

class MeetingRecorder:
    def __init__(self, root):
        self.root = root
        self.root.title("議事録登録ツール")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = int(screen_width * 0.2)
        height = screen_height
        x_position = screen_width - width
        y_position = 0
        self.root.geometry(f"{width}x{height}+{x_position}+{y_position}")

        self.customize_design()
        
        self.meeting_name = tk.StringVar()
        self.meeting_place = tk.StringVar()
        self.meeting_var = tk.StringVar()
        self.hour_var = tk.StringVar()
        self.minute_var = tk.StringVar()
        self.recorder = tk.StringVar()
        self.participants = []
        self.start_time = None
        self.check_vars = []

        current_time = datetime.datetime.now().time()
        self.hour_var.set(current_time.hour)
        self.minute_var.set(current_time.minute)
        
        self.init_ui()
        
    def init_ui(self):
        ttk.Label(self.root, text="会議名").pack(pady=2)
        ttk.Entry(self.root, textvariable=self.meeting_name).pack(pady=3)
        
        ttk.Label(self.root, text="会議場所").pack(pady=2)
        ttk.Entry(self.root, textvariable=self.meeting_place).pack(pady=3)
        
        ttk.Label(self.root, text="日付").pack(pady=3)
        self.meeting_date = DateEntry(self.root, textvariable=self.meeting_var)
        self.meeting_date.pack(pady=3)

        ttk.Label(self.root, text="時間").pack(pady=3)
        time_frame = ttk.Frame(self.root)
        time_frame.pack(pady=3)

        self.hour_spin = ttk.Spinbox(time_frame, from_=0, to=23, width=5, format="%02.0f", textvariable=self.hour_var)
        self.hour_spin.pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT, padx=(0, 0))

        self.minute_spin = ttk.Spinbox(time_frame, from_=0, to=59, width=5, format="%02.0f", textvariable=self.minute_var)
        self.minute_spin.pack(side=tk.LEFT)
        
        ttk.Label(self.root, text="議事録者").pack(pady=2)
        ttk.Entry(self.root, textvariable=self.recorder).pack(pady=3)
        
        ttk.Label(self.root, text="参加者").pack(pady=2)
        self.participant_entry = ttk.Entry(self.root)
        self.participant_entry.pack(pady=3)
        ttk.Button(self.root, text="参加者追加", command=self.add_participant).pack(pady=2)
        self.participant_listbox = tk.Listbox(self.root,height=8)
        self.participant_listbox.pack(pady=3, fill=tk.BOTH, expand=False)    
        
        startStyle = ttk.Style()
        startStyle.configure("RedText.TButton", foreground="red", background="white")
        ttk.Button(self.root, text="開始", command=self.start_meeting, style="RedText.TButton").pack(pady=3)

        # バインド設定
        self.participant_entry.bind("<Return>", lambda event=None: self.add_participant())
        
    # メソッド
    def customize_design(self):
        # フォントの定義
        self.font_large = ("Arial", 12, "bold")
        self.font_normal = ("Arial", 10)

        # ルートウィンドウの背景色の設定
        self.root.configure(bg="#f7f7f7")

        # ボタンのスタイルの設定
        style = ttk.Style()
        style.configure("TButton", 
                        font=self.font_normal, 
                        padding=5, 
                        background="#4e4e4e", 
                        foreground="black",
                        relief="flat",
                        borderwidth=0)
        style.map("TButton",
                  background=[('active', '#6e6e6e')],
                  foreground=[('active', 'black')])

        style.configure("TLabel", 
                        font=self.font_normal, 
                        background="#f7f7f7", 
                        padding=(5, 2))

        style.configure("TEntry", 
                        font=self.font_normal, 
                        padding=5, 
                        fieldbackground="#e1e1e1", 
                        foreground="#333333",
                        insertcolor="#333333")

    def add_participant(self):
        participant = self.participant_entry.get()
        if participant and participant not in self.participants:
            self.participants.append(participant)
            self.participant_listbox.insert(tk.END, participant)
            self.participant_entry.delete(0, tk.END)
        
    def update_timer(self):
        elapsed_time = int(time.time() - self.start_time)
        mins, sec = divmod(elapsed_time, 60)
        hours, mins = divmod(mins, 60)
        self.timer_label.config(text=f"{hours:02}:{mins:02}:{sec:02}")
        self.root.after(1000, self.update_timer)

    def start_meeting(self):
        if not self.meeting_name.get().strip():
            tk.messagebox.showwarning("警告", "会議名を入力してください")
            return
        
        for widget in self.root.winfo_children():
            widget.destroy()
        
        self.start_time = time.time()

        recorder_name = self.recorder.get()
        if recorder_name and recorder_name not in self.participants:
            self.participants.append(recorder_name)
        
        ttk.Label(self.root, textvariable=self.meeting_name, font=self.font_large).pack(pady=3)
        self.timer_label = ttk.Label(self.root, text="00:00:00")
        self.timer_label.pack(pady=2)
        self.update_timer()
        
        ttk.Label(self.root, text="発言者").pack(pady=3)
        self.speaker_combobox = ttk.Combobox(self.root, values=self.participants)
        self.speaker_combobox.pack(pady=3)
        self.speaker_combobox.bind("<FocusIn>", self.save_listbox_selection)  # フォーカスイベントをバインド
        self.speaker_combobox.bind("<<ComboboxSelected>>", self.restore_listbox_selection_after_combobox)  # 選択イベントをバインド
        self.speaker_combobox.bind("<FocusOut>", self.restore_listbox_selection)  # <FocusOut>イベントをバインド
        self.speaker_combobox["postcommand"] = self.save_listbox_selection  # postcommandオプションを設定

        
        ttk.Label(self.root, text="発言内容").pack(pady=3)
        self.speech_text = tk.Text(self.root, height=5)
        self.speech_text.pack(pady=3)
        self.speech_text.bind("<Control-Return>", self.register_speech_event)  # Ctrl+Enterで発言を登録
        
        ttk.Button(self.root, text="登録", command=self.register_speech).pack(pady=3)

        self.speech_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE)  # selectmodeをMULTIPLEに設定
        self.speech_listbox.pack(fill=tk.BOTH, expand=False)
        self.speech_listbox.bind("<Double-Button-1>", self.add_memo)
        
        tk.Button(self.root, text="終了", command=self.end_meeting, fg="white", bg="red", relief="flat").pack(pady=3)
        
    def register_speech(self):
        speaker = self.speaker_combobox.get()
        speech = self.speech_text.get(1.0, tk.END).strip()
        if speaker and speech:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            self.speech_listbox.insert(tk.END, f"{timestamp} - {speaker}: {speech}")
            self.speech_text.delete(1.0, tk.END)
        if speaker and speaker not in self.participants:
            self.participants.append(speaker)
            self.participant_listbox.insert(tk.END, speaker)
        
    def add_memo(self, event):
        # 選択されたアイテムのインデックスを取得
        index = self.speech_listbox.curselection()[0]
        item = self.speech_listbox.get(index)
        
        # ダイアログを表示してメモを入力
        memo = simpledialog.askstring("メモの追加", "メモを入力してください：")
        if memo:
            # メモをアイテムに追加
            new_item = item + f" (メモ: {memo})"
            self.speech_listbox.delete(index)
            self.speech_listbox.insert(index, new_item)

    def end_meeting(self):
        date_str = self.meeting_var.get()
        time_str = f"{self.hour_var.get()}:{self.minute_var.get()}"
        datetime_str = f"{date_str} {time_str}"
        file_path = self.meeting_name.get() + ".txt"
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(f"会議名: {self.meeting_name.get()}\n")
            f.write(f"会議場所: {self.meeting_place.get()}\n")
            f.write(f"日時: {datetime_str}\n")
            f.write(f"議事録者: {self.recorder.get()}\n")
            f.write(f"参加者: {', '.join(self.participants)}\n")
            elapsed_time = int(time.time() - self.start_time)
            mins, sec = divmod(elapsed_time, 60)
            hours, mins = divmod(mins, 60)
            f.write(f"会議時間: {hours:02}:{mins:02}:{sec:02}\n\n")
            
            for idx, time_speaker in enumerate(self.speech_listbox.get(0, tk.END)):
                prefix = "＊" if "(メモ:" in time_speaker or idx in self.speech_listbox.curselection() else ""  # メモがある場合、冒頭に「＊」を付ける
                f.write(f"{prefix}{time_speaker}\n")
        
        print("出力完了")
        tk.messagebox.showinfo("情報", f"議事録が{file_path}に保存されました。")

        self.root.quit()

    # バインド時のメソッド
    def save_listbox_selection(self, event):
        # Listboxの選択状態を保存
        if self.speech_listbox.curselection():
            self.saved_selection = self.speech_listbox.curselection()
        
            # フォーカスが移動した後に選択状態を復元するためのイベントをバインド
            self.speech_listbox.after(100, self.restore_listbox_selection)

    def restore_listbox_selection(self):
        # Listboxの選択状態を復元
        for index in self.saved_selection:
            self.speech_listbox.select_set(index)
    
    def restore_listbox_selection_after_combobox(self, event):
        # 少し遅延してListboxの選択状態を復元
        self.speech_listbox.after(10, self.restore_listbox_selection)
    
    def register_speech_event(self, event):
        # Ctrl+Enterで発言を登録するためのイベントハンドラ
        self.register_speech()


if __name__ == "__main__":
    root = tk.Tk()
    app = MeetingRecorder(root)
    root.mainloop()
