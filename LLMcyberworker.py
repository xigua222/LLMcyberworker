import csv
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import requests
import time
import os
import hashlib
from datetime import datetime
from threading import Thread, Event

# =================默认配置=================
DEFAULT_CONFIG = {
    "api_url": "https://api.example.com/v1/chat/completions",
    "api_key": "your-api-key-here",
    "model_name": "gpt-3.5-turbo",
    "system_prompt": "你是一个有帮助的助手。",
    "user_prompt": "请处理以下文本：\n{text}",
    "id_column": "id",
    "text_column": "text",
    "max_retries": 3,
    "max_text_length": 4000,
    "max_workers": 10,
    "requests_per_second": 10,
    "rate_bucket": 15
}


class ConfigManager:
    CONFIG_FILE = "config.json"

    @classmethod
    def load_config(cls):
        if not os.path.exists(cls.CONFIG_FILE):
            return DEFAULT_CONFIG.copy()

        try:
            with open(cls.CONFIG_FILE, 'r') as f:
                data = json.load(f)

            config_hash = data.pop('config_hash', '')
            calculated_hash = cls._calculate_hash(data)

            if config_hash != calculated_hash:
                messagebox.showwarning("配置验证", "配置文件校验失败，已恢复默认配置")
                return DEFAULT_CONFIG.copy()

            return data
        except Exception as e:
            messagebox.showerror("配置错误", f"加载配置失败: {str(e)}")
            return DEFAULT_CONFIG.copy()

    @classmethod
    def save_config(cls, config):
        try:
            config_hash = cls._calculate_hash(config)
            config_with_hash = config.copy()
            config_with_hash['config_hash'] = config_hash

            with open(cls.CONFIG_FILE, 'w') as f:
                json.dump(config_with_hash, f, indent=2)

            return True
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存配置: {str(e)}")
            return False

    @staticmethod
    def _calculate_hash(config_dict):
        sorted_str = json.dumps(config_dict, sort_keys=True).encode('utf-8')
        return hashlib.sha256(sorted_str).hexdigest()


class GenericAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("LLMcyberworkerV1.0")

        self.config = ConfigManager.load_config()
        self.running_config = self.config.copy()

        self.id_column = self.running_config['id_column']
        self.text_column = self.running_config['text_column']
        self.model_name = self.running_config['model_name']

        self.stop_event = Event()
        self.current_write_index = 0
        self.checkpoint_file = ".progress_checkpoint"
        self.current_input_hash = ""
        self.lock = threading.Lock()
        self.results_buffer = {}
        self.total_tokens = 0  # Token统计

        self.setup_ui()
        self.setup_menu()
        self.update_api_headers()
        self.setup_rate_limiter()

    def setup_rate_limiter(self):
        self.rate_limiter = threading.BoundedSemaphore(
            self.running_config['rate_bucket']
        )
        self.last_request_time = 0

    def update_api_headers(self):
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.running_config['api_key']}",
            "Accept": "application/json"
        }

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="API配置", command=self.configure_api)
        config_menu.add_command(label="列名配置", command=self.configure_columns)
        config_menu.add_command(label="处理配置", command=self.configure_processing)
        menubar.add_cascade(label="配置", menu=config_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="输入文件:").grid(row=0, column=0, sticky=tk.W)
        self.input_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.input_path, width=50).grid(row=0, column=1)
        ttk.Button(main_frame, text="浏览...", command=self.select_input_file).grid(row=0, column=2)

        ttk.Label(main_frame, text="输出文件:").grid(row=1, column=0, sticky=tk.W)
        self.output_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=1, column=1)
        ttk.Button(main_frame, text="浏览...", command=self.select_output_file).grid(row=1, column=2)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=15)
        self.start_btn = ttk.Button(btn_frame, text="开始处理", command=self.start_processing)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(btn_frame, text="停止处理", command=self.stop_processing, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=10)

        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)

        self.token_var = tk.StringVar(value="已用Token: 0")
        ttk.Label(status_frame, textvariable=self.token_var).pack(side=tk.RIGHT)

        self.log_text = tk.Text(main_frame, height=10, width=70)
        self.log_text.grid(row=5, column=0, columnspan=3, pady=10)
        self.log_text.tag_config('error', foreground='red')

        prompt_frame = ttk.LabelFrame(main_frame, text="处理指令配置")
        prompt_frame.grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=5)

        ttk.Label(prompt_frame, text="系统指令:").grid(row=0, column=0, sticky=tk.W)
        self.system_prompt_entry = tk.Text(prompt_frame, height=3, width=60)
        self.system_prompt_entry.insert(tk.END, self.running_config['system_prompt'])
        self.system_prompt_entry.grid(row=0, column=1)

        ttk.Label(prompt_frame, text="用户指令:").grid(row=1, column=0, sticky=tk.W)
        self.user_prompt_entry = tk.Text(prompt_frame, height=3, width=60)
        self.user_prompt_entry.insert(tk.END, self.running_config['user_prompt'])
        self.user_prompt_entry.grid(row=1, column=1)

    def configure_api(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("API配置")

        entries = {}
        fields = [
            ('api_url', "API地址:"),
            ('api_key', "API密钥:"),
            ('model_name', "模型名称:")
        ]

        for i, (field, label) in enumerate(fields):
            ttk.Label(dialog, text=label).grid(row=i, column=0, padx=5, pady=5)
            entry = ttk.Entry(dialog, width=40)
            entry.grid(row=i, column=1, padx=5, pady=5)
            entry.insert(0, self.running_config[field])
            entries[field] = entry

        def save_config():
            for field, entry in entries.items():
                self.running_config[field] = entry.get().strip()

            self.model_name = self.running_config['model_name']
            self.update_api_headers()

            if ConfigManager.save_config(self.running_config):
                dialog.destroy()
                messagebox.showinfo("成功", "API配置已保存")

        ttk.Button(dialog, text="保存", command=save_config).grid(row=3, column=1, pady=10)

    def configure_processing(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("处理配置")

        entries = {}
        fields = [
            ('max_retries', "最大重试次数:"),
            ('requests_per_second', "请求速率(次/秒):"),
            ('max_workers', "并发线程数:"),
            ('rate_bucket', "流量控制窗口:")
        ]

        for i, (field, label) in enumerate(fields):
            ttk.Label(dialog, text=label).grid(row=i, column=0, padx=5, pady=5)
            entry = ttk.Entry(dialog, width=15)
            entry.grid(row=i, column=1, padx=5, pady=5)
            entry.insert(0, str(self.running_config[field]))
            entries[field] = entry

        def save_config():
            try:
                for field, entry in entries.items():
                    value = int(entry.get())
                    if value <= 0:
                        raise ValueError
                    self.running_config[field] = value

                self.setup_rate_limiter()
                if ConfigManager.save_config(self.running_config):
                    dialog.destroy()
                    messagebox.showinfo("成功", "处理配置已保存")
            except ValueError:
                messagebox.showerror("错误", "请输入正整数")

        ttk.Button(dialog, text="保存", command=save_config).grid(row=4, column=1, pady=10)

    def configure_columns(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("列名配置")
        dialog.columnconfigure(1, weight=1)  # 添加弹性布局

        ttk.Label(dialog, text="ID列名:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        id_entry = ttk.Entry(dialog, width=25)
        id_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        id_entry.insert(0, self.running_config['id_column'])

        ttk.Label(dialog, text="文本列名:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        text_entry = ttk.Entry(dialog, width=25)
        text_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        text_entry.insert(0, self.running_config['text_column'])

        def save_columns():
            self.running_config['id_column'] = id_entry.get()
            self.running_config['text_column'] = text_entry.get()
            self.id_column = self.running_config['id_column']
            self.text_column = self.running_config['text_column']

            if ConfigManager.save_config(self.running_config):
                dialog.destroy()
                messagebox.showinfo("成功", "列名配置已更新")
            else:
                messagebox.showerror("错误", "配置保存失败")

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky=tk.EW)
        ttk.Button(btn_frame, text="保存", command=save_columns).pack(side=tk.RIGHT)

    def select_input_file(self):
        path = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if path:
            new_hash = self.calculate_file_fingerprint(path)
            if path != self.input_path.get() or new_hash != self.current_input_hash:
                self.clean_checkpoint()
            self.input_path.set(path)
            self.current_input_hash = new_hash
            self.log_message(f"输入文件已选择: {path}")

    def select_output_file(self):
        path = filedialog.asksaveasfilename(
            title="保存结果文件",
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv")]
        )
        if path:
            self.output_path.set(path)
            self.log_message(f"输出路径已设置: {path}")

    def clean_checkpoint(self):
        if os.path.exists(self.checkpoint_file):
            try:
                os.remove(self.checkpoint_file)
                self.log_message("已清除旧检查点")
            except Exception as e:
                self.log_message(f"检查点清理失败: {str(e)}", is_error=True)

    def log_message(self, message, is_error=False):
        timestamp = datetime.now().strftime("%H:%M:%S")
        msg = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, msg)
        if is_error:
            self.log_text.tag_add('error', "end-2l linestart", "end-2l lineend")
        self.log_text.see(tk.END)

    def update_progress(self, value, total):
        self.progress['value'] = value
        self.progress['maximum'] = total
        self.status_var.set(f"处理进度: {value}/{total} ({value / total:.1%})")

    def safe_api_call(self, text, index):
        if self.stop_event.is_set():
            return (index, "已中止")

        if pd.isna(text) or str(text).strip() == "":
            return (index, "空输入")

        clean_text = str(text).strip()[:self.running_config['max_text_length']]
        payload = {
            "messages": [
                {"role": "system", "content": self.system_prompt_entry.get("1.0", tk.END).strip()},
                {"role": "user", "content": self.user_prompt_entry.get("1.0", tk.END).strip().format(text=clean_text)}
            ],
            "model": self.model_name,
            "temperature": 0.7,
            "max_tokens": 1000
        }

        for attempt in range(self.running_config['max_retries']):
            try:
                with self.rate_limiter:
                    now = time.time()
                    elapsed = now - self.last_request_time
                    interval = 1.0 / self.running_config['requests_per_second']

                    if elapsed < interval:
                        time.sleep(interval - elapsed)

                    self.last_request_time = time.time()

                response = requests.post(
                    self.running_config['api_url'],
                    headers=self.headers,
                    json=payload,
                    timeout=(10, 30)
                )
                response.raise_for_status()

                response_data = response.json()
                usage = response_data.get('usage', {})
                tokens = usage.get('total_tokens', 0)
                self.root.after(0, self.update_token_count, tokens)

                raw_result = response_data['choices'][0]['message']['content'].strip()
                return (index, raw_result)

            except requests.exceptions.HTTPError as e:
                status_code = e.response.status_code
                error_msg = f"API错误[{status_code}]: {str(e)}"

                if status_code == 429:
                    retry_after = int(e.response.headers.get('Retry-After', 5))
                    self.log_message(f"请求过载，等待{retry_after}秒重试... (尝试 {attempt + 1})")
                    time.sleep(retry_after)
                    continue

                elif 500 <= status_code < 600:
                    self.log_message(f"服务器错误，等待{2 ** attempt}秒重试...", is_error=True)
                    time.sleep(2 ** attempt)
                    continue

                else:
                    self.log_message(error_msg, is_error=True)
                    return (index, f"API错误{status_code}")

            except Exception as e:
                self.log_message(f"请求失败: {str(e)} (尝试 {attempt + 1})", is_error=True)
                time.sleep(2 ** attempt)
                continue

        return (index, "超过最大重试次数")

    def update_token_count(self, tokens):
        self.total_tokens += tokens
        self.token_var.set(f"已用Token: {self.total_tokens}")

    def processing_worker(self):
        try:
            input_path = self.input_path.get()
            current_hash = self.calculate_file_fingerprint(input_path)
            if current_hash != self.current_input_hash:
                self.clean_checkpoint()
                self.current_write_index = 0
                self.current_input_hash = current_hash

            if input_path.lower().endswith('.csv'):
                df = pd.read_csv(input_path, dtype={self.id_column: str})
            else:
                df = pd.read_excel(
                    input_path,
                    engine='openpyxl',
                    dtype={self.id_column: str}
                )

            total = len(df)
            self.current_input_hash = current_hash

            output_path = self.output_path.get()
            write_header = not os.path.exists(output_path) or os.stat(output_path).st_size == 0

            with open(output_path, 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                if write_header:
                    writer.writerow(["ID", "原始文本", "模型响应"])

                if os.path.exists(self.checkpoint_file):
                    with open(self.checkpoint_file, 'r') as f:
                        self.current_write_index = int(f.read().strip())
                        self.log_message(f"从检查点恢复，继续处理第 {self.current_write_index} 行")

                with ThreadPoolExecutor(max_workers=self.running_config['max_workers']) as executor:
                    futures = {}
                    for idx in range(self.current_write_index, total):
                        if self.stop_event.is_set():
                            break
                        text = df.iloc[idx][self.text_column]
                        future = executor.submit(self.safe_api_call, text, idx)
                        futures[future] = idx

                    for future in as_completed(futures):
                        if self.stop_event.is_set():
                            break
                        idx = futures[future]
                        try:
                            _, result = future.result()

                            with self.lock:
                                self.results_buffer[idx] = result

                                while self.current_write_index in self.results_buffer:
                                    row = df.iloc[self.current_write_index]
                                    writer.writerow([
                                        row[self.id_column],
                                        row[self.text_column],
                                        self.results_buffer.pop(self.current_write_index)
                                    ])
                                    csvfile.flush()

                                    self.current_write_index += 1
                                    with open(self.checkpoint_file, 'w') as f:
                                        f.write(str(self.current_write_index))

                                    self.root.after(0, self.update_progress, self.current_write_index, total)

                        except Exception as e:
                            self.log_message(f"处理行 {idx} 出错: {str(e)}", is_error=True)

            if not self.stop_event.is_set():
                self.clean_checkpoint()

        except Exception as e:
            self.log_message(f"处理失败: {str(e)}", is_error=True)
        finally:
            self.root.after(0, self.processing_finished)

    def calculate_file_fingerprint(self, file_path):
        try:
            with open(file_path, 'rb') as f:
                file_hash = hashlib.sha256()
                while chunk := f.read(8192):
                    file_hash.update(chunk)
                return file_hash.hexdigest()
        except Exception as e:
            self.log_message(f"文件校验失败: {str(e)}", is_error=True)
            return ""

    def start_processing(self):
        if not self.input_path.get():
            messagebox.showerror("错误", "请选择输入文件")
            return
        if not self.output_path.get():
            messagebox.showerror("错误", "请指定输出文件")
            return

        self.stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        worker = Thread(target=self.processing_worker)
        worker.daemon = True
        worker.start()

    def stop_processing(self):
        self.stop_event.set()
        self.log_message("正在停止... 最后一条请求处理完成后将完全停止")
        self.stop_btn.config(state=tk.DISABLED)

    def processing_finished(self):
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("处理完成")

    def show_about(self):
        """显示关于信息"""
        about_info = """LLMcyberworkerV1.0
- 适配 海量大模型API调用
- 支持 自定义系统指令和用户指令
- 支持 断点续传功能
- 支持 实时显示token消耗量
- 支持 Excel/CSV输入
        """
        messagebox.showinfo("关于", about_info)


if __name__ == "__main__":
    root = tk.Tk()
    app = GenericAnalysisApp(root)
    root.mainloop()
