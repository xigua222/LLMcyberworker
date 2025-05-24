import csv
import json
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import requests
import time
import os
import hashlib
import re
from datetime import datetime
from urllib.parse import urlparse
from threading import Thread, Event

DEFAULT_CONFIG = {
    "MAX_RETRIES": 7,
    "MAX_TEXT_LENGTH": 4000,
    "MAX_WORKERS": 30,
    "REQUESTS_PER_SECOND": 5,
    "AUTO_SAVE_INTERVAL": 30,
    "CHECKPOINT_FILE": ".analysis_checkpoint.tmp",
    "MAX_REASON_LENGTH": 200,
    "RATE_BUCKET_CAPACITY": 17
}

SENTIMENT_MAP = {
    "正面": 1,
    "中性": 0,
    "负面": -1,
    "无相关信息": 0
}


class AnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("大模型接口分析")

        self.api_url = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
        self.api_key = "ccafe75a420116ea842794ea2797525a.cMMgUrnZpvzHdSZq"
        self.system_prompt = """以下文本是一家上市公司业绩说明会记录的高管回答。你是一名经济学家。仅根据此文本，回答问题。 
    宏观经济认知的分析维度：  
        - 管理层明示或可能暗示表述的宏观经济环境判断
        - 对国家政策、市场、行业的趋势感知、预测和分析
        - 经济指标（GDP、CPI等）的预期表述
    请严格按以下规则处理：
        1. 分类选项（程度递增）：负面、中性、正面、无相关信息。
        2. 输出格式：分类结果: 解释原因（50字内）
    分类请参考以下示例：
        - 负面：涉及宏观经济环境，出现类似“行业调控、趋紧的融资环境、需求不明朗、不确定性大、面临着极大的挑战、较大负面影响、更加困难”等字眼
        - 正面：涉及宏观经济环境，出现类似“巨大机遇、经济上行、更有信心、看好经济、高景气增长、增速均在两位数以上、政策利好、国家战略”等字眼
        """
        self.user_prompt = "该回答对于宏观经济状态认知态度如何（注意：不是公司自身经营状况，不要因为企业困难反向过度推断和联想），请严格按照以下输出格式：分类结果: 解释原因：\n{text}"
        self.stop_event = Event()
        self.current_index = 0
        self.df = None
        self.model_name ="glm-4-air"

        self.id_column = "uid"
        self.year_column = "year"
        self.text_column = "acntet"

        self.setup_ui()
        self.setup_menu()
        self.update_api_headers()
        self.rate_limiter = threading.BoundedSemaphore(DEFAULT_CONFIG["RATE_BUCKET_CAPACITY"])
        self.last_request_time = 0

        self.current_write_index = 0
        self.checkpoint_file = ".progress_checkpoint"
        self.csv_file = None
        self.writer = None
        self.results_buffer = {}
        self.lock = threading.Lock()
        self.current_input_hash = ""
        self.checkpoint_file = os.path.abspath(".progress_checkpoint")

    def update_api_headers(self):
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
            "Accept": "application/json"
        }

    def setup_menu(self):
        menubar = tk.Menu(self.root)

        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="模型配置", command=self.configure_model)
        config_menu.add_command(label="列名配置", command=self.configure_columns)
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
        self.start_btn = ttk.Button(btn_frame, text="开始分析", command=self.start_analysis)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(btn_frame, text="停止分析", command=self.stop_analysis, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=10)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=4, column=0, columnspan=3, sticky=tk.W)
        self.log_text = tk.Text(main_frame, height=10, width=70)
        self.log_text.grid(row=5, column=0, columnspan=3, pady=10)
        self.log_text.tag_config('error', foreground='red')
        prompt_frame = ttk.LabelFrame(main_frame, text="当前分析指令")
        prompt_frame.grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=5)
        ttk.Label(prompt_frame, text="系统角色:").grid(row=0, column=0, sticky=tk.W)
        self.system_prompt_entry = tk.Text(prompt_frame, height=3, width=60, state=tk.DISABLED)
        self.system_prompt_entry.grid(row=0, column=1)
        ttk.Label(prompt_frame, text="分析指令:").grid(row=1, column=0, sticky=tk.W)
        self.user_prompt_entry = tk.Text(prompt_frame, height=3, width=60, state=tk.DISABLED)
        self.user_prompt_entry.grid(row=1, column=1)
        self.edit_btn = ttk.Button(prompt_frame, text="编辑指令", command=self.toggle_prompt_edit)
        self.edit_btn.grid(row=0, column=2, rowspan=2, padx=5)
        self.refresh_prompt_display()

    def refresh_prompt_display(self):
        self.system_prompt_entry.config(state=tk.NORMAL)
        self.user_prompt_entry.config(state=tk.NORMAL)
        self.system_prompt_entry.delete(1.0, tk.END)
        self.user_prompt_entry.delete(1.0, tk.END)
        self.system_prompt_entry.insert(tk.END, self.system_prompt)
        self.user_prompt_entry.insert(tk.END, self.user_prompt)
        self.system_prompt_entry.config(state=tk.DISABLED)
        self.user_prompt_entry.config(state=tk.DISABLED)

    def configure_model(self):
        config_dialog = tk.Toplevel(self.root)
        config_dialog.title("模型配置")

        ttk.Label(config_dialog, text="API地址:").grid(row=0, column=0, padx=5, pady=5)
        api_url_entry = ttk.Entry(config_dialog, width=40)
        api_url_entry.grid(row=0, column=1, padx=5, pady=5)
        api_url_entry.insert(0, self.api_url)

        ttk.Label(config_dialog, text="API密钥:").grid(row=1, column=0, padx=5, pady=5)
        api_key_entry = ttk.Entry(config_dialog, width=40)
        api_key_entry.grid(row=1, column=1, padx=5, pady=5)
        api_key_entry.insert(0, self.api_key)

        ttk.Label(config_dialog, text="模型名称:").grid(row=2, column=0, padx=5, pady=5)
        model_entry = ttk.Entry(config_dialog, width=40)
        model_entry.grid(row=2, column=1, padx=5, pady=5)
        model_entry.insert(0, self.model_name)

        def save_config():
            new_url = api_url_entry.get().strip()
            if not self.validate_url(new_url):
                messagebox.showerror("错误", "无效的API地址格式，请使用http/https开头的完整URL")
                return

            self.api_url = new_url
            self.api_key = api_key_entry.get().strip()
            self.model_name = model_entry.get().strip()
            self.update_api_headers()
            config_dialog.destroy()
            messagebox.showinfo("成功", "模型配置已更新")

        ttk.Button(config_dialog, text="保存", command=save_config).grid(row=3, column=1, pady=10)

    def validate_url(self, url):
        try:
            result = urlparse(url)
            return all([result.scheme in ["http", "https"], result.netloc])
        except:
            return False

    def configure_columns(self):
        config_dialog = tk.Toplevel(self.root)
        config_dialog.title("列名配置")

        ttk.Label(config_dialog, text="公司代码列名:").grid(row=0, column=0, padx=5, pady=5)
        id_entry = ttk.Entry(config_dialog, width=20)
        id_entry.grid(row=0, column=1, padx=5, pady=5)
        id_entry.insert(0, self.id_column)

        ttk.Label(config_dialog, text="年份列名:").grid(row=1, column=0, padx=5, pady=5)
        year_entry = ttk.Entry(config_dialog, width=20)
        year_entry.grid(row=1, column=1, padx=5, pady=5)
        year_entry.insert(0, self.year_column)

        ttk.Label(config_dialog, text="文本列名:").grid(row=2, column=0, padx=5, pady=5)
        text_entry = ttk.Entry(config_dialog, width=20)
        text_entry.grid(row=2, column=1, padx=5, pady=5)
        text_entry.insert(0, self.text_column)

        def save_columns():
            self.id_column = id_entry.get()
            self.year_column = year_entry.get()
            self.text_column = text_entry.get()
            config_dialog.destroy()
            messagebox.showinfo("成功", "列名配置已更新")

        ttk.Button(config_dialog, text="保存", command=save_columns).grid(row=3, column=1, pady=10)

    def select_input_file(self):
        old_path = self.input_path.get()
        path = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            new_hash = self.calculate_file_fingerprint(path)
            if path != old_path or new_hash != self.current_input_hash:
                self.clean_checkpoint()

            self.input_path.set(path)
            self.current_input_hash = new_hash
            self.log_message(f"输入文件已选择: {path}")

    def calculate_file_fingerprint(self, file_path):
        try:
            stat = os.stat(file_path)
            return f"{stat.st_size}_{stat.st_mtime_ns}"
        except:
            return ""

    def clean_checkpoint(self):
        if os.path.exists(self.checkpoint_file):
            try:
                os.remove(self.checkpoint_file)
                self.log_message("检测到新文件，已清除旧检查点")
            except Exception as e:
                self.log_message(f"检查点清理失败: {str(e)}", is_error=True)

    def select_output_file(self):
        path = filedialog.asksaveasfilename(
            title="保存结果文件",
            defaultextension=".csv",
            filetypes=[("Excel文件", "*.csv")]
        )
        if path:
            self.output_path.set(path)
            self.log_message(f"输出路径已设置: {path}")

    def toggle_prompt_edit(self):
        if self.system_prompt_entry['state'] == tk.DISABLED:
            self.system_prompt_entry.config(state=tk.NORMAL)
            self.user_prompt_entry.config(state=tk.NORMAL)
            self.edit_btn.config(text="保存指令")
        else:
            self.system_prompt = self.system_prompt_entry.get("1.0", tk.END).strip()
            self.user_prompt = self.user_prompt_entry.get("1.0", tk.END).strip()
            self.refresh_prompt_display()
            self.edit_btn.config(text="编辑指令")
            self.log_message("分析指令已更新")

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

    def get_data_fingerprint(self):
        return hashlib.md5(pd.util.hash_pandas_object(self.df).values).hexdigest()

    def load_checkpoint(self):
        if os.path.exists(DEFAULT_CONFIG['CHECKPOINT_FILE']):
            try:
                checkpoint = pd.read_pickle(DEFAULT_CONFIG['CHECKPOINT_FILE'])
                if checkpoint['data_fingerprint'] == self.get_data_fingerprint():
                    self.current_index = checkpoint['current_index']
                    self.df['sentiment_score'] = checkpoint['scores']
                    self.df['sentiment_reason'] = checkpoint.get('reasons', [""] * len(self.df))
                    self.log_message(f"检测到断点，将从第 {self.current_index + 1} 条继续")
                    return True
            except Exception as e:
                self.log_message(f"加载断点失败: {str(e)}", is_error=True)
        return False

    def save_checkpoint(self):
        checkpoint_data = {
            'current_index': self.current_write_index,
            'input_hash': self.current_input_hash,
            'file_path': self.input_path.get()
        }

        try:
            with tempfile.NamedTemporaryFile(mode='w', delete=False, encoding='utf-8') as tf:
                json.dump(checkpoint_data, tf)
                tf.flush()
                os.fsync(tf.fileno())
            os.replace(tf.name, self.checkpoint_file)
        except Exception as e:
            self.log_message(f"检查点保存失败: {str(e)}", is_error=True)

    def load_checkpoint(self):
        if not os.path.exists(self.checkpoint_file):
            return False

        try:
            with open(self.checkpoint_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            current_hash = self.calculate_file_fingerprint(self.input_path.get())
            if (data['file_path'] == self.input_path.get() and
                    data['input_hash'] == current_hash and
                    os.path.exists(self.input_path.get())):
                self.current_write_index = data['current_index']
                self.log_message(f"从检查点恢复，将继续从第 {self.current_write_index} 行开始")
                return True

            self.log_message("检查点与当前文件不匹配，将重新开始")
            return False

        except Exception as e:
            self.log_message(f"加载检查点失败: {str(e)}", is_error=True)
            return False

    def safe_api_call(self, text, index):
        if self.stop_event.is_set():
            return (index, 0, "已中止")

        if pd.isna(text) or str(text).strip() == "":
            return (index, 0, "空输入")

        clean_text = str(text).strip()[:DEFAULT_CONFIG['MAX_TEXT_LENGTH']]
        payload = {
            "messages": [
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": self.user_prompt.format(text=clean_text)}
            ],
            "model": self.model_name,
            "temperature": 0.1,
            "top_p": 0.9,
            "frequency_penalty": 0.2,
            "max_tokens": 100
        }

        for attempt in range(DEFAULT_CONFIG['MAX_RETRIES']):
            try:
                with self.rate_limiter:
                    now = time.time()
                    elapsed = now - self.last_request_time
                    interval = 1.0 / DEFAULT_CONFIG['REQUESTS_PER_SECOND']

                    if elapsed < interval:
                        sleep_time = interval - elapsed
                        time.sleep(sleep_time)

                    self.last_request_time = time.time()

                response = requests.post(
                    self.api_url,
                    headers=self.headers,
                    json=payload,
                    timeout=(10, 30)
                )
                response.raise_for_status()

                raw_result = response.json()['choices'][0]['message']['content'].strip()
                score, reason = self.parse_api_response(raw_result)
                return (index, score, reason)

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
                    return (index, 0, f"API错误{status_code}")

            except requests.exceptions.Timeout:
                self.log_message(f"请求超时，等待{2 ** attempt}秒重试... (尝试 {attempt + 1})", is_error=True)
                time.sleep(2 ** attempt)
                continue

            except Exception as e:
                self.log_message(f"请求失败: {str(e)}", is_error=True)
                if attempt < DEFAULT_CONFIG['MAX_RETRIES'] - 1:
                    time.sleep(2 ** attempt)
                continue

        return (index, 0, "超过重试次数")

    def parse_api_response(self, raw_text):
        clean_text = raw_text.replace("\n", " ").strip()
        if ":" in clean_text:
            parts = clean_text.split(":", 1)
            if len(parts) == 2:
                category, reason = parts[0].strip(), parts[1].strip()
                if category in SENTIMENT_MAP:
                    return SENTIMENT_MAP[category], reason[:DEFAULT_CONFIG['MAX_REASON_LENGTH']]

        pattern = r"(" + "|".join(SENTIMENT_MAP.keys()) + ")"
        match = re.search(pattern, clean_text)
        if match:
            category = match.group(0)
            reason = clean_text.replace(category, "").strip(" :")
            return SENTIMENT_MAP[category], reason[:DEFAULT_CONFIG['MAX_REASON_LENGTH']]
        for category in SENTIMENT_MAP:
            if category in clean_text:
                reason = clean_text.replace(category, "").strip(" :")
                return SENTIMENT_MAP[category], reason[:DEFAULT_CONFIG['MAX_REASON_LENGTH']]
        return 0, "无法解析响应格式"

    def analysis_worker(self):
        try:
            self.df = pd.read_excel(
                self.input_path.get(),
                engine='openpyxl',
                dtype={self.id_column: str, self.year_column: int}
            )
            total = len(self.df)
            self.current_input_hash = self.calculate_file_fingerprint(self.input_path.get())
            if not self.load_checkpoint():
                self.current_write_index = 0
            output_path = self.output_path.get()
            write_header = not os.path.exists(output_path) or os.stat(output_path).st_size == 0

            with open(output_path, 'a', newline='', encoding='utf-8') as csvfile:
                self.writer = csv.writer(csvfile)
                if write_header:
                    self.writer.writerow(["公司代码", "年份", "原始文本", "情感得分", "分析依据"])
                if os.path.exists(self.checkpoint_file):
                    with open(self.checkpoint_file, 'r') as f:
                        self.current_write_index = int(f.read().strip())
                        self.log_message(f"从检查点恢复，将继续从第 {self.current_write_index} 行开始")

                with ThreadPoolExecutor(max_workers=DEFAULT_CONFIG['MAX_WORKERS']) as executor:
                    futures = {}
                    for idx in range(self.current_write_index, total):
                        if self.stop_event.is_set():
                            break
                        text = self.df.iloc[idx][self.text_column]
                        future = executor.submit(self.safe_api_call, text, idx)
                        futures[future] = idx

                    for future in as_completed(futures):
                        if self.stop_event.is_set():
                            break
                        idx = futures[future]
                        try:
                            _, score, reason = future.result()
                            with self.lock:
                                self.results_buffer[idx] = (score, reason)
                                while self.current_write_index in self.results_buffer:
                                    score, reason = self.results_buffer.pop(self.current_write_index)
                                    row = self.df.iloc[self.current_write_index]
                                    self.writer.writerow([
                                        row[self.id_column],
                                        row[self.year_column],
                                        row[self.text_column],
                                        score,
                                        reason[:DEFAULT_CONFIG['MAX_REASON_LENGTH']]]
                                    )
                                    csvfile.flush()
                                    self.current_write_index += 1
                                    with open(self.checkpoint_file, 'w') as f:
                                        f.write(str(self.current_write_index))
                                    self.root.after(0, self.update_progress, self.current_write_index, total)

                        except Exception as e:
                            self.log_message(f"处理行 {idx} 出错: {str(e)}", is_error=True)

            if not self.stop_event.is_set() and os.path.exists(self.checkpoint_file):
                try:
                    os.remove(self.checkpoint_file)
                except Exception as e:
                    self.log_message(f"删除检查点失败: {str(e)}（可手动删除）", is_error=True)

        except Exception as e:
            self.log_message(f"致命错误: {str(e)}", is_error=True)
        finally:
            self.root.after(0, self.generate_summary)
            self.root.after(0, self.analysis_finished)

    def start_analysis(self):
        if not self.input_path.get():
            messagebox.showerror("错误", "请先选择输入文件")
            return
        if not self.output_path.get():
            messagebox.showerror("错误", "请先指定输出文件")
            return
        current_hash = self.calculate_file_fingerprint(self.input_path.get())
        if current_hash != self.current_input_hash:
            self.clean_checkpoint()
            self.current_write_index = 0
        self.stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.last_save_time = time.time()

        worker = Thread(target=self.analysis_worker)
        worker.daemon = True
        worker.start()

    def stop_analysis(self):
        self.stop_event.set()
        if hasattr(self, 'csvfile') and not self.csvfile.closed:
            try:
                self.csvfile.flush()
                os.fsync(self.csvfile.fileno())
                self.csvfile.close()
            except Exception as e:
                self.log_message(f"文件关闭失败: {str(e)}", is_error=True)
        if self.current_write_index > 0:
            self.log_message(f"分析已暂停，下次可从第 {self.current_write_index} 行继续")
        else:
            self.log_message("分析已停止")

    def analysis_finished(self):
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")

    def generate_summary(self):
        try:
            raw_df = pd.read_csv(self.output_path.get())
            summary_df = raw_df.groupby(["公司代码", "年份"]).agg(
                宏观认知指数=("情感得分", "sum"),
                有效文本数量=("情感得分", "count"),
                有关文本数量=("情感得分", lambda x: (x != 0).sum()),
                典型分析依据=("分析依据", lambda x: x.mode()[0] if not x.mode().empty else "")
            ).reset_index()
            summary_path = os.path.splitext(self.output_path.get())[0] + "_汇总.csv"
            summary_df.to_csv(summary_path, index=False)
            self.log_message(f"汇总统计结果已保存至 {summary_path}")
        except Exception as e:
            self.log_message(f"生成汇总失败: {str(e)}", is_error=True)

    def show_about(self):
        about_text = """
        版本号：20250221-6.7
        改进 写入算法 彻底避免处理过大文件时内存占用问题
        新增 实时结果 支持任意时间节点暂停后查看原始文件
        新增 文件指纹 大幅改进断点接力体验
        修复 一系列bug
        ---
        此前版本
        新增 应用基础功能
        新增 用户界面（GUI)
        新增 自定义模型接口，基本适配市面主流云服务
        新增 令牌桶机制，合理控制api调用速率，大幅降低报错
        新增 prompt自定义能力
        新增 断点接力 支持中断后在断点继续任务
        """
        messagebox.showinfo("关于", about_text)


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = AnalysisApp(root)
        root.mainloop()
    except Exception as e:
        print(f"启动失败: {str(e)}")
        input("按回车退出...")
