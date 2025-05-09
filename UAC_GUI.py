import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.simpledialog import Dialog
import pandas as pd
import os
import glob
import datetime
import openpyxl
from openpyxl.styles.builtins import styles


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("仿真配置工具")
        self.geometry("720x480")
        self.create_widgets()
        self.auto_load_config()
    def create_widgets(self):
        # 创建Notebook
        self.notebook = ttk.Notebook(self)

        # 创建各个分页
        self.node_frame = NodeSettingsFrame(self.notebook)
        self.network_frame = NetworkSettingsFrame(self.notebook)
        self.communication_frame = CommunicationSettingsFrame(self.notebook)
        self.hydrology_frame = HydrologySettingsFrame(self.notebook)

        self.notebook.add(self.node_frame, text="节点位置设置")
        self.notebook.add(self.network_frame, text="网络仿真设置")
        self.notebook.add(self.communication_frame, text="通信仿真设置")
        self.notebook.add(self.hydrology_frame, text="水文数据设置")
        self.notebook.pack(expand=True, fill='both')

        # 创建底部按钮
        button_frame = ttk.Frame(self)
        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="读取配置", command=self.load_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="检查配置", command=self.check_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="开始仿真", command=self.start_simulation).pack(side=tk.LEFT, padx=5)
        button_frame.pack(side=tk.BOTTOM, pady=10)

    def auto_load_config(self):
        try:
            files = glob.glob("*.xlsx")
            valid_files = []
            for f in files:
                try:
                    datetime.datetime.strptime(os.path.splitext(f)[0], "%Y_%m_%d_%H_%M_%S")
                    valid_files.append(f)
                except:
                    continue
            if valid_files:
                latest = max(valid_files, key=os.path.getctime)
                self.load_config(latest)
        except Exception as e:
            pass

    def save_config(self, filename=None):
        try:
            if not filename:
                filename = filedialog.asksaveasfilename(defaultextension=".xlsx")
                if not filename:
                    return

            writer = pd.ExcelWriter(filename, engine='openpyxl')

            # 保存节点设置
            node_data = {
                'center_lon': [self.node_frame.center_lon_entry.get()],
                'center_lat': [self.node_frame.center_lat_entry.get()]
            }
            pd.DataFrame(node_data).to_excel(writer, sheet_name='NodeSettings', index=False)

            # 保存节点表格
            nodes = []
            for item in self.node_frame.tree.get_children():
                nodes.append([self.node_frame.tree.item(item)['values'][i] for i in range(7)])
            pd.DataFrame(nodes, columns=self.node_frame.tree['columns']).to_excel(writer, sheet_name='NodeTable',
                                                                                  index=False)

            # 保存网络设置
            network_data = {
                '仿真总时间': [self.network_frame.entries['仿真总时间'].get()],
                '迭代间隔': [self.network_frame.entries['迭代间隔'].get()],
                '数据速率': [self.network_frame.entries['数据速率'].get()],
                '包大小': [self.network_frame.entries['包大小'].get()],
                'mac_protocol': [self.network_frame.mac_var.get()],
                'routing_protocol': [self.network_frame.routing_var.get()]
            }
            pd.DataFrame(network_data).to_excel(writer, sheet_name='NetworkSettings', index=False)

            # 保存通信设置
            comm_data = {
                'BWIndex': [self.communication_frame.bw_var.get()],
                'modOrder': [self.communication_frame.entries['modOrder'].get()],
                'codeRateIndex': [self.communication_frame.code_rate_var.get()],
                'numSymPerFrame': [self.communication_frame.entries['numSymPerFrame'].get()],
                'numFrames': [self.communication_frame.entries['numFrames'].get()],
                'fc': [self.communication_frame.entries['fc'].get()],
                'enableFading': [int(self.communication_frame.fading_var.get())],
                'chanVisual': [int(self.communication_frame.visual_var.get())],
                'enableCFO': [int(self.communication_frame.cfo_var.get())],
                'enableCPE': [int(self.communication_frame.cpe_var.get())]
            }
            pd.DataFrame(comm_data).to_excel(writer, sheet_name='CommSettings', index=False)

            # 保存水文数据
            hydrology_data = []
            for item in self.hydrology_frame.tree.get_children():
                hydrology_data.append(self.hydrology_frame.tree.item(item)['values'])
            if hydrology_data:
                # 获取表格列头
                columns = self.hydrology_frame.tree['columns']
                # 创建带列名的DataFrame
                df = pd.DataFrame(hydrology_data, columns=columns)
                # 保存时保留列头
                df.to_excel(writer, sheet_name='HydrologyData', index=False, header=True)

            writer.close()
            messagebox.showinfo("成功", "配置保存成功！")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")

    def load_config(self, filename=None):
        try:
            if not filename:
                filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
                if not filename:
                    return

            # 加载节点设置
            node_settings = pd.read_excel(filename, sheet_name='NodeSettings')
            self.node_frame.center_lon_entry.delete(0, tk.END)
            self.node_frame.center_lon_entry.insert(0, node_settings['center_lon'][0])
            self.node_frame.center_lat_entry.delete(0, tk.END)
            self.node_frame.center_lat_entry.insert(0, node_settings['center_lat'][0])

            # 加载节点表格
            node_table = pd.read_excel(filename, sheet_name='NodeTable')
            self.node_frame.tree.delete(*self.node_frame.tree.get_children())
            for _, row in node_table.iterrows():
                self.node_frame.tree.insert("", tk.END, values=list(row))

            # 加载网络设置
            network_settings = pd.read_excel(filename, sheet_name='NetworkSettings')
            for key in self.network_frame.entries:
                self.network_frame.entries[key].delete(0, tk.END)
                self.network_frame.entries[key].insert(0, str(network_settings[key][0]))
            self.network_frame.mac_var.set(network_settings['mac_protocol'][0])
            self.network_frame.routing_var.set(network_settings['routing_protocol'][0])

            # 加载通信设置
            comm_settings = pd.read_excel(filename, sheet_name='CommSettings')
            self.communication_frame.bw_var.set(str(comm_settings['BWIndex'][0]))
            self.communication_frame.code_rate_var.set(comm_settings['codeRateIndex'][0])
            for key in self.communication_frame.entries:
                self.communication_frame.entries[key].delete(0, tk.END)
                self.communication_frame.entries[key].insert(0, str(comm_settings[key][0]))
            self.communication_frame.fading_var.set(bool(comm_settings['enableFading'][0]))
            self.communication_frame.visual_var.set(bool(comm_settings['chanVisual'][0]))
            self.communication_frame.cfo_var.set(bool(comm_settings['enableCFO'][0]))
            self.communication_frame.cpe_var.set(bool(comm_settings['enableCPE'][0]))

            # 加载水文数据
            try:
                # 读取带表头的数据
                hydrology_df = pd.read_excel(filename, sheet_name='HydrologyData', header=0)

                # 清空现有数据
                self.hydrology_frame.tree.delete(*self.hydrology_frame.tree.get_children())

                # 设置列头
                columns = hydrology_df.columns.tolist()
                self.hydrology_frame.tree["columns"] = columns

                # 配置列标题和宽度
                for col in columns:
                    self.hydrology_frame.tree.heading(col, text=col)
                    self.hydrology_frame.tree.column(col, width=100, anchor='center')

                # 插入数据
                for _, row in hydrology_df.iterrows():
                    self.hydrology_frame.tree.insert("", tk.END, values=row.tolist())
                self.update()

            except Exception as e:
                print(f"水文数据加载警告: {str(e)}")  # 或使用messagebox显示非阻塞警告
                # 清空可能存在的部分数据
                self.hydrology_frame.tree.delete(*self.hydrology_frame.tree.get_children())
                self.hydrology_frame.tree["columns"] = []

            messagebox.showinfo("成功", "配置加载成功！")

        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {str(e)}")

    def check_config(self):
        # 检查节点设置
        if not self.node_frame.center_lon_entry.get() or not self.node_frame.center_lat_entry.get() :
            messagebox.showwarning("配置错误", "经纬度未填写！")
            self.notebook.select(self.node_frame)
            return False

        if len(self.node_frame.tree.get_children()) == 0:
            messagebox.showwarning("配置错误", "至少需要添加一个节点！")
            self.notebook.select(self.node_frame)
            return False

        # 检查网络设置
        for key in self.network_frame.entries:
            if not self.network_frame.entries[key].get():
                messagebox.showwarning("配置错误", f"网络设置中{key}不能为空！")
                self.notebook.select(self.network_frame)
                return False

        # 检查通信设置
        required_fields = ['modOrder', 'numSymPerFrame', 'numFrames', 'fc']
        for key in required_fields:
            if not self.communication_frame.entries[key].get():
                messagebox.showwarning("配置错误", f"通信设置中{key}不能为空！")
                self.notebook.select(self.communication_frame)
                return False
        messagebox.showinfo("检查完成", f"所有设置全部填写完成")
        return True

    def start_simulation(self):
        if self.check_config():
            filename = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S") + ".xlsx"
            self.save_config(filename)
            messagebox.showinfo("仿真开始", "配置检查通过，开始仿真...")


class NodeSettingsFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets()
        self._create_editable_tree()

    def _create_editable_tree(self):
        # 创建编辑框
        self.entry = ttk.Entry(self)
        self.entry.editing_item = None
        self.entry.editing_column = None

        # 绑定双击事件
        self.tree.bind("<Double-1>", self.on_double_click)
        self.entry.bind("<Return>", self.on_edit_confirm)
        self.entry.bind("<FocusOut>", self.on_edit_confirm)

    def on_double_click(self, event):
        # 获取点击位置
        region = self.tree.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return

        # 获取选中项和列
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)

        # 禁止编辑节点编号列（第0列）
        if column == "#1":  # Treeview列索引从#1开始
            return

        # 设置编辑框位置
        x, y, width, height = self.tree.bbox(item, column)
        self.entry.place(x=x+4, y=y+49, width=width, height=height)

        # 获取当前值
        col_index = int(column[1:]) - 1  # 转换为0-based索引
        current_value = self.tree.item(item, "values")[col_index]

        # 配置编辑框
        self.entry.delete(0, tk.END)
        self.entry.insert(0, current_value)
        self.entry.editing_item = item
        self.entry.editing_column = col_index
        self.entry.focus_set()

    def on_edit_confirm(self, event=None):
        if not self.entry.editing_item:
            return

        # 获取新值
        new_value = self.entry.get()

        # 更新tree的值
        values = list(self.tree.item(self.entry.editing_item, "values"))
        values[self.entry.editing_column] = new_value
        self.tree.item(self.entry.editing_item, values=values)

        # 清理编辑状态
        self.entry.place_forget()
        self.entry.editing_item = None
        self.entry.editing_column = None

    def create_widgets(self):
        # 中心经纬度设置
        center_frame = ttk.Frame(self)
        ttk.Label(center_frame, text="中心经度：").pack(side=tk.LEFT, padx=5)
        self.center_lon_entry = ttk.Entry(center_frame)
        self.center_lon_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(center_frame, text="中心纬度：").pack(side=tk.LEFT, padx=5)
        self.center_lat_entry = ttk.Entry(center_frame)
        self.center_lat_entry.pack(side=tk.LEFT, padx=5)
        center_frame.pack(pady=10)

        # 节点表格
        columns = ("节点编号", "x坐标", "y坐标", "z坐标", "x轴速度", "y轴速度", "z轴速度")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=20, anchor='center')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 按钮
        btn_frame = ttk.Frame(self)
        ttk.Button(btn_frame, text="添加节点", command=self.add_node).pack(pady=5)
        ttk.Button(btn_frame, text="删除节点", command=self.delete_node).pack(pady=5)
        btn_frame.pack(side=tk.RIGHT, padx=5)

    def add_node(self):
        next_id = len(self.tree.get_children())
        self.tree.insert("", tk.END, values=(next_id, "", "", "", "", "", ""))

    def delete_node(self):
        selected = self.tree.selection()
        if selected:
            self.tree.delete(selected)
            # 重新编号
            for idx, item in enumerate(self.tree.get_children()):
                self.tree.item(item, values=(idx, *self.tree.item(item)['values'][1:]))


class NetworkSettingsFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.entries = {}
        self.create_widgets()

    def create_widgets(self):
        fields = ["仿真总时间", "迭代间隔", "数据速率", "包大小"]
        for i, field in enumerate(fields):
            frame = ttk.Frame(self)
            ttk.Label(frame, text=field + ":").pack(side=tk.LEFT)
            entry = ttk.Entry(frame)
            entry.pack(side=tk.LEFT, padx=5)
            self.entries[field] = entry
            frame.pack(pady=5)

        # MAC协议选择
        mac_frame = ttk.Frame(self)
        ttk.Label(mac_frame, text="MAC协议:").pack(side=tk.LEFT)
        self.mac_var = tk.StringVar()
        ttk.Combobox(mac_frame, textvariable=self.mac_var, values=["Aloha", "Jamming"], state="readonly").pack(
            side=tk.LEFT, padx=5)
        mac_frame.pack(pady=5)

        # 路由协议选择
        routing_frame = ttk.Frame(self)
        ttk.Label(routing_frame, text="路由协议:").pack(side=tk.LEFT)
        self.routing_var = tk.StringVar()
        ttk.Combobox(routing_frame, textvariable=self.routing_var, values=["Dummy"], state="readonly").pack(
            side=tk.LEFT, padx=5)
        routing_frame.pack(pady=5)


class CommunicationSettingsFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.entries = {}
        self.create_widgets()

    def create_widgets(self):
        # BWIndex
        bw_frame = ttk.Frame(self)
        ttk.Label(bw_frame, text="BWIndex:").pack(side=tk.LEFT)
        self.bw_var = tk.StringVar()
        ttk.Combobox(bw_frame, textvariable=self.bw_var, values=list(range(1, 8)), state="readonly").pack(side=tk.LEFT,
                                                                                                          padx=5)
        bw_frame.pack(pady=5)

        # modOrder
        mod_frame = ttk.Frame(self)
        ttk.Label(mod_frame, text="modOrder:").pack(side=tk.LEFT)
        self.entries['modOrder'] = ttk.Entry(mod_frame)
        self.entries['modOrder'].pack(side=tk.LEFT, padx=5)
        mod_frame.pack(pady=5)

        # codeRateIndex
        code_frame = ttk.Frame(self)
        ttk.Label(code_frame, text="codeRateIndex:").pack(side=tk.LEFT)
        self.code_rate_var = tk.StringVar()
        ttk.Combobox(code_frame, textvariable=self.code_rate_var,
                     values=["1/2", "2/3", "3/4", "5/6"], state="readonly").pack(side=tk.LEFT, padx=5)
        code_frame.pack(pady=5)

        # 其他输入字段
        fields = ["numSymPerFrame", "numFrames", "fc"]
        for field in fields:
            frame = ttk.Frame(self)
            ttk.Label(frame, text=field + ":").pack(side=tk.LEFT)
            self.entries[field] = ttk.Entry(frame)
            self.entries[field].pack(side=tk.LEFT, padx=5)
            frame.pack(pady=5)

        # 勾选框
        check_frame = ttk.Frame(self)
        self.fading_var = tk.BooleanVar()
        ttk.Checkbutton(check_frame, text="enableFading", variable=self.fading_var).pack(side=tk.LEFT, padx=5)
        self.visual_var = tk.BooleanVar()
        ttk.Checkbutton(check_frame, text="chanVisual", variable=self.visual_var).pack(side=tk.LEFT, padx=5)
        self.cfo_var = tk.BooleanVar()
        ttk.Checkbutton(check_frame, text="enableCFO", variable=self.cfo_var).pack(side=tk.LEFT, padx=5)
        self.cpe_var = tk.BooleanVar()
        ttk.Checkbutton(check_frame, text="enableCPE", variable=self.cpe_var).pack(side=tk.LEFT, padx=5)
        check_frame.pack(pady=5)


class HydrologySettingsFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets()

    def create_widgets(self):
        btn = ttk.Button(self, text="选择水文数据文件", command=self.load_csv)
        btn.pack(pady=10)

        self.tree = ttk.Treeview(self, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def load_csv(self):
        filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filepath:
            try:
                df = pd.read_csv(filepath,encoding='ANSI')
                self.tree.delete(*self.tree.get_children())
                self.tree["columns"] = list(df.columns)
                for col in df.columns:
                    self.tree.heading(col, text=col)
                    self.tree.column(col, width=100, anchor='center')
                for _, row in df.iterrows():
                    self.tree.insert("", tk.END, values=list(row))
            except Exception as e:
                messagebox.showerror("错误", f"读取文件失败: {str(e)}")


if __name__ == "__main__":
    app = Application()
    app.mainloop()