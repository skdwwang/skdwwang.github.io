# main.py
from PIL import Image, ImageDraw, ImageFont
import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime
import configparser

# 资源路径处理
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 时间格式处理函数
def format_time(now=None):
    """格式化时间为：年(后两位)月日 时分（去除前导零）"""
    now = now or datetime.now()
    year = now.strftime("%y")
    month = str(now.month)
    day = str(now.day)
    hour = str(now.hour)
    minute = str(now.minute)
    return f"{year}年{month}月{day}日 {hour}时{minute}分"

def parse_time(time_str):
    """将格式化时间字符串解析为datetime对象"""
    try:
        date_part, time_part = time_str.split(" ")
        year = int(date_part.split("年")[0]) + 2000
        month = int(date_part.split("年")[1].split("月")[0])
        day = int(date_part.split("月")[1].split("日")[0])
        hour = int(time_part.split("时")[0])
        minute = int(time_part.split("时")[1].split("分")[0])
        return datetime(year, month, day, hour, minute)
    except Exception:
        return datetime.now()

# 配置加载
def load_config():
    config = configparser.ConfigParser()
    config_path = resource_path("config.ini")
    
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件 {config_path} 不存在")
    
    try:
        config.read(config_path, encoding='utf-8')
    except Exception as e:
        raise RuntimeError(f"配置文件读取失败: {str(e)}")
    
    required_keys = [
        '学院', '年级', '专业', '班级', '姓名','性别','民族', '学号', '电话',
        '家长姓名', '家长电话', '去处', '原因', '导员姓名'
    ]
    
    if not config.has_section('Defaults'):
        raise KeyError("配置文件中缺少 [Defaults] 节")
    
    missing = [key for key in required_keys if not config.has_option('Defaults', key)]
    if missing:
        raise KeyError(f"缺少以下必要配置项: {', '.join(missing)}")
    
    return config['Defaults']

# 字体路径获取
def get_font_path():
    fonts = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
    ]
    for path in fonts:
        if os.path.exists(path):
            return path
    raise FileNotFoundError("微软雅黑字体未安装")

# GUI组件创建函数
def create_radio_group(parent, options, default):
    var = tk.StringVar(value=default)
    frame = ttk.Frame(parent)
    for text, value in options:
        rb = ttk.Radiobutton(frame, text=text, value=value, variable=var, width=6)
        rb.pack(side=tk.LEFT, padx=2)
    return frame, var

def create_date_picker(parent, default_date):
    date_entry = DateEntry(
        parent,
        date_pattern='yyyy-mm-dd',
        width=18,
        background='darkblue',
        foreground='white',
        borderwidth=2
    )
    date_entry.set_date(default_date)
    return date_entry

# 输入窗口创建
def create_input_window(defaults):
    should_generate = False
    root = tk.Tk()
    root.title("请假条信息录入")
    root.geometry("600x450")
    
    container = ttk.Frame(root)
    canvas = tk.Canvas(container)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    frame = ttk.Frame(canvas)

    # 布局设置
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    container.pack(fill="both", expand=True)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # 输入字段定义
    labels = [
        "院系", "年级", "专业", "班级", "姓名", "学号", "性别", "民族",
        "申请日期", "联系方式", "请假类型", "是否代请假", "请假开始日期",
        "请假结束日期", "请假天数", "家长是否知晓", "家长姓名", "家长联系电话",
        "请假去向类型", "请假去向（具体地点）", "请假事由", "审批意见",
        "导员签名", "审批时间"
    ]

    entries = []
    radio_vars = {}
    comboboxes = []
    date_pickers = []
    radio_indices = []
    combobox_indices = []
    date_indices = []
    time_entry = None

    # 动态创建输入组件
    for i, (label, default) in enumerate(zip(labels, defaults)):
        row = i % 12
        col = 0 if i < 12 else 2
        
        ttk.Label(frame, text=label).grid(row=row, column=col, padx=5, pady=5, sticky=tk.E)
        
        if label in ["申请日期", "请假开始日期", "请假结束日期"]:
            date_picker = create_date_picker(frame, default)
            date_picker.grid(row=row, column=col+1, padx=5, pady=5, sticky=tk.W)
            date_pickers.append((i, date_picker))
            date_indices.append(i)
        elif label in ["性别", "是否代请假", "请假类型", "家长是否知晓"]:
            options = {
                "性别": [("男", "男"), ("女", "女")],
                "是否代请假": [("是", "是"), ("否", "否")],
                "请假类型": [("事假", "事假"), ("病假", "病假")],
                "家长是否知晓": [("是", "是"), ("否", "否")]
            }[label]
            radio_frame, var = create_radio_group(frame, options, default)
            radio_frame.grid(row=row, column=col+1, sticky=tk.W)
            radio_vars[label] = var
            radio_indices.append(i)
        elif label == "请假去向类型":
            combobox = ttk.Combobox(
                frame,
                values=["校内", "离校（在郑）", "离郑（省内）", "省外及其他区域"],
                width=18
            )
            combobox.set(default)
            combobox.grid(row=row, column=col+1, padx=5, pady=5, sticky=tk.W)
            comboboxes.append((i, combobox))
            combobox_indices.append(i)
        elif label == "审批时间":
            time_entry = ttk.Entry(frame)
            time_entry.insert(0, format_time(parse_time(default)))
            time_entry.grid(row=row, column=col+1, padx=5, pady=5, sticky=tk.W)
            entries.append((i, time_entry))
        else:
            entry = ttk.Entry(frame)
            entry.insert(0, default)
            entry.grid(row=row, column=col+1, padx=5, pady=5, sticky=tk.W)
            entries.append((i, entry))

    # 确认按钮逻辑
    def on_confirm():
        nonlocal should_generate
        should_generate = True
        result = []
        for i, default in enumerate(defaults):
            if i in date_indices:
                dp = next(dp for idx, dp in date_pickers if idx == i)
                result.append(dp.get())
            elif i in radio_indices:
                label = labels[i]
                result.append(radio_vars[label].get() or default)
            elif i in combobox_indices:
                cb = next(cb for idx, cb in comboboxes if idx == i)
                result.append(cb.get() or default)
            else:
                entry = next(ent for idx, ent in entries if idx == i)
                result.append(entry.get() or default)
        
        try:
            start_date = datetime.strptime(result[12], "%Y-%m-%d")
            end_date = datetime.strptime(result[13], "%Y-%m-%d")
            delta = end_date - start_date
            result[14] = str(delta.days + 1)
        except ValueError:
            messagebox.showerror("错误", "日期格式不正确，请检查开始和结束日期")
            return
        
        try:
            parsed_time = parse_time(result[23])
            result[23] = format_time(parsed_time)
        except Exception:
            messagebox.showerror("错误", "审批时间格式不正确，请使用'23年1月1日 9时5分'格式")
            return
        
        root.destroy()
        return result

    # 窗口关闭处理
    def on_close():
        root.destroy()
        return None

    root.protocol("WM_DELETE_WINDOW", on_close)
    ttk.Button(frame, text="生成", command=on_confirm).grid(row=12, column=0, columnspan=4, pady=10)
    
    root.mainloop()
    return getattr(sys, 'result', None)

# 主生成函数
def generate_leave_form():
    try:
        defaults = load_config()
    except Exception as e:
        messagebox.showerror("配置错误", f"加载配置失败: {str(e)}")
        return

    today = datetime.today().strftime("%Y-%m-%d")
    current_time = format_time()

    text_config = [
        (355,605, defaults['学院'], 41, (51,51,51)),
        (355,725, defaults['年级'], 41, (51,51,51)),
        (355,846, defaults['专业'], 41, (51,51,51)),
        (355,965, defaults['班级'], 41, (51,51,51)),
        (355,1085, defaults['姓名'], 41, (51,51,51)),
        (355,1205, defaults['学号'], 41, (51,51,51)),
        (355,1325, defaults['性别'], 41, (51,51,51)),
        (355,1445, defaults['民族'], 41, (51,51,51)),
        (355,1565, today, 41, (51,51,51)),
        (355,1685, defaults['电话'], 41, (51,51,51)),
        (355,1805, "事假", 41, (51,51,51)),
        (355,1925, "否", 41, (51,51,51)),
        (355,2081, today, 41, (51,51,51)),
        (355,2269, today, 41, (51,51,51)),
        (355,2423, "1", 41, (51,51,51)),
        (355,2579, "是", 41, (51,51,51)),
        (355,2733, defaults['家长姓名'], 41, (51,51,51)),
        (355,2887, defaults['家长电话'], 41, (51,51,51)),
        (355,3075, "校内", 41, (51,51,51)),
        (355,3300, defaults['去处'], 41, (51,51,51)),
        (355,3488, defaults['原因'], 41, (51,51,51)),
        (355,3599, "同意。", 41, (51,51,51)),
        (355,3671, defaults['导员姓名'], 41, (8,116,5)),
        (355,3742, current_time, 41, (51,51,51)),
    ]

    user_inputs = create_input_window([x[2] for x in text_config])

    if not user_inputs:
        print("操作已取消")
        return

    try:
        img = Image.open(resource_path("input.jpg"))
        draw = ImageDraw.Draw(img)
        font_path = get_font_path()
        
        for i, (x, y, _, size, color) in enumerate(text_config):
            text = user_inputs[i]
            font = ImageFont.truetype(font_path, size)
            draw.text((x, y - font.getmetrics()[0]), text, fill=color, font=font)
            
        output_path = os.path.join(os.getcwd(), "output.jpg")
        img.save(output_path, quality=95)
        os.startfile(output_path)
        messagebox.showinfo("成功", f"请假条已生成至:\n{output_path}")
            
    except Exception as e:
        messagebox.showerror("生成错误", f"生成失败: {str(e)}")

if __name__ == "__main__":
    generate_leave_form()
