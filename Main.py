import os
import sys
import configparser
from datetime import datetime
from collections import Counter
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import tkinter as tk
from tkinter import (
    ttk,
    messagebox,
    Listbox,
    Frame,
    Button,
    END,
    filedialog
)
from tkcalendar import DateEntry
# 新增代码 --------------------------------------------------
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)  # EXE所在目录
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 开发环境目录
# ----------------------------------------------------------
# 全局配置
config = None
downsample_length = 0
sampling_length = 0

# ████████████████████████████████████████████████████████████████████████████████
# 配置文件相关功能
# ████████████████████████████████████████████████████████████████████████████████

def select_config_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="选择配置文件",
        filetypes=[("INI配置文件", "*.ini")],
        initialdir=BASE_DIR  # 关键修改：始终从EXE所在目录开始
    )
    root.destroy()
    return file_path

def validate_config(config):
    """验证配置文件完整性"""
    required_sections = {
        'StudentInfo': ['学院', '年级', '专业', '班级', '姓名', '性别', '民族', '学号', '电话'],
        'ParentContact': ['家长姓名', '家长电话'],
        'LeaveRequest': ['去处', '原因', '导员姓名'],
        'SystemParams': ['向下压缩长度', '取样长度']
    }
    
    missing_sections = [s for s in required_sections if s not in config]
    if missing_sections:
        raise ValueError(f"缺少必要配置节: {', '.join(missing_sections)}")
    
    for section, keys in required_sections.items():
        missing_keys = [k for k in keys if k not in config[section]]
        if missing_keys:
            raise ValueError(f"在节 [{section}] 中缺少配置项: {', '.join(missing_keys)}")

def load_config(config_path):
    """加载并验证配置文件"""
    global config, downsample_length, sampling_length
    
    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')
    validate_config(config)
    
    downsample_length = int(config['SystemParams']['向下压缩长度'])
    sampling_length = int(config['SystemParams']['取样长度'])

# ████████████████████████████████████████████████████████████████████████████████
# 图片处理核心功能
# ████████████████████████████████████████████████████████████████████████████████

def find_image_in_current_dir(prompt, default_name=None):
    """使用系统原生文件选择对话框"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title=prompt,
        filetypes=[
            ("图片文件", "*.jpg *.jpeg *.png"),
            ("所有文件", "*.*")
        ],
        initialdir=os.getcwd()
    )
    root.destroy()
    return file_path

def get_dominant_color(img, exclude_light=True, exclude_dark=True):
    """获取图片的主色调（可排除极亮/极暗色）"""
    hsv = img.convert("HSV")
    pixels = list(hsv.getdata())
    
    if exclude_light:
        pixels = [p for p in pixels if p[1] > 30 or p[2] < 220]
    if exclude_dark:
        pixels = [p for p in pixels if p[2] > 30]
    
    if not pixels:
        return (0, 0, 0)
    
    color_counts = Counter(pixels)
    dominant = color_counts.most_common(1)[0][0]
    return dominant[:3]

def create_smart_mask(img, base_color, tolerance=30):
    """创建智能蒙版（主色区域透明）"""
    hsv_img = img.convert("HSV")
    h, s, v = base_color
    
    mask = Image.new("L", img.size, 255)
    pixels = hsv_img.load()
    
    for y in range(img.height):
        for x in range(img.width):
            ph, ps, pv = pixels[x, y]
            h_diff = min(abs(ph - h), 360 - abs(ph - h))
            s_diff = abs(ps - s)
            v_diff = abs(pv - v)
            
            if h_diff < tolerance and s_diff < tolerance and v_diff < tolerance:
                similarity = 1 - max(h_diff, s_diff, v_diff)/tolerance
                mask.putpixel((x, y), int(255 * (1 - similarity**2)))

    return mask.filter(ImageFilter.GaussianBlur(1))

def prooo():
    input_file = find_image_in_current_dir("选择辅助图片") 
    if not input_file: 
        messagebox.showinfo("提示", "未选择任何图片")  # 新增提示
        return
    
    output_file = "output.jpg"
    try:
        with Image.open(os.path.join(BASE_DIR, output_file)) as output_img:
            if output_img.height <= downsample_length:
                raise ValueError(f"output图片高度不足{downsample_length}像素")
            
            cropped_output = output_img.crop((0, downsample_length, output_img.width, output_img.height))
            
            with Image.open(os.path.join(BASE_DIR, input_file)).convert("RGB") as img:
                width, height = img.size
                scale = 1080 / min(width, height)
                new_size = (int(width*scale), int(height*scale))
                img = img.resize(new_size, Image.LANCZOS)
                
                crop_width = min(img.width, 1080)
                crop_height = sampling_length 
                if img.height < crop_height:
                    print(f"错误：图片高度不足{crop_height}像素")
                    return
                
                top_part = img.crop((0, 0, crop_width, crop_height))
                dominant_color = get_dominant_color(top_part)
                mask = create_smart_mask(top_part, dominant_color)
                
                result = cropped_output.copy()
                x_offset = (cropped_output.width - crop_width) // 2
                result.paste(top_part, (x_offset, 0), mask)
                
                save_path = "output.jpg"
                result.save(save_path, quality=100)
                print(f"\n处理完成！保存为: {os.path.abspath(save_path)}")

    except Exception as e:
        print(f"处理出错: {str(e)}")

# ████████████████████████████████████████████████████████████████████████████████
# GUI 界面相关功能
# ████████████████████████████████████████████████████████████████████████████████

def resource_path(relative_path):
    """资源文件路径处理"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_font_path():
    """获取系统字体路径"""
    fonts = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
    ]
    for path in fonts:
        if os.path.exists(path):
            return path
    raise FileNotFoundError("微软雅黑字体未安装")

def create_radio_group(parent, options, default):
    """创建单选按钮组"""
    var = tk.StringVar(value=default)
    frame = ttk.Frame(parent)
    for text, value in options:
        rb = ttk.Radiobutton(
            frame,
            text=text,
            value=value,
            variable=var,
            width=6
        )
        rb.pack(side=tk.LEFT, padx=2)
    return frame, var

def create_date_picker(parent, default_date):
    """创建日期选择组件"""
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

def create_input_window(defaults):
    """创建输入窗口"""
    root = tk.Tk()
    root.title("请假条信息录入")
    root.geometry("600x450")
    
    should_generate = False
    
    container = ttk.Frame(root)
    canvas = tk.Canvas(container)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    frame = ttk.Frame(canvas)

    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    container.pack(fill="both", expand=True)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    labels = [
        "院系", "年级", "专业", "班级",
        "姓名", "学号", "性别", "民族",
        "申请日期", "联系方式", "请假类型", "是否代请假",
        "请假开始日期", "请假结束日期", "请假天数", "家长是否知晓",
        "家长姓名", "家长联系电话", "请假去向类型", "请假去向（具体地点）",
        "请假事由", "审批意见", "导员签名", "审批时间"
    ]

    entries = []
    radio_vars = {}
    comboboxes = []
    date_pickers = []
    radio_indices = []
    combobox_indices = []
    date_indices = []
    time_entry = None

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

    def on_close():
        root.destroy()
        return None

    root.protocol("WM_DELETE_WINDOW", on_close)
    
    ttk.Button(frame, text="生成", command=lambda: [setattr(sys, 'result', on_confirm())]).grid(row=12, column=0, columnspan=4, pady=10)
    
    root.mainloop()
    return getattr(sys, 'result', None)

# ████████████████████████████████████████████████████████████████████████████████
# 主程序
# ████████████████████████████████████████████████████████████████████████████████

def generate_leave_form():
    config_path = select_config_file()
    if not config_path:
        print("操作已取消")
        return
    
    try:
        load_config(config_path)
    except Exception as e:
        messagebox.showerror("配置错误", f"配置文件加载失败: {str(e)}")
        return

    text_config = [
        (355,605, config['StudentInfo']['学院'], 41, (51,51,51)),
        (355,725, config['StudentInfo']['年级'], 41, (51,51,51)),
        (355,846, config['StudentInfo']['专业'], 41, (51,51,51)),
        (355,965, config['StudentInfo']['班级'], 41, (51,51,51)),
        (355,1085, config['StudentInfo']['姓名'], 41, (51,51,51)),
        (355,1205, config['StudentInfo']['学号'], 41, (51,51,51)),
        (355,1325, config['StudentInfo']['性别'], 41, (51,51,51)),
        (355,1445, config['StudentInfo']['民族'], 41, (51,51,51)),
        (355,1565, datetime.today().strftime("%Y-%m-%d"), 41, (51,51,51)),
        (355,1685, config['StudentInfo']['电话'], 41, (51,51,51)),
        (355,1805, "事假", 41, (51,51,51)),
        (355,1925, "否", 41, (51,51,51)),
        (355,2081, datetime.today().strftime("%Y-%m-%d"), 41, (51,51,51)),
        (355,2269, datetime.today().strftime("%Y-%m-%d"), 41, (51,51,51)),
        (355,2423, "1", 41, (51,51,51)),
        (355,2579, "是", 41, (51,51,51)),
        (355,2733, config['ParentContact']['家长姓名'], 41, (51,51,51)),
        (355,2887, config['ParentContact']['家长电话'], 41, (51,51,51)),
        (355,3075, "校内", 41, (51,51,51)),
        (355,3300, config['LeaveRequest']['去处'], 41, (51,51,51)),
        (355,3488, config['LeaveRequest']['原因'], 41, (51,51,51)),
        (355,3599, "同意。", 41, (51,51,51)),
        (355,3671, config['LeaveRequest']['导员姓名'], 41, (8,116,5)),
        (355,3742, format_time(), 41, (51,51,51)),
    ]

    default_values = [x[2] for x in text_config]
    user_inputs = create_input_window(default_values)

    if user_inputs is None:
        return

    try:
        input_path = os.path.join(BASE_DIR, "input.jpg")  # 动态获取完整路径
        img = Image.open(input_path)
        draw = ImageDraw.Draw(img)
        
        for i, (x, y, _, size, color) in enumerate(text_config):
            text = user_inputs[i]
            font = ImageFont.truetype(get_font_path(), size)
            draw.text(
                (x, y - font.getmetrics()[0]),
                text,
                fill=color,
                font=font
            )
            
        output_path = os.path.join(os.getcwd(), "output.jpg")
        img.save(output_path, quality=95)
        prooo()
    except ValueError as e:
        messagebox.showerror("错误", f"日期格式错误：{str(e)}")
    except Exception as e:
        messagebox.showerror("错误", f"生成图片时出错：{str(e)}")

if __name__ == "__main__":
    generate_leave_form()
