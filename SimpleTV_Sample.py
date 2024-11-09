import pyaudio
import wave
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import deque
import time
import win32com.client
import numpy as np
import tkinter as tk
import os

voice = win32com.client.Dispatch("SAPI.SpVoice")

# 设置音频参数
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 8000
PRE_SECONDS = 1.5
POST_SECONDS = 4.5

# 设定循环次数变量
num = 9  # 录音次数限制

# 初始化 PyAudio
audio = pyaudio.PyAudio()
stream = audio.open(format=FORMAT, channels=CHANNELS,
                    rate=RATE, input=True, frames_per_buffer=CHUNK)

# GUI 设置
root = tk.Tk()
root.title("Audio Capture Interface")

# 左侧框架用于显示图形和阈值设置
left_frame = tk.Frame(root)
left_frame.pack(side=tk.LEFT, padx=10, pady=10)

# 右侧框架用于显示录音指令
right_frame = tk.Frame(root)
right_frame.pack(side=tk.RIGHT, padx=10, pady=10)

# 声音幅值阈值输入
threshold_label = tk.Label(left_frame, text="Set Threshold:")
threshold_label.pack()
threshold_entry = tk.Entry(left_frame)
threshold_entry.insert(0, "15000")
threshold_entry.pack()

# 录音次数计数器
n = 0

# 录音文件存储路径
audio_files = {}

# 检查是否有已存在的录音文件，并将路径存储在 audio_files 中
for i in range(1, num + 1):
    file_path = f'{i}_Audio_time.wav'
    if os.path.exists(file_path):
        audio_files[i] = file_path  # 如果文件已存在，添加到字典


# 开始按钮功能
def start_recording():
    global stop_program, n
    stop_program = False
    n = 0  # 计数器初始化为0
    update_audio()  # 开始音频更新


# 开始按钮
start_button = tk.Button(left_frame, text="Start", command=start_recording)
start_button.pack()

# 图形窗口
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8, 6))
canvas = FigureCanvasTkAgg(fig, master=left_frame)
canvas.get_tk_widget().pack()

# 设置图形显示
line1, = ax1.plot(np.zeros(6 * RATE))
line2, = ax2.plot(np.zeros(2 * RATE))
ax1.set_ylim([-50000, 50000])
ax1.set_xlim([0, 6 * RATE])
ax1.set_title("Real-Time Audio Signal", fontsize=12)
ax2.set_ylim([-50000, 50000])
ax2.set_xlim([0, 2 * RATE])
ax2.set_title("Last Captured Sample", fontsize=12)

# 调整图表之间的间隔
fig.subplots_adjust(hspace=0.6)

# 使用队列保存缓冲数据
plot_buffer = deque(maxlen=int(6 * RATE / CHUNK))
pre_buffer = deque(maxlen=int(PRE_SECONDS * RATE / CHUNK))
post_buffer = deque(maxlen=int(POST_SECONDS * RATE / CHUNK))

paused = False
exceeded_threshold = False
stop_program = False

# 在右侧创建文本框
labels_text = [
    "1. Power_APPLE",
    "2. Voice up_KIVI",
    "3. Voice down_DRAGON FRUIT",
    "4. Enter_ORANGE",
    "5. Back_BANANA",
    "6. Up_COCONUT",
    "7. Down_PIYATA",
    "8. Left_LEMON",
    "9. Right_RED DELICIOUS"
]
label_widgets = []


# 播放音频文件函数
def play_audio(file_path):
    if not os.path.exists(file_path):
        print("Audio file does not exist.")
        return

    # 打开音频文件
    wf = wave.open(file_path, 'rb')
    stream = audio.open(format=audio.get_format_from_width(wf.getsampwidth()),
                        channels=wf.getnchannels(),
                        rate=wf.getframerate(),
                        output=True)

    # 读取并播放音频文件
    data = wf.readframes(CHUNK)
    while data:
        stream.write(data)
        data = wf.readframes(CHUNK)

    # 关闭流和文件
    stream.stop_stream()
    stream.close()
    wf.close()


# 创建文本框并绑定点击事件
for i, text in enumerate(labels_text):
    file_path = f'{i + 1}_Audio_time.wav'
    font_style = ('Arial', 12, 'bold') if os.path.exists(file_path) else ('Arial', 12)
    label = tk.Label(right_frame, text=text, font=font_style, width=20, anchor='w', relief='ridge')
    label.pack(pady=5)

    # 绑定点击事件
    label.bind("<Button-1>", lambda event, idx=i + 1: play_audio(audio_files.get(idx, "")))
    label_widgets.append(label)


# 定义音频更新函数
def update_audio():
    global paused, exceeded_threshold, stop_program, n

    if stop_program:
        return

    # 获取用户输入的阈值
    try:
        THRESHOLD = int(threshold_entry.get())
    except ValueError:
        THRESHOLD = 15000

    # 读取音频数据
    data = stream.read(CHUNK)
    audio_data = np.frombuffer(data, dtype=np.int16)
    peak = np.max(np.abs(audio_data))

    # 更新实时图像缓冲区
    plot_buffer.append(audio_data)
    plot_data = np.hstack(list(plot_buffer))

    # 确保 plot_data 的长度为 6 * RATE
    if len(plot_data) < 6 * RATE:
        plot_data = np.pad(plot_data, (0, 6 * RATE - len(plot_data)), 'constant')
    else:
        plot_data = plot_data[:6 * RATE]

    # 更新实时图形
    line1.set_ydata(plot_data)
    canvas.draw()

    # 高亮当前的文本框
    for label in label_widgets:
        label.config(bg='white')  # 重置背景颜色
    if 0 <= n < num:
        label_widgets[n].config(bg='yellow')  # 高亮当前录音次数的文本框

    if not paused:
        pre_buffer.append(audio_data)

        # 检查峰值是否超过阈值
        if peak > THRESHOLD and not exceeded_threshold:
            exceeded_threshold = True
            paused = True
            n += 1  # 增加录音次数
            print(f"声音幅值超过阈值，开始录音阶段 {n} ...")

            # 把后3秒的数据也缓存
            for _ in range(int(POST_SECONDS * RATE / CHUNK)):
                data = stream.read(CHUNK)
                post_buffer.append(np.frombuffer(data, dtype=np.int16))

            # 合并前后音频
            combined_audio = np.hstack(list(pre_buffer) + list(post_buffer))
            first_4_sec = combined_audio[:4 * RATE]
            remaining_2_sec = combined_audio[4 * RATE:]
            min_negative_value = np.min(remaining_2_sec[remaining_2_sec < 0])
            dynamic_threshold = min_negative_value * 3.5
            below_threshold_indices = np.where(first_4_sec < dynamic_threshold)[0]

            if len(below_threshold_indices) > 0:
                start_idx = below_threshold_indices[0]
                end_idx = below_threshold_indices[-1]
                extracted_segment = first_4_sec[start_idx:end_idx + 1]
            else:
                extracted_segment = first_4_sec

            # 确保 extracted_segment_padded 长度一致
            extracted_segment_padded = np.pad(extracted_segment, (0, max(0, 2 * RATE - len(extracted_segment))),
                                              'constant')
            extracted_segment_padded = extracted_segment_padded[:2 * RATE]  # 截断到 2 * RATE 的长度
            line2.set_ydata(extracted_segment_padded)
            canvas.draw()

            # 保存音频文件
            filename = f'{n}_Audio_time.wav'
            audio_files[n] = filename  # 记录文件路径
            wf = wave.open(filename, 'wb')
            wf.setnchannels(CHANNELS)
            wf.setsampwidth(audio.get_sample_size(FORMAT))
            wf.setframerate(RATE)
            wf.writeframes(extracted_segment.tobytes())
            wf.close()

            filename = f'{n}_Audio_time.npy'
            np.save(filename, extracted_segment)

            voice.Speak(f"Sample {n} has been recorded")
            print(f"Sample {n} has been recorded")

            # 检查是否超过 num 次
            if n >= num:
                voice.Speak("Finished")
                print("Finished")
                stop_button.config(text="Finished")  # 更改按钮文本为 "Finish"
                stop_program = True  # 停止继续采集
                return

            time.sleep(2)
            exceeded_threshold = False
            paused = False
            print("继续收集实时语音")

    root.after(10, update_audio)  # 定期调用以刷新界面


# 停止按钮功能
def stop_program_func():
    global stop_program
    stop_program = True
    stream.stop_stream()
    stream.close()
    audio.terminate()
    root.quit()


# 停止按钮
stop_button = tk.Button(left_frame, text="Stop", command=stop_program_func)
stop_button.pack()

# 启动 GUI 主循环
root.mainloop()
