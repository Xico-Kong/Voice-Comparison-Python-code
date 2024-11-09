import pyaudio
import wave
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import deque
import time
import win32com.client
import librosa
import numpy as np
from fastdtw import fastdtw
from scipy.spatial.distance import euclidean
import tkinter as tk
import os

# Initialize voice
voice = win32com.client.Dispatch("SAPI.SpVoice")

# Set audio parameters
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 8000
PRE_SECONDS = 1.5
POST_SECONDS = 4.5

# Initialize PyAudio
audio = pyaudio.PyAudio()

# GUI setup
root = tk.Tk()
root.title("Voice Command Recognition")

# Left frame for threshold input and audio display
left_frame = tk.Frame(root)
left_frame.pack(side=tk.LEFT, padx=10, pady=10)

# Right frame for command labels
right_frame = tk.Frame(root)
right_frame.pack(side=tk.RIGHT, padx=10, pady=10)

# Set threshold input
threshold_label = tk.Label(left_frame, text="Set Threshold:")
threshold_label.pack()
threshold_entry = tk.Entry(left_frame)
threshold_entry.insert(0, "15000")
threshold_entry.pack()


# Start and Stop buttons
def start_recording():
    global stop_program
    stop_program = False
    update_audio()


start_button = tk.Button(left_frame, text="Start", command=start_recording)
start_button.pack()

# Place Stop button at the bottom
stop_button = tk.Button(left_frame, text="Stop", command=root.quit)
stop_button.pack(side=tk.BOTTOM)

# Command labels and audio association
commands = [
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
command_labels = []
similarity_labels = []
audio_files = {}  # Dictionary to hold audio file paths if they exist


# Function to play audio
def play_audio(file_path):
    if not os.path.exists(file_path):
        print("Audio file does not exist.")
        return
    wf = wave.open(file_path, 'rb')
    stream = audio.open(format=audio.get_format_from_width(wf.getsampwidth()),
                        channels=wf.getnchannels(),
                        rate=wf.getframerate(),
                        output=True)
    data = wf.readframes(CHUNK)
    while data:
        stream.write(data)
        data = wf.readframes(CHUNK)
    stream.stop_stream()
    stream.close()
    wf.close()


# Create command labels
for i, command in enumerate(commands):
    file_path = f'{i + 1}_Audio_time.wav'
    font_style = ('Arial', 12, 'bold') if os.path.exists(file_path) else ('Arial', 12)

    # Command label with similarity placeholder
    command_frame = tk.Frame(right_frame)
    command_label = tk.Label(command_frame, text=command, font=font_style, width=20, anchor='w', relief='ridge')
    command_label.pack(side=tk.LEFT)
    similarity_label = tk.Label(command_frame, text="Difference: N/A", font=('Arial', 10))
    similarity_label.pack(side=tk.RIGHT, padx=5)
    command_frame.pack(pady=5)

    command_labels.append(command_label)
    similarity_labels.append(similarity_label)

    # If audio file exists, store its path and bind click event to play audio
    if os.path.exists(file_path):
        audio_files[i] = file_path
        command_label.bind("<Button-1>", lambda event, idx=i: play_audio(audio_files[idx]))


# Extract MFCC
def extract_mfcc(y, sr, n_mfcc=13):
    y = y.astype(np.float32)
    n_fft = min(2048, len(y) // 3)
    hop_length = n_fft // 3
    return librosa.feature.mfcc(y=y, sr=sr, n_mfcc=n_mfcc, n_fft=n_fft, hop_length=hop_length)


# Compare MFCC with DTW
def compare_mfcc_dtw(mfcc1, mfcc2):
    distance, _ = fastdtw(mfcc1.T, mfcc2.T, dist=euclidean)
    return distance


# Initialize real-time audio display
fig, ax = plt.subplots()
canvas = FigureCanvasTkAgg(fig, master=left_frame)
canvas.get_tk_widget().pack()
line, = ax.plot(np.zeros(6 * RATE))
ax.set_ylim([-50000, 50000])
ax.set_xlim([0, 6 * RATE])
ax.set_xlabel("Samples")
ax.set_ylabel("Amplitude")

# Buffer for real-time audio
plot_buffer = deque(maxlen=int(6 * RATE / CHUNK))
pre_buffer = deque(maxlen=int(PRE_SECONDS * RATE / CHUNK))
post_buffer = deque(maxlen=int(POST_SECONDS * RATE / CHUNK))
paused = False
exceeded_threshold = False
stop_program = False


# Update audio function
def update_audio():
    global paused, exceeded_threshold, stop_program

    if stop_program:
        return

    try:
        THRESHOLD = int(threshold_entry.get())
    except ValueError:
        THRESHOLD = 15000

    # Start audio stream
    stream = audio.open(format=FORMAT, channels=CHANNELS, rate=RATE, input=True, frames_per_buffer=CHUNK)
    data = stream.read(CHUNK)
    audio_data = np.frombuffer(data, dtype=np.int16)
    peak = np.max(np.abs(audio_data))

    # Update plot buffer
    plot_buffer.append(audio_data)
    plot_data = np.hstack(list(plot_buffer))
    if len(plot_data) < 6 * RATE:
        plot_data = np.pad(plot_data, (0, 6 * RATE - len(plot_data)), 'constant')
    line.set_ydata(plot_data[:6 * RATE])
    canvas.draw()

    if not paused:
        pre_buffer.append(audio_data)

        # Check if audio threshold is exceeded
        if peak > THRESHOLD and not exceeded_threshold:
            exceeded_threshold = True
            paused = True

            # Buffer post-audio data
            for _ in range(int(POST_SECONDS * RATE / CHUNK)):
                data = stream.read(CHUNK)
                post_buffer.append(np.frombuffer(data, dtype=np.int16))

            # Process audio data
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

            # Load sample audio data
            samples = [np.load(f'{i}_Audio_time.npy').astype(np.float32) for i in range(1, 10)]
            extracted_segment = extracted_segment.astype(np.float32)
            mfcc_record = extract_mfcc(extracted_segment, RATE)

            # Compute similarities
            similarities = [compare_mfcc_dtw(mfcc_record, extract_mfcc(sample, RATE)) for sample in samples]
            min_similarity_idx = np.argmin(similarities)
            min_similarity_value = similarities[min_similarity_idx]

            # Display similarities and highlight command
            for idx, (label, similarity) in enumerate(zip(similarity_labels, similarities)):
                label.config(text=f"Difference: {similarity/1750:.2f}")
                command_labels[idx].config(bg="white")
            command_labels[min_similarity_idx].config(bg="yellow")

            # Speak command or indicate unrecognized command
            if min_similarity_value > 1750:
                voice.Speak("I can not understand")
                print("I can not understand")
            else:
                voice.Speak(commands[min_similarity_idx].split(". ")[1])

            time.sleep(2)
            exceeded_threshold = False
            paused = False

    root.after(10, update_audio)


root.mainloop()

