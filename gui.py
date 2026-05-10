# ============================================
# PERSONALIZED AI EMOTION ADAPTIVE MUSIC SYSTEM
# ============================================

import customtkinter as ctk
import tkinter as tk
from tkinter import simpledialog
import cv2
from PIL import Image, ImageTk
from keras.models import load_model
from keras.preprocessing.image import img_to_array
from win32com.client import Dispatch
import numpy as np
import os
import random
import csv
from datetime import datetime
import pyttsx3
import time
import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter

# -----------------------------
# CUSTOM TKINTER SETTINGS
# -----------------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# -----------------------------
# LOAD FACE DETECTOR
# -----------------------------
face_classifier = cv2.CascadeClassifier(
    './haarcascade_frontalface_default.xml'
)

# -----------------------------
# LOAD EMOTION MODEL
# -----------------------------
classifier = load_model('model1.keras')

# -----------------------------
# GLOBAL VARIABLES
# -----------------------------
photo = None

last_music_emotion = ""

last_play_time = 0

music_paused = False

stable_emotion = ""

emotion_start_time = 0

emotion_threshold = 3

username = ""

# -----------------------------
# MUSIC PLAYER
# -----------------------------
mp = Dispatch("WMPlayer.OCX")

# -----------------------------
# VOICE ENGINE
# -----------------------------
engine = pyttsx3.init()

# -----------------------------
# OPEN CAMERA
# -----------------------------
cap = cv2.VideoCapture(
    0,
    cv2.CAP_DSHOW
)

width = int(
    cap.get(cv2.CAP_PROP_FRAME_WIDTH)
)

height = int(
    cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
)

# -----------------------------
# CREATE WINDOW
# -----------------------------
window = ctk.CTk()

window.title(
    "AI Emotion Music Recommendation"
)

window.geometry(
    "1200x950"
)

# -----------------------------
# USER LOGIN
# -----------------------------
username = simpledialog.askstring(
    "User Login",
    "Enter Your Name:"
)

if username is None or username == "":

    username = "Guest"

# Voice Welcome
engine.say(
    f"Welcome {username}"
)

engine.runAndWait()

# -----------------------------
# PLAY MUSIC FUNCTION
# -----------------------------
def play_music(emotion):

    global last_music_emotion
    global last_play_time
    global music_paused

    if music_paused:

        return

    if emotion == last_music_emotion:

        return

    current_time = time.time()

    if current_time - last_play_time < 5:

        return

    last_play_time = current_time

    last_music_emotion = emotion

    if emotion == "Angry":

        songs_directory = "./songs/angry"

    elif emotion == "Happy":

        songs_directory = "./songs/happy"

    elif emotion == "Sad":

        songs_directory = "./songs/sad"

    elif emotion == "Neutral":

        songs_directory = "./songs/neutral"

    elif emotion == "Surprise":

        songs_directory = "./songs/surprised"

    else:

        return

    all_songs = os.listdir(
        songs_directory
    )

    random_song = random.choice(
        all_songs
    )

    song_path = os.path.join(
        songs_directory,
        random_song
    )

    status_label.configure(
        text=f"Playing: {emotion}"
    )

    mp.controls.stop()

    mp.currentPlaylist.clear()

    tune = mp.newMedia(song_path)

    mp.currentPlaylist.appendItem(tune)

    mp.controls.play()

    engine.say(
        f"You look {emotion}"
    )

    engine.runAndWait()

# -----------------------------
# SHOW ANALYTICS GRAPH
# -----------------------------
def show_analytics():

    try:

        df = pd.read_csv(
            f"{username}_emotion_log.csv",
            header=None
        )

        emotions = df[1]

        count = Counter(emotions)

        plt.figure(figsize=(8,6))

        plt.bar(
            count.keys(),
            count.values()
        )

        plt.title(
            f"{username} Emotion Analytics"
        )

        plt.xlabel(
            "Emotion"
        )

        plt.ylabel(
            "Count"
        )

        plt.show()

    except Exception as e:

        print(
            "Analytics Error:",
            e
        )

# -----------------------------
# REAL TIME UPDATE FUNCTION
# -----------------------------
def update():

    global photo
    global stable_emotion
    global emotion_start_time

    if cap.isOpened():

        ret, frame = cap.read()

        if ret:

            frame = cv2.flip(
                frame,
                1
            )

            gray = cv2.cvtColor(
                frame,
                cv2.COLOR_BGR2GRAY
            )

            faces = face_classifier.detectMultiScale(
                gray,
                scaleFactor=1.1,
                minNeighbors=3,
                minSize=(30,30)
            )

            class_labels = [
                'Angry',
                'Happy',
                'Neutral',
                'Sad',
                'Surprise'
            ]

            if len(faces) > 0:

                faces = sorted(
                    faces,
                    key=lambda x: x[2] * x[3],
                    reverse=True
                )

                (x, y, w, h) = faces[0]

                # ORANGE BOX
                cv2.rectangle(
                    frame,
                    (x, y),
                    (x+w, y+h),
                    (255,140,0),
                    3
                )

                roi_gray = gray[
                    y:y+h,
                    x:x+w
                ]

                roi_gray = cv2.resize(
                    roi_gray,
                    (48,48)
                )

                roi = roi_gray.astype(
                    "float"
                ) / 255.0

                roi = img_to_array(
                    roi
                )

                roi = np.expand_dims(
                    roi,
                    axis=0
                )

                prediction = classifier.predict(
                    roi,
                    verbose=0
                )[0]

                confidence = np.max(
                    prediction
                ) * 100

                emotion = class_labels[
                    prediction.argmax()
                ]

                label = f"{emotion} {confidence:.2f}%"

                # EMOTION LABEL
                cv2.putText(
                    frame,
                    label,
                    (x, y-10),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.7,
                    (255,140,0),
                    2
                )

                status_label.configure(
                    text=f"Detected Emotion: {label}"
                )

                # -----------------
                # STABILITY AI
                # -----------------
                if stable_emotion != emotion:

                    stable_emotion = emotion

                    emotion_start_time = time.time()

                else:

                    elapsed = time.time() - emotion_start_time

                    if elapsed >= emotion_threshold:

                        play_music(emotion)

                # -----------------
                # WELLNESS MODE
                # -----------------
                if emotion == "Sad":

                    suggestion_label.configure(
                        text="Suggestion: Relax and listen to calm music."
                    )

                elif emotion == "Angry":

                    suggestion_label.configure(
                        text="Suggestion: Try breathing exercises."
                    )

                elif emotion == "Happy":

                    suggestion_label.configure(
                        text="Great mood detected!"
                    )

                else:

                    suggestion_label.configure(
                        text="Emotion monitoring active."
                    )

                # -----------------
                # SAVE USER CSV LOG
                # -----------------
                with open(
                    f"{username}_emotion_log.csv",
                    "a",
                    newline=""
                ) as file:

                    writer = csv.writer(file)

                    writer.writerow([
                        datetime.now(),
                        emotion,
                        f"{confidence:.2f}%"
                    ])

            else:

                status_label.configure(
                    text="No Face Detected"
                )

            frame_rgb = cv2.cvtColor(
                frame,
                cv2.COLOR_BGR2RGB
            )

            img = Image.fromarray(
                frame_rgb
            )

            photo = ImageTk.PhotoImage(
                image=img
            )

            canvas.create_image(
                0,
                0,
                image=photo,
                anchor=tk.NW
            )

    window.after(30, update)

# -----------------------------
# STOP MUSIC
# -----------------------------
def stop_music():

    global last_music_emotion
    global music_paused

    try:

        music_paused = True

        mp.controls.stop()

        mp.currentPlaylist.clear()

        last_music_emotion = ""

        status_label.configure(
            text="Music Stopped"
        )

        suggestion_label.configure(
            text="AI Monitoring Paused"
        )

    except Exception as e:

        print(
            "Error stopping music:",
            e
        )

# -----------------------------
# RESUME MUSIC
# -----------------------------
def resume_music():

    global music_paused

    music_paused = False

    status_label.configure(
        text="Auto Music Resumed"
    )

    suggestion_label.configure(
        text="AI Monitoring Active"
    )

# -----------------------------
# CLOSE APP
# -----------------------------
def close():

    mp.controls.stop()

    cap.release()

    window.destroy()

# -----------------------------
# TITLE LABEL
# -----------------------------
title_label = ctk.CTkLabel(
    window,
    text="Personalized AI Emotion Adaptive Music System",
    font=("Arial", 30, "bold")
)

title_label.pack(
    pady=15
)

# -----------------------------
# USER LABEL
# -----------------------------
user_label = ctk.CTkLabel(
    window,
    text=f"Logged in User: {username}",
    font=("Arial",18)
)

user_label.pack(
    pady=5
)

# -----------------------------
# CAMERA CANVAS
# -----------------------------
canvas = tk.Canvas(
    window,
    width=width,
    height=height,
    bg="black",
    highlightthickness=0
)

canvas.pack()

# -----------------------------
# STATUS LABEL
# -----------------------------
status_label = ctk.CTkLabel(
    window,
    text="AI System Started...",
    font=("Arial", 20)
)

status_label.pack(
    pady=10
)

# -----------------------------
# SUGGESTION LABEL
# -----------------------------
suggestion_label = ctk.CTkLabel(
    window,
    text="Mental Wellness Monitoring Active",
    font=("Arial", 18)
)

suggestion_label.pack(
    pady=5
)

# -----------------------------
# BUTTON FRAME
# -----------------------------
button_frame = ctk.CTkFrame(window)

button_frame.pack(
    pady=20
)

# -----------------------------
# STOP BUTTON
# -----------------------------
btn_stop = ctk.CTkButton(
    button_frame,
    text="Stop Music",
    width=180,
    height=45,
    command=stop_music
)

btn_stop.grid(
    row=0,
    column=0,
    padx=15
)

# -----------------------------
# RESUME BUTTON
# -----------------------------
btn_resume = ctk.CTkButton(
    button_frame,
    text="Resume AI Music",
    width=180,
    height=45,
    command=resume_music
)

btn_resume.grid(
    row=0,
    column=1,
    padx=15
)

# -----------------------------
# ANALYTICS BUTTON
# -----------------------------
btn_analytics = ctk.CTkButton(
    button_frame,
    text="Show Analytics",
    width=180,
    height=45,
    command=show_analytics
)

btn_analytics.grid(
    row=0,
    column=2,
    padx=15
)

# -----------------------------
# EXIT BUTTON
# -----------------------------
btn_close = ctk.CTkButton(
    button_frame,
    text="Exit",
    width=180,
    height=45,
    fg_color="red",
    hover_color="darkred",
    command=close
)

btn_close.grid(
    row=0,
    column=3,
    padx=15
)

# -----------------------------
# START REAL TIME AI
# -----------------------------
update()

# -----------------------------
# MAIN LOOP
# -----------------------------
window.mainloop()