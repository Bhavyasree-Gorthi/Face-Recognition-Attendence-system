from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

# ===================== Speak Function =====================
def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

# ===================== Load Face Detection Model =====================
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default2.xml')

# ===================== Load Face Data and Labels =====================
with open('data/names.pkl', 'rb') as name_file:
    LABELS = pickle.load(name_file)

with open('data/faces_data.pkl', 'rb') as face_file:
    FACES = pickle.load(face_file)

print('Shape of Faces matrix -->', FACES.shape)

# ===================== Train KNN =====================
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# ===================== Load Background Image =====================
background_path = os.path.join("data", "background.png")
imgBackground = cv2.imread(background_path)

if imgBackground is None:
    print(f"Error: Could not open {background_path}. Please check if the file exists and is not corrupted.")
    exit(1)

# ===================== Resize Background to Fit Screen =====================
screen_res = 1280, 720
imgBackground = cv2.resize(imgBackground, screen_res)

# ===================== Frame and Embed Settings =====================
frame_width = 300
frame_height = 300

# Red box dimensions adjusted for resized image
box_x, box_y, box_w, box_h = 62, 140, 630, 430
embed_x = box_x + (box_w - frame_width) // 2
embed_y = box_y + (box_h - frame_height) // 2

# ===================== Set Webcam Resolution =====================
video = cv2.VideoCapture(0)
video.set(3, frame_width)
video.set(4, frame_height)

COL_NAMES = ['NAME', 'TIME']
print("Press 'o' to take attendance. Press 'q' to quit.")

# ===================== Main Loop =====================
while True:
    ret, frame = video.read()
    frame = cv2.resize(frame, (frame_width, frame_height))

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    attendance = None

    for (x, y, w, h) in faces:
        crop_img = frame[y:y + h, x:x + w]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)

        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        csv_path = f"Attendance/Attendance_{date}.csv"
        file_exists = os.path.isfile(csv_path)

        # Draw name on the webcam frame
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x + 5, y - 10),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)

        attendance = [str(output[0]), timestamp]

    # ===================== Embed Single Frame =====================
    imgDisplay = imgBackground.copy()
    imgDisplay[embed_y:embed_y + frame_height, embed_x:embed_x + frame_width] = frame

    cv2.imshow("Face Recognition Attendance", imgDisplay)

    key = cv2.waitKey(1)

    # ===================== Take Attendance =====================
    if key == ord('o') and attendance:
        os.makedirs("Attendance", exist_ok=True)
        already_taken = False

        if file_exists:
            with open(csv_path, "r", newline='') as csvfile:
                reader = csv.reader(csvfile)
                next(reader, None)
                for row in reader:
                    if row and row[0] == attendance[0]:
                        already_taken = True
                        break

        if already_taken:
            speak("Attendance already completed.")
        else:
            with open(csv_path, "a", newline='') as csvfile:
                writer = csv.writer(csvfile)
                if not file_exists:
                    writer.writerow(COL_NAMES)
                writer.writerow(attendance)
            speak("Attendance taken.")
        time.sleep(2)

    if key == ord('q'):
        break

video.release()
cv2.destroyAllWindows()
