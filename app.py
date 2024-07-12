import cv2
import numpy as np
import tkinter as tk
from PIL import Image, ImageTk
import mediapipe as mp
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Initialize pycaw for volume control
def initialize_pycaw():
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        return volume
    except Exception as e:
        print(f"Failed to initialize pycaw: {e}")
        return None

# Initialize MediaPipe Hands
mp_hands = mp.solutions.hands
hands = mp_hands.Hands()
mp_draw = mp.solutions.drawing_utils

# Initialize tkinter window
root = tk.Tk()
root.title("Hand Gesture Volume Control")
root.geometry("800x600")

# Function to update webcam frame and perform hand gesture recognition
def update_webcam_frame():
    global prev_loc_x, prev_loc_y

    # Capture frame-by-frame from the camera
    success, img = cap.read()

    if success:
        # Convert the image to RGB
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        results = hands.process(img_rgb)

        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                mp_draw.draw_landmarks(img, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                landmarks = hand_landmarks.landmark

                # Get the coordinates of the thumb tip and index finger tip
                thumb_tip = landmarks[mp_hands.HandLandmark.THUMB_TIP]
                index_tip = landmarks[mp_hands.HandLandmark.INDEX_FINGER_TIP]

                # Convert normalized coordinates to pixel values
                h, w, c = img.shape
                thumb_x, thumb_y = int(thumb_tip.x * w), int(thumb_tip.y * h)
                index_x, index_y = int(index_tip.x * w), int(index_tip.y * h)

                # Draw circles on thumb tip and index finger tip
                cv2.circle(img, (thumb_x, thumb_y), 15, (255, 0, 255), cv2.FILLED)
                cv2.circle(img, (index_x, index_y), 15, (255, 0, 255), cv2.FILLED)
                cv2.line(img, (thumb_x, thumb_y), (index_x, index_y), (255, 0, 255), 3)

                # Calculate the distance between thumb tip and index finger tip
                length = np.hypot(index_x - thumb_x, index_y - thumb_y)

                # Map the distance to the volume range
                vol = np.interp(length, [30, 300], [-65.25, 0.0])  # Adjust according to your volume range

                try:
                    volume.SetMasterVolumeLevel(vol, None)
                except Exception as e:
                    print(f"Failed to set volume: {e}")

                # Display volume level as a bar
                vol_bar = np.interp(length, [30, 300], [400, 150])
                cv2.rectangle(img, (50, 150), (85, 400), (0, 255, 0), 3)
                cv2.rectangle(img, (50, int(vol_bar)), (85, 400), (0, 255, 0), cv2.FILLED)
                cv2.putText(img, f'Vol: {int(length)}%', (40, 450), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)

        # Convert OpenCV image to PIL format
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        img = cv2.resize(img, (640, 480))
        img_pil = Image.fromarray(img)
        img_tk = ImageTk.PhotoImage(image=img_pil)

        # Update the label with the new image
        label_webcam.img_tk = img_tk
        label_webcam.config(image=img_tk)

    # Call this function again after 10ms
    label_webcam.after(10, update_webcam_frame)

# Initialize the webcam
cap = cv2.VideoCapture(0)
cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

# Create a frame for webcam feed
frame_webcam = tk.Frame(root, width=640, height=480)
frame_webcam.pack(pady=20)

# Label to display webcam feed
label_webcam = tk.Label(frame_webcam)
label_webcam.pack()

# Start updating webcam frame
update_webcam_frame()

# Exit button
btn_exit = tk.Button(root, text="Exit", command=root.quit)
btn_exit.pack()

# Start the tkinter main loop
root.mainloop()

# When everything is done, release the capture
cap.release()
cv2.destroyAllWindows()
