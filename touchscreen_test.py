import cv2
import mediapipe as mp
import pyautogui
import numpy as np
import math

print("Touchless Touchscreen Test Mode starting...")

# Screen size
screen_w, screen_h = pyautogui.size()

# Webcam
cap = cv2.VideoCapture(0)

# Mediapipe hand tracking
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(
    max_num_hands=1,
    min_detection_confidence=0.7,
    min_tracking_confidence=0.7
)
mp_draw = mp.solutions.drawing_utils

click_delay = 0

while True:
    success, frame = cap.read()
    if not success:
        continue

    frame = cv2.flip(frame, 1)
    h, w, _ = frame.shape
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb)

    if results.multi_hand_landmarks:
        hand = results.multi_hand_landmarks[0]
        lm = hand.landmark

        # Index finger tip
        ix, iy = int(lm[8].x * w), int(lm[8].y * h)
        # Thumb tip
        tx, ty = int(lm[4].x * w), int(lm[4].y * h)
        # Middle finger tip
        mx, my = int(lm[12].x * w), int(lm[12].y * h)

        # Map to screen coordinates
        screen_x = np.interp(ix, [0, w], [0, screen_w])
        screen_y = np.interp(iy, [0, h], [0, screen_h])
        pyautogui.moveTo(screen_x, screen_y, duration=0.01)

        # Pinch detection → click
        pinch = math.hypot(ix - tx, iy - ty)
        if pinch < 30 and click_delay == 0:
            pyautogui.click()
            click_delay = 15

        # Middle finger scroll
        if abs(my - iy) < 40:
            pyautogui.scroll(40 if my < iy else -40)

        # Draw landmarks
        mp_draw.draw_landmarks(frame, hand, mp_hands.HAND_CONNECTIONS)

    if click_delay > 0:
        click_delay -= 1

    cv2.imshow("Touchscreen Mode", frame)
    if cv2.waitKey(1) & 0xFF == 27:  # Press ESC to exit
        break

cap.release()
cv2.destroyAllWindows()
print("Touchscreen Test Mode exited.")
