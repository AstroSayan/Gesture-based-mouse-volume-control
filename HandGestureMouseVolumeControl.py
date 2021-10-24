import cv2
import time
import numpy as np
import autopy
import cvzone.HandTrackingModule as htm
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

##########################
wCam, hCam = 640, 480
frameR = 100  # Frame Reduction
smoothening = 7
#########################

cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)
# Hand Detector object
detector = htm.HandDetector(detectionCon=0.7, maxHands=1)
# Info regarding screen size
wScr, hScr = autopy.screen.size()
# Initializing audio drivers and controller
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volRange = volume.GetVolumeRange()

# Initialization
minVol = volRange[0]
maxVol = volRange[1]
pTime = 0
plocX, plocY = 0, 0
clocX, clocY = 0, 0
# flag_vol = 0
vol = 0
volBar = 400
volPer = 0
area = 0
colorVol = (255, 0, 0)

while True:
    success, img = cap.read()
    hands, img = detector.findHands(img)
    if hands:
        # Hand 1
        hand1 = hands[0]
        lmList1 = hand1["lmList"]
        bbox1 = hand1["bbox"]
        h_t = hand1["type"]

        if len(lmList1) != 0:
            # Triggering volume control by using Left hand
            if h_t == "Left":
                area = (bbox1[2] - bbox1[0]) * (bbox1[3] - bbox1[1]) // 100
                if 150 < area < 1000:
                    # Find Distance between index and Thumb
                    length, lineInfo, img = detector.findDistance(lmList1[4], lmList1[8], img)
                    # Convert Volume
                    volBar = np.interp(length, [50, 200], [400, 150])
                    volPer = np.interp(length, [50, 200], [0, 100])
                    # Reduce Resolution to make it smoother
                    smoothness = 10
                    volPer = smoothness * round(volPer / smoothness)
                    # Check fingers up
                    fingers = detector.fingersUp(hand1)
                    # If middle is down set volume
                    if not fingers[4]:
                        volume.SetMasterVolumeLevelScalar(volPer / 100, None)
                        cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                        colorVol = (0, 255, 0)
                    else:
                        colorVol = (255, 0, 0)

                # Drawings
                cv2.rectangle(img, (50, 150), (85, 400), (255, 0, 0), 3)
                cv2.rectangle(img, (50, int(volBar)), (85, 400), (255, 0, 0), cv2.FILLED)
                cv2.putText(img, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX,
                            1, (255, 0, 0), 3)
                cVol = int(volume.GetMasterVolumeLevelScalar() * 100)
                cv2.putText(img, f'Vol Set: {int(cVol)}', (400, 50), cv2.FONT_HERSHEY_COMPLEX,
                            1, colorVol, 3)
            elif h_t == "Right":
                x1, y1 = lmList1[8]
                x2, y2 = lmList1[12]

                fingers = detector.fingersUp(hand1)
                cv2.rectangle(img, (frameR, frameR), (wCam - frameR, hCam - frameR),
                              (255, 0, 255), 2)
                # 4. Only Index Finger : Moving Mode
                if fingers[1] == 1 and fingers[2] == 0:
                    # 5. Convert Coordinates
                    x = np.interp(x1, (frameR, wCam - frameR), (0, wScr))
                    y = np.interp(y1, (frameR, hCam - frameR), (0, hScr))
                    # 6. Smoothen Values
                    clocX = plocX + (x - plocX) / smoothening
                    clocY = plocY + (y - plocY) / smoothening

                    # 7. Move Mouse
                    autopy.mouse.move(wScr - clocX, clocY)
                    cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
                    plocX, plocY = clocX, clocY
                # 8. Both Index and middle fingers are up : Clicking Mode
                if fingers[1] == 1 and fingers[2] == 1:
                    # 9. Find distance between fingers
                    length, lineInfo, img = detector.findDistance(lmList1[8], lmList1[12], img)
                    # print(length)
                    # 10. Click mouse if distance short
                    if length < 30:
                        cv2.circle(img, (lineInfo[4], lineInfo[5]),
                                   15, (0, 255, 0), cv2.FILLED)
                        autopy.mouse.click()

    # Frame Rate
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS: {int(fps)}', (40, 50), cv2.FONT_HERSHEY_PLAIN,
                3, (255, 0, 0), 3)
    cv2.imshow("Img", img)
    cv2.waitKey(1)
