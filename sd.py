# import the opencv library  https://pastebin.com/ixWD0BPC
import win32com.client      #needed to load COM objects
import cv2
from statistics import mean
import time
import numpy as np
# define a video capture object
vid = cv2.VideoCapture(0, cv2.CAP_DSHOW)
currentPov = 0
currentVis = 0






contcont = 0
NormalizationList = []
tel = win32com.client.Dispatch("ASCOM.SynScanMobile.Telescope")

if tel.Connected:
    print("Telescope was already connected")
else:
    tel.Connected = True
    if tel.Connected:
        print("	Connected to telescope now")
    else:
        print("	Unable to connect to telescope, expect exception")

print(tel.Description)

# tel.SlewToAltAz(currentPov, currentVis)
#
#
# time.sleep(70)
print("0 0 0 ok")
from skyfield.api import N, E, wgs84, load
planets = load('de421.bsp')
ts = load.timescale()
t = ts.now()


earth, mars = planets['earth'], planets['sun']
boston = earth + wgs84.latlon(	55.96 * N, 37.50636 * E,  elevation_m=180)
astrometric = boston.at(t).observe(mars)
alt, az, d = astrometric.apparent().altaz()

currentPov = az.degrees
currentVis = alt.degrees


tel.SlewToAltAz(currentPov, currentVis)

i = 63500
j = 0
sr = 0
r = 50
g = 40
b = 70
while (True):
    i = i + 1

    # try:

    # Capture the video frame
    # by frame
    ret, frame = vid.read()
    frame2 = frame
    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    th, im_th = cv2.threshold(frame, 60, 255, cv2.THRESH_BINARY)
    th2, im_th2_flag = cv2.threshold(frame, 200, 255, cv2.THRESH_BINARY)
    # print(th)

    # Load image, convert to grayscale, and Otsu's threshold
    image = frame2
    # gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh =im_th
    width = thresh.shape[1]
    height = thresh.shape[0]
    centreImX = width / 2
    centreImY = height / 2
    cx = centreImX
    cy = centreImY
    # Find contours and extract the bounding rectangle coordintes
    # then find moments to obtain the centroid
    cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    if len(cnts) > 1:
        cnts_1 = max(cnts, key=cv2.contourArea)
        cnts = [1]
        cnts[0] = cnts_1

    cnts2 = cv2.findContours(im_th2_flag, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
    cnts2 = cnts2[0] if len(cnts2) == 2 else cnts2[1]

    if len(cnts2) > 1:
        cnts_12 = max(cnts2, key=cv2.contourArea)
        cnts2 = [1]
        cnts2[0] = cnts_12


    # print(len(cnts))
    # if len(cnts) != 0:
    #     cnts = max(cnts, key=cv2.contourArea)
    #
    # contcont = 0
    for c in cnts:
            # Obtain bounding box coordinates and draw rectangle
            x, y, w, h = cv2.boundingRect(c)
            cv2.rectangle(image, (x, y), (x + w, y + h), (b, g, r), 2)

            # Find center coordinate and draw center point
            M = cv2.moments(c)
            if M['m00'] ==0:
                M['m00'] = 1
            cx = int(M['m10'] / M['m00'])
            cy = int(M['m01'] / M['m00'])
            cv2.circle(image, (cx, cy), 2, (36, 255, 12), -1)
            cv2.circle(image, (int(centreImX), int(centreImY)), 5, (255, 100, 50), 2)



    if len(cnts2) > 0:
        cnts_12 = cnts2[0]
        contcont2 = cv2.contourArea(cnts_12)
        if (i % 30) == 0:
            NormalizationList.append(contcont2)
            sr = mean(NormalizationList)
            if (sr - contcont2) > sr/4:
                r = 255
                b = 20
                g = 40
            else:
                r = 20
                b = 40
                g = 255


    width = thresh.shape[1]
    height = thresh.shape[0]





    # if (i % 60) == 0:
    #     if abs(cx - centreImX) > 80:
    #         if cx > centreImX:
    #             currentPov = currentPov - 1
    #         if cx < centreImX:
    #             currentPov = currentPov + 1
    #     if abs(cy - centreImX) > 80:
    #         if cy > centreImY:
    #             currentVis = currentVis - 1
    #         if cy < centreImY:
    #             currentVis = currentVis + 1
    if (i % 150) == 0 and currentVis < 87 and currentVis > - 6 :
        if cx - centreImX > 100:
            currentPov = currentPov - 1
        if cx - centreImX < (-100):
            currentPov = currentPov + 1
        if cy - centreImY > 100:
            currentVis = currentVis - 1
        if cy - centreImY < (-100):
            currentVis = currentVis + 1


    if (i % 38) == 0:
        if cx - centreImX > 4:
            currentPov = currentPov - 0.03
        if cx - centreImX < (-4):
            currentPov = currentPov + 0.03
        if cy - centreImY > 4:
            currentVis = currentVis - 0.03
        if cy - centreImY < (-4):
            currentVis = currentVis + 0.03
        if cx + cy > 4:
            currentPov1 = currentPov
            if currentPov < 0:
                currentPov1 = 360 + currentPov
            tel.SlewToAltAzAsync(currentPov1, currentVis)
    if (i % 100) == 0:
        cv2.imwrite('C:/Users/MIPT-1/Pictures/Sosolar/' +str(i) + "2" + 'Test_gray.jpg', image)
    if abs(cx - centreImX) < 4:
        tel.AbortSlew
    if abs(cy - centreImY) < 4:
        tel.AbortSlew
    # print("Raznica:")
    # print(cx - centreImX)
    # print(cy - centreImY)
    s = ''
    s += "centre:"
    s += f"{cx,cy}"
    s += "   "
    s += "diff:"
    s += f"{cx - centreImX, cy - centreImY}"
    s += "   "
    s += "Area:"
    s += f"{contcont, sr}"

    print(s, end="\r")





    cv2.imshow('image', image)
    cv2.waitKey(1)
















    # 128.0

    # Display the resulting frame
    # cv2.imshow('frame',im_th)

    # the 'q' button is set as the
    # quitting button you may use any
    # desired button of your choice
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
    # except:
    #     print(Exception)

# After the loop release the cap object
vid.release()
# Destroy all the windows
cv2.destroyAllWindows()
