import win32api
import win32con
import keyboard as kb
import time
import cv2
import numpy as np 
import imutils
import glob
import pyautogui


panorama = int(360 // 85)
panorama_ver = int(180 // 85)
total = panorama * panorama_ver
print(total)
save = "C:/Users/blap/Videos/Captures/"
image_paths = glob.glob('C:/Users/blap/Videos/Captures/*.png')
images = []
image_nb = 0

def resize_images(images):
    #return cv2.resize(images, (int(images.shape[1] / 4), int(images.shape[0] / 4)))
    return cv2.resize(images, (854*2, 480*2))


def execute_01(save_path):
    for image in image_paths:
        img = cv2.imread(image)
        images.append(resize_images(img))

    imageStitcher = cv2.Stitcher_create()
    #imageStitcher.setRegistrationResol(0.1)
    #imageStitcher.setSeamEstimationResol(0.1)
    #imageStitcher.setPanoConfidenceThresh(0.1)

    error, stitched_img = imageStitcher.stitch(images)

    if not error:
        try:
            cv2.imwrite(save_path+"stitchedOutput.png", stitched_img)
        except Exception as e:
            print(f"Error: {e}")
        stitched_img = cv2.copyMakeBorder(stitched_img, 10, 10, 10, 10, cv2.BORDER_CONSTANT, (0,0,0))

        gray = cv2.cvtColor(stitched_img, cv2.COLOR_BGR2GRAY)
        thresh_img = cv2.threshold(gray, 0, 255 , cv2.THRESH_BINARY)[1]

        contours = cv2.findContours(thresh_img.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        contours = imutils.grab_contours(contours)
        areaOI = max(contours, key=cv2.contourArea)

        mask = np.zeros(thresh_img.shape, dtype="uint8")
        x, y, w, h = cv2.boundingRect(areaOI)
        cv2.rectangle(mask, (x,y), (x + w, y + h), 255, -1)

        minRectangle = mask.copy()
        sub = mask.copy()

        while cv2.countNonZero(sub) > 0:
            minRectangle = cv2.erode(minRectangle, None)
            sub = cv2.subtract(minRectangle, thresh_img)


        contours = cv2.findContours(minRectangle.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        contours = imutils.grab_contours(contours)
        areaOI = max(contours, key=cv2.contourArea)

        x, y, w, h = cv2.boundingRect(areaOI)

        stitched_img = stitched_img[y:y + h, x:x + w]

        try:
            cv2.imwrite(save_path+"stitchedOutputProcessed.png", stitched_img)
        except Exception as e:
            print(f"Error: {e}")

        print("Done!")

    else:
        print(f"Images could not be stitched! {error}")
        print("Likely not enough keypoints being detected!")


def press(nb_image):
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(rf'{save}\Image{nb_image}.png')
    #win32api.keybd_event(0x5B, 0,win32con.KEYEVENTF_EXTENDEDKEY | 0 ,0)
    #win32api.keybd_event(0x12, 0,win32con.KEYEVENTF_EXTENDEDKEY | 0 ,0)
    #win32api.keybd_event(0x2C, 0,win32con.KEYEVENTF_EXTENDEDKEY | 0 ,0)
    #time.sleep(.1)
    #win32api.keybd_event(0x5B, 0,win32con.KEYEVENTF_KEYUP ,0)
    #win32api.keybd_event(0x12, 0,win32con.KEYEVENTF_KEYUP ,0)
    #win32api.keybd_event(0x2C, 0,win32con.KEYEVENTF_KEYUP ,0)

def turn(nb_image):
    press(nb_image)
    time.sleep(.1)
    centre_x = win32api.GetSystemMetrics(win32con.SM_CXSCREEN) // 2
    centre_y = win32api.GetSystemMetrics(win32con.SM_CYSCREEN) // 2
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, centre_x - int(360 // total), 0,0,0)


def turn_down():
    time.sleep(.1)
    centre_y = win32api.GetSystemMetrics(win32con.SM_CYSCREEN) // 2
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, 0, centre_y - (180 // total),0,0)


#execute_01(save)
while True:
    if kb.is_pressed('j'):
        i = 0
        print("Screen Start")
        for _ in range(5):
            for _ in range(total):
                turn(i)
                time.sleep(.8)
            turn_down()

        print("Screen End")
        print("-"*16)
        print("Stitcher Start")
        start_time = time.time()
        execute_01(save)
        print("Stitcher End")
        print(f'Execution Time : {time.time() - start_time}')
