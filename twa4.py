import speech_recognition as sr
import win32com.client as wincl
import pyautogui
import time
import pyperclip
from mlask import MLAsk
import cv2

r = sr.Recognizer()
mic = sr.Microphone()
emotion_analyzer = MLAsk()

im = cv2.imread("pic/normal.png", cv2.IMREAD_UNCHANGED)
cv2.imshow("Face", im)
cv2.waitKey(5)
cv2.moveWindow("Face", 0, 0)

while True:

    print("Say something ...")

    with mic as source:
        r.adjust_for_ambient_noise(source) #雑音対策
        audio = r.listen(source)

    print ("Now to recognize it...")

    try:
        recogn = r.recognize_google(audio, language='ja-JP')
        print(recogn)

        # "さようなら" と言ったら音声認識を止める
        if r.recognize_google(audio, language='ja-JP') == "さようなら" :
            print("end")
            #cv2.destroyWindows('Face')
            break

        pyperclip.copy(recogn)

        # チャット欄
        pyautogui.click(x=990, y=795, interval=0.5, button="left")
        pyautogui.hotkey("ctrl", "v")
        pyautogui.hotkey("enter")
        time.sleep(5)

        # エアフレンドの発言
        pyautogui.click(x=1081, y=697, interval=0.5, button="right")

        # コピー
        pyautogui.click(x=1135, y=724, interval=0.5, button="left")
        time.sleep(1)
        clip_str = pyperclip.paste()
        print(clip_str)

        # MLAskで感情分析
        analy = emotion_analyzer.analyze(clip_str)
        rep = "normal"
        try:
            rep = analy["representative"][0]
        except:
            rep = "normal"

        print(analy)
        print(rep)
        emo = "pic/" + rep + ".png"
        im = cv2.imread(emo, cv2.IMREAD_UNCHANGED)
        cv2.waitKey(5)
        cv2.imshow("Face", im)

        voice = wincl.Dispatch("SAPI.SpVoice")
        voice.Rate = 0  # [-10 to 10]
        voice.Speak(clip_str)


    # 以下は認識できなかったときに止まらないように。
    except sr.UnknownValueError:
        print("could not understand audio")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
