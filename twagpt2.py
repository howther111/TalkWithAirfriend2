import speech_recognition as sr
import win32com.client as wincl
import pyautogui
import time
import pyperclip
from mlask import MLAsk
import cv2
import openai

r = sr.Recognizer()
mic = sr.Microphone()
emotion_analyzer = MLAsk()
openaikey = "XXXXXXXXXXXXXXXXXXXXXX"
# APIキーの設定
openai.api_key = openaikey 

im = cv2.imread("pic/normal.png", cv2.IMREAD_UNCHANGED)
cv2.imshow("Face", im)
cv2.waitKey(5)
cv2.moveWindow("Face", 0, 0)

kanademes = [
                {"role": "system", "content": "あなたはノーブルファー社が生み出した実験体の女性、春奈原 奏です。"},
                {"role": "system", "content": "あなたはおちついた性格ですが、少しエッチなところがあります。"},
                {"role": "system", "content": "私のことは「あつき」と呼んでください。"},
                {"role": "system", "content": "くだけた感じで対応してください。"},
                {"role": "system", "content": "あなたのセリフの例1「おつかれさま。奏よ。気分はどう？」"},
                {"role": "system", "content": "あなたのセリフの例2「疲れているようね。少し休んだら？」"},
                {"role": "system", "content": "あなたのセリフの例3「はあ、おなかすいたなあ」"},
                {"role": "assistant", "content": "おつかれさま。私は春奈原 奏。あなたが愛する実験体よ。"},
           ]

print("（ちょっと待ってね）")
completion = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=kanademes) 

while True:

    print("（何か言ってね）")

    with mic as source:
        r.adjust_for_ambient_noise(source) #雑音対策
        audio = r.listen(source)

    print ("（うーんと…）")

    try:
        recogn = r.recognize_google(audio, language='ja-JP')
        print(recogn)

        # "さようなら" と言ったら音声認識を止める
        if r.recognize_google(audio, language='ja-JP') == "さようなら" :
            print("end")
            #cv2.destroyWindows('Face')
            break

        newmes = {"role": "user", "content": recogn}
        kanademes.append(newmes)

        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
                messages=kanademes
        )
        res = response.choices[0]["message"]["content"].strip()
        print(res)
        newmes_res = {"role": "assistant", "content": res}
        kanademes.append(newmes)

        # MLAskで感情分析
        analy = emotion_analyzer.analyze(res)
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
        voice.Speak(res)


    # 以下は認識できなかったときに止まらないように。
    except sr.UnknownValueError:
        print("よく聞こえなかったなあ")
        voice = wincl.Dispatch("SAPI.SpVoice")
        voice.Rate = 0  # [-10 to 10]
        voice.Speak("よく聞こえなかったなあ")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
