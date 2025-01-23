from googletrans import Translator
from vosk import Model , KaldiRecognizer
import pyaudio
import time 
from playsound import playsound
import json
import edge_tts
from openai import OpenAI
# fa-IR-DilaraNeural
VOICE = "fa-IR-DilaraNeural"
OUTPUT_FILE = "Path-to-test.mp3"
messages = []
client = OpenAI(
    base_url="https://api-inference.huggingface.co/v1/",
    api_key="your_hugging-face_api_key"
)
model = Model("path-to-persian-vosk-model")
recognizer = KaldiRecognizer(model, 16000)
with open("tt.txt", 'w') as file:  
    pass 
translatorx = Translator() 
# Start audio stream
mic = pyaudio.PyAudio()
stream = mic.open(            format=pyaudio.paInt16,
            channels=1,
            rate=16000,
            input=True,
            frames_per_buffer=8000)
stream.start_stream()


last_two_partials = ["", ""]
is_recognizing = False
while True:
    print("Listening...")
    data = stream.read(4000, exception_on_overflow=False)  
    
    ########################
    if recognizer.AcceptWaveform(data):
        result = recognizer.Result()
    else :
        result = recognizer.PartialResult()
    try:
        result_dict = json.loads(result)  
    except json.JSONDecodeError:
        continue
    if "شروع" in result_dict.get("text", ""):
        if not is_recognizing: 
            is_recognizing = True
            print("شناسایی شروع شد...")
            playsound("C:/Users/Mobile Gandom/Desktop/project_files/ding.mp3")
    elif "پایان" in result_dict.get("text", ""):
        if is_recognizing:  
            is_recognizing = False
            print("شناسایی متوقف شد.")
            playsound("C:/Users/Mobile Gandom/Desktop/project_files/ding.mp3")
    ###################
    if "partial" in result_dict:

        partial_text = result_dict["partial"].strip()

        if partial_text and is_recognizing:
            last_two_partials[0], last_two_partials[1] = last_two_partials[
                            1], partial_text 
            print(f"در حال صحبت: {partial_text}")
        else:  # اگر partial خالی بود، آخرین partial غیرخالی را در فایل ذخیره می‌کنیم
            if last_two_partials[1]:  # اگر آخرین partial معتبر است
                with open("tt.txt", mode="a", encoding="utf-8") as f:
                    f.write(("User's question : "+(last_two_partials[1]).replace("پایان"," ")) + "\n")
                    text = (last_two_partials[1]).replace("پایان","")
                    answer = translatorx.translate(text)
                    translated_text = answer.text
                    username = "your username"
                    assistant_name = "your ai username"
                    prompt = f"{username} is a user who seeks knowledge and assistance. " \
                  f"You, as a helpful assistant, should refer to yourself as '{assistant_name}'. " \
                  f"Respond to the following question from {username}: {translated_text}"
                    messages.append({"role":"user","content":translated_text})
                    response = client.chat.completions.create(
                                    model="Qwen/Qwen2.5-72B-Instruct",
                                    messages=[{"role": "user", "content": prompt}],
                                    temperature=0.5,
                                    max_tokens=64,
                                    top_p=0.7,
                                )
                    msg = response.choices[0].message.content
                    messages.append({"role": "assistant", "content": msg})
                    with open("tt.txt", mode="a", encoding="utf-8") as f:
                        f.write((last_two_partials[1]).replace("پایان","") + "\n")
                        f.write('Translated question : '+(translated_text) + "\n")
                        f.write(('English answer : '+msg) + "\n")
                        a = translatorx.translate(msg,src="en",dest="fa")
                        f.write(("Persian answer : "+a.text) + "\n")
                    communicate = edge_tts.Communicate(a.text, VOICE)
                    with open(OUTPUT_FILE, "wb") as file:
                        for chunk in communicate.stream_sync():
                            if chunk["type"] == "audio":
                                file.write(chunk["data"])
                    playsound(OUTPUT_FILE)
                    is_recognizing = False

                last_two_partials = ["", ""]
                


