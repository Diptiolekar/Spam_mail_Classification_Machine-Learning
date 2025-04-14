import pickle
import streamlit as st
import win32com.client
import pythoncom 
from PIL import Image
import base64
from io import BytesIO

def speak(text):
	pythoncom.CoInitialize()
	speak=win32com.client.Dispatch(("SAPI.SpVoice"))
	speak.Speak(text)

def image_to_base64(image_path):
    img = Image.open(image_path)
    buffered = BytesIO()
    img.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    return img_str

def load_css(img_str):
    css = f"""
    <style>
    .stApp {{
        background-image: url('data:image/jpg;base64,{img_str}');
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
    }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

img_str = image_to_base64("email1.jpg")
load_css(img_str)


model=pickle.load(open("spam.pkl","rb"))
cv=pickle.load(open("vectorizer.pkl","rb"))



def main():
	st.title("Email Spam Classification App")
	st.subheader("Build with IT VEDANT")
	msg=st.text_input("Enter a Text: ")
	if st.button("Predict"):
		data=[msg]
		vect=cv.transform(data).toarray()
		prediction=model.predict(vect)
		result=prediction[0]
		if result==1:
			st.error("This is a SPAM EMAIL !!!")
			speak("This is a SPAM EMAIL")
		else:
			st.success("This is a HAM Email")
			speak("This is a HAM EMAIL")


if __name__ == "__main__":
    main()