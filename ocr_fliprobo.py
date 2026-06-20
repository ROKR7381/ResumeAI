import os
import base64
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv(override=True)

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def main():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("OPENAI_API_KEY is not set.")
        return
        
    client = OpenAI(api_key=api_key)
    base64_image = encode_image("p1.png")
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text", 
                            "text": "Look at the section 'Employment History' -> 'DATA SCIENTIST at FLIPROBO TECHNOLOGY' in the preview on the right side. Represent the complete visible text under this section (the project title '1.Insurance Claims - Fraud Detection' and all bullet points/lines under it) as a structured JSON array of strings. Do not omit any word that is visible."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=1000
        )
        print(response.choices[0].message.content)
    except Exception as e:
        print("Error:", e)

if __name__ == '__main__':
    main()
