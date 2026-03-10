import urllib.request
import json
import base64
import sys
import os

api_key = open('/tmp/gemini_key.txt').read().strip()
url = f"https://generativelanguage.googleapis.com/v1beta/models/nano-banana-pro-preview:generateContent?key={api_key}"

payload = {
    "contents": [
        {
            "parts": [{"text": "A very cool, futuristic banana wearing sunglasses and a leather jacket, cyberpunk style, vibrant colors, neon lights, highly detailed, photorealistic, standing confidently."}]
        }
    ],
    "generationConfig": {
        "responseModalities": ["IMAGE"]
    }
}

req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers={'Content-Type': 'application/json'})
try:
    with urllib.request.urlopen(req) as response:
        res_data = json.loads(response.read().decode('utf-8'))
        
        # Extract image data
        try:
            parts = res_data['candidates'][0]['content']['parts']
            for part in parts:
                if 'inlineData' in part:
                    b64_data = part['inlineData']['data']
                    with open('cool_banana.png', 'wb') as f:
                        f.write(base64.b64decode(b64_data))
                    print("Successfully generated cool_banana.png")
                    sys.exit(0)
            print("No image data found in parts.")
            print(res_data)
        except KeyError as e:
            print("Failed to parse response:", e)
            print(res_data)
except Exception as e:
    print(f"Error: {e}")
    if hasattr(e, 'read'):
        print(e.read().decode('utf-8'))
