import requests


# Endpoint URL
url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"

# Headers
headers = {
    "Content-Type": "application/json"
}

# Prompt the user for a question
user_question = input("Enter your question for Gemini AI: ")

# Request body
data = {
    "contents": [
        {
            "parts": [
                {
                    "text": user_question
                }
            ]
        }
    ]
}

# Make the POST request
response = requests.post(url, headers=headers, json=data)

response_json = response.json()
try:
    generated_text = response_json['candidates'][0]['content']['parts'][0]['text']
    print("\n" + "="*40)
    print("Gemini AI's Response:")
    print("-"*40)
    print(generated_text)
    print("="*40 + "\n")
except (KeyError, IndexError):
    print("\n[Error] Could not parse the response from Gemini AI.")
    print("Raw response:", response_json)
