from openai import OpenAI
from openai.types.chat.chat_completion import ChatCompletion
from openai.types.chat.chat_completion_message_param import ChatCompletionMessageParam
import time

client: OpenAI = None  # Global OpenAI client instance

def configure_openai(api_key: str):
    global client
    client = OpenAI(api_key=api_key)

def classify_response(email_subject: str, email_preview: str) -> str:
    prompt = f"""
You are an assistant helping someone track job applications.

Given the following email:
Subject: "{email_subject}"
Preview: "{email_preview}"

What type of response is this? Choose only one of:
- Applied
- Rejected
- Interview
- Offer
- No Reply Yet
- Other

Respond with only the label, nothing else.
"""

    try:
        response: ChatCompletion = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[ChatCompletionMessageParam(role="user", content=prompt)],
            max_tokens=5,
            temperature=0,
        )
        classification = response.choices[0].message.content.strip()
        return classification
    except Exception as e:
        print("OpenAI API error:", e)
        time.sleep(2)
        return "Error"
