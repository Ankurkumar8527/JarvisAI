from openai import OpenAI
from config import api_key

client = OpenAI(api_key=api_key)

response = client.responses.create(
    model="gpt-4.1-mini",
    input="Write a professional resignation email to my boss."
)

print(response.output_text)
