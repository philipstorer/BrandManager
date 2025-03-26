import openai
openai.api_key = "YOUR_ACTUAL_API_KEY"
response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[{"role": "user", "content": "Hello"}],
)
print(response)
