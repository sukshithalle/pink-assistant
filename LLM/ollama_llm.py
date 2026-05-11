import ollama

def ask_llm(prompt):
    response = ollama.chat(
        model="llama3:8b",
        messages=[
            {"role": "system", "content": "You are Pink Assistant, a helpful AI."},
            {"role": "user", "content": prompt}
        ]
    )
    return response["message"]["content"]