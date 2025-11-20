#AI Assisted in generation of quizzes based on the data feed to it

import requests
import json
import csv

OLLAMA_URL = "http://192.168.1.17:11434/api/generate"
HEADER = ["Question","Answer"]

def ask_ollama(prompt, model="llama3"):
    data = {
        "model": model,
        "prompt": prompt,
        "stream": False
    }
    response = requests.post(OLLAMA_URL, json=data)
    return response.json()["response"]


def main():
    result = ask_ollama("Create a 3 multiple choice question with 4 choices and provide the correct answer.")

    # Extract fields from the response
    row = {
    "question": result.get("question", ""),
    "choice1": result.get("choice1", ""),
    "choice2": result.get("choice2", ""),
    "choice3": result.get("choice3", ""),
    "choice4": result.get("choice4", ""),
    "answer": result.get("answer", "")
    }

    # CSV filename
    csv_file = "questions.csv"

   # Write to CSV
    with open(csv_file, "a", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=row.keys())

    # Write header only if file is empty
    if file.tell() == 0:
        writer.writeheader()

    writer.writerow(row)

    print("Saved to questions.csv!")





if __name__ == "__main__":
    main()
