import json
with open('d:/parsing_questions/parsed_questions.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print(f"Total processed: {len(data['processed'])}")
print(f"Total failures: {len(data['parsing_failures'])}")

for i, fail in enumerate(data['parsing_failures']):
    print(f"\n--- Failure {i+1} ---")
    print(f"Reason: {fail['reason']}")
    print(f"Q Number: {fail.get('q_number', 'N/A')}")
    print(f"Text Snippet: {fail['text'][:200]}...")
