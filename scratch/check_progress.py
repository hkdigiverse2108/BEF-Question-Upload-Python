import json
try:
    with open('d:/parsing_questions/parsed_questions.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    print(f"Total processed: {len(data['processed'])}")
    print(f"Total failures: {len(data['parsing_failures'])}")
    if data['processed']:
        print(f"Last question number: {data['processed'][-1]['number']}")
except Exception as e:
    print(f"Error: {e}")
