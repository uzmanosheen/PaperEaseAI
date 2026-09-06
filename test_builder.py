import json
import doc_builder

with open("sample_extracted.json", "r", encoding="utf-8") as f:
    data = json.load(f)   # save the JSON you got from Day 2 into this file first

header_info = {
    "school_name": "City Public School",
    "subject": "Mathematics",
    "class_name": "9th",
    "term": "Mid Term",
    "session": "2025-26",
    "total_marks": "75",
    "time": "2 Hours",
}

doc_builder.build_docx(data, header_info, logo_path=None, template="template1", output_path="test_output.docx")
print("Saved test_output.docx")