import extractor



data = extractor.extract_from_docx("sample_exam.docx")

import json
print(json.dumps(data, indent=2, ensure_ascii=False))

