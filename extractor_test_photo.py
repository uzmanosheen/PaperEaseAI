import extractor


with open("sample_exam.jpg", "rb") as f:
    data = extractor.extract_from_images([(f.read(), "image/jpeg")])

import json
print(json.dumps(data, indent=2, ensure_ascii=False))