import os
import pandas as pd
from pyresparser import ResumeParser

data_list = []

resume_folder = "../resumes"

for file in os.listdir(resume_folder):

    if file.endswith(".pdf"):

        file_path = os.path.join(resume_folder, file)

        data = ResumeParser(file_path).get_extracted_data()

        if data:
            data_list.append({
                "Name": data.get("name"),
                "Email": data.get("email"),
                "Skills": data.get("skills")
            })

df = pd.DataFrame(data_list)

df.to_csv("../output/resume_data.csv", index=False)

print("Resume Parsing Completed")