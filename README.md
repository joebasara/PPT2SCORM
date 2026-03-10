# PPTX to SCORM Converter

This Streamlit app converts PowerPoint (.pptx) files into SCORM 1.2 packages to upload in LMS platforms such as Blackboard.

## Features

- Slide navigation
- Hyperlinks preserved
- Internal slide links
- Images preserved
- SCORM 1.2 package output

## Run locally

Install dependencies:

pip install -r requirements.txt

Run the app:

streamlit run app.py

## Deploy on Streamlit Cloud

1. Upload this project to GitHub
2. Connect the repository in Streamlit Cloud
3. Deploy `app.py`

## Limitations

This MVP does not fully reproduce PowerPoint rendering. The following are not supported:

- animations
- transitions
- SmartArt

- advanced triggers
