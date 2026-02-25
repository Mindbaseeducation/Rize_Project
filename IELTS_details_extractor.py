import streamlit as st
import openai
import fitz
import base64
import tempfile
import pandas as pd

# -----------------------
# Config
# -----------------------
openai.api_key = st.secrets["openai"]["api_key"]

st.set_page_config(page_title="IELTS Extractor → Excel", layout="centered")
st.title("📄 IELTS Details Extractor")

# -----------------------
# Vision Extraction Function
# -----------------------
def extract_details_from_image(file_bytes):

    # Convert first page of PDF to image
    pdf = fitz.open(stream=file_bytes, filetype="pdf")
    page = pdf[0]

    pix = page.get_pixmap(dpi=300)
    img_bytes = pix.tobytes("png")
    base64_image = base64.b64encode(img_bytes).decode("utf-8")

    prompt = """
You are reading an official IELTS Test Report Form image.

CRITICAL RULES:

1) Extract the TEST DATE shown near "Centre Number".
2) DO NOT extract Date of Birth.
3) DO NOT extract the bottom issue date.
4) Extract each score strictly from its labeled box:
   - Listening → number next to Listening
   - Reading → number next to Reading
   - Writing → number next to Writing
   - Speaking → number next to Speaking
   - Overall Band Score → number next to it
   - CEFR Level → value next to it
5) Do NOT rearrange scores.
6) Do NOT guess.
7) If something is missing, leave blank.
8) Return EXACTLY 9 comma-separated values.
9) No labels. No explanations.

Return in this order:

First Name,
Family Name,
Date of Examination,
Listening Score,
Reading Score,
Writing Score,
Speaking Score,
Overall Band Score,
CEFR Level
"""

    response = openai.ChatCompletion.create(
        model="gpt-4o",
        temperature=0,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{base64_image}"
                        }
                    }
                ],
            }
        ],
    )

    output = response.choices[0].message.content.strip()
    st.write("🧠 GPT Output:", output)

    fields = [x.strip() for x in output.split(",")]

    # Ensure exactly 9 columns
    if len(fields) < 9:
        fields += [""] * (9 - len(fields))
    elif len(fields) > 9:
        fields = fields[:9]

    return fields


# -----------------------
# Streamlit UI
# -----------------------
uploaded_files = st.file_uploader(
    "Upload IELTS Report PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if st.button("Extract & Download Excel"):
    if not uploaded_files:
        st.error("Please upload at least one report.")
        st.stop()

    rows = []

    for file in uploaded_files:
        st.write(f"📄 Processing: {file.name}")
        file_bytes = file.read()
        extracted = extract_details_from_image(file_bytes)
        rows.append(extracted)

    df = pd.DataFrame(rows, columns=[
        "First Name",
        "Family Name",
        "Date of Examination",
        "Listening Score",
        "Reading Score",
        "Writing Score",
        "Speaking Score",
        "Overall Band Score",
        "CEFR Level"
    ])

    st.dataframe(df)

    # Export to Excel
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        df.to_excel(tmp.name, index=False)

        st.download_button(
            "📥 Download Excel",
            open(tmp.name, "rb").read(),
            file_name="IELTS Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        