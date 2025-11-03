import streamlit as st
import pandas as pd
import requests
from docx import Document
from io import BytesIO, StringIO
import io
import re
import json
import html as html_module
import streamlit.components.v1 as components

# ----------------------------
# APP CONFIG
# ----------------------------
st.set_page_config(page_title="AI Test Case Generator (Excel Output)", layout="wide")
st.title("🤖 AI-Powered Test Case Generator — Excel Output")

st.markdown("""
Upload a Word document containing your User Story + Acceptance Criteria.  
The tool will generate Functional, Validation/UI and Database test cases per AC and provide a single Excel download.
""")

# ----------------------------
# API KEY (replace for production)
# ----------------------------
OPENROUTER_API_KEY = "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"

# ----------------------------
# SESSION STATE INIT
# ----------------------------
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

# ----------------------------
# HELPERS
# ----------------------------
def extract_text_from_docx(file):
    doc = Document(file)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(paragraphs)

def extract_acceptance_criteria(full_text):
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
    ac_pattern = re.compile(r'^(AC\s*\d+|AC\d+)\b[:\-\–]?\s*(.*)', re.IGNORECASE)
    ac_indices = []
    for i, ln in enumerate(lines):
        if ac_pattern.match(ln):
            ac_indices.append(i)

    results = []
    if ac_indices:
        for idx, start in enumerate(ac_indices):
            end = ac_indices[idx+1] if idx+1 < len(ac_indices) else len(lines)
            block_lines = lines[start:end]
            label_match = ac_pattern.match(block_lines[0])
            label = label_match.group(1).replace(" ", "").upper()
            block_lines[0] = ac_pattern.sub(r'\2', block_lines[0]).strip()
            ac_text = " ".join([ln for ln in block_lines if ln])
            results.append((label, ac_text))
        return results

    heading_idx = None
    for i, ln in enumerate(lines):
        if re.search(r'acceptance criteria', ln, re.IGNORECASE):
            heading_idx = i
            break

    if heading_idx is not None:
        tail = lines[heading_idx+1:]
        current = []
        for ln in tail:
            if re.match(r'^(AC\s*\d+|AC\d+)\b', ln, re.IGNORECASE):
                if current:
                    results.append((f"AC{len(results)+1}", " ".join(current)))
                    current = []
                current.append(re.sub(r'^(AC\s*\d+[:\-\–]?\s*)', '', ln, flags=re.IGNORECASE))
            else:
                current.append(ln)
        if current:
            results.append((f"AC{len(results)+1}", " ".join(current)))
        results = [(lbl if lbl.upper().startswith("AC") else f"AC{i+1}", txt) for i,(lbl,txt) in enumerate(results)]
        return results

    paragraphs = re.split(r'\n{2,}', full_text)
    for i, p in enumerate(paragraphs):
        p = p.strip()
        if p:
            results.append((f"AC{i+1}", p))
    return results

def call_openrouter_api(prompt):
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "gpt-4o-mini",
        "messages": [
            {"role": "system", "content": "You are an expert QA engineer who returns only CSV rows (no headers, no explanations)."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.0
    }
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        resp.raise_for_status()
        d = resp.json()
        if "choices" in d and len(d["choices"])>0:
            content = d["choices"][0].get("message", {}).get("content") or d["choices"][0].get("text") or ""
            return content
        return ""
    except Exception as e:
        return f"⚠️ Error: {str(e)}"

def build_prompt_for_ac(ac_label, ac_text):
    ac_text_clean = ac_text.strip().replace('\n', ' ')
    prompt = f"""
Acceptance Criterion: {ac_text_clean}

For this single Acceptance Criterion, generate multiple test case rows in CSV format with columns:
Testing Type,Test Scenario,Test Steps,Expected Result

Rules:
- Include test cases of types: Functional, Validation/UI, Database (if applicable).
- For Testing Type use exactly one of: Functional, Validation/UI, Database.
- Keep Test Scenario short (single sentence).
- Keep Test Steps numbered and concise (e.g., "1. Do X 2. Do Y").
- Return ONLY CSV rows (no header row, no explanation, no markdown). Each CSV row must contain exactly 3 commas.
- Provide 2-4 logical test cases covering positive, negative, and boundary scenarios as appropriate.

Generate rows now.
"""
    return prompt

def parse_ai_csv_rows_to_df(ai_text):
    txt = (ai_text or "").strip()
    txt = re.sub(r'```.*?```', '', txt, flags=re.DOTALL)
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    cand = [ln for ln in lines if ln.count(",") >= 3]
    if not cand:
        cand = [ln for ln in lines if ln.count(",") >= 2]
    if not cand:
        return None
    header = "Testing Type,Test Scenario,Test Steps,Expected Result"
    csv_content = header + "\n" + "\n".join(cand)
    try:
        df = pd.read_csv(StringIO(csv_content))
        for c in df.select_dtypes(include=['object']).columns:
            df[c] = df[c].astype(str).str.strip()
        return df
    except Exception:
        return None

def assign_tc_ids(df, ac_label):
    mapping = {
        'Functional': 'FUNC',
        'Database': 'DB',
        'Validation/UI': 'UI',
        'Validation': 'UI',
        'UI': 'UI'
    }
    counters = {}
    ids = []
    for _, row in df.iterrows():
        t = str(row.get('Testing Type', '')).strip()
        key = mapping.get(t, 'OTH')
        counters.setdefault(key, 0)
        counters[key] += 1
        ids.append(f"{ac_label}-{key}-{counters[key]:03d}")
    df.insert(0, "Test Case ID", ids)
    return df

# ----------------------------
# UI - Inputs
# ----------------------------
uploaded_file = st.file_uploader("📄 Upload a Word Document (.docx)", type=["docx"])
text_area_story = st.text_area("📝 Or paste your user story here (optional):", value="")

col1, col2 = st.columns([3,1])
with col1:
    st.write("Generating: Functional + Validation/UI + Database testcases per AC")
with col2:
    st.write("Defaults:")
    st.write("- Role: **Super User**")
    st.write("- Excel-only output (no raw AI text on screen)")

max_acs = st.number_input("Max ACs to process (0 = all)", min_value=0, max_value=100, value=0)
generate = st.button("🚀 Generate & Download Excel")

# ----------------------------
# GENERATE PROCESS
# ----------------------------
if generate:
    # validate input
    if not uploaded_file and not text_area_story.strip():
        st.error("Please upload a Word document or paste your user story/AC text.")
        st.stop()

    # read full text
    if uploaded_file:
        try:
            full_text = extract_text_from_docx(uploaded_file)
        except Exception as e:
            st.error(f"Error reading DOCX: {e}")
            st.stop()
    else:
        full_text = text_area_story.strip()

    # extract ACs
    ac_blocks = extract_acceptance_criteria(full_text)
    if not ac_blocks:
        st.error("No Acceptance Criteria found in the document.")
        st.stop()

    if max_acs and max_acs > 0:
        ac_blocks = ac_blocks[:max_acs]

    final_rows = []
    progress_bar = st.progress(0)
    total = len(ac_blocks)
    for idx, (ac_label, ac_text) in enumerate(ac_blocks, start=1):
        prompt = build_prompt_for_ac(ac_label, ac_text)
        ai_reply = call_openrouter_api(prompt)
        df_ai = parse_ai_csv_rows_to_df(ai_reply)

        if df_ai is None:
            heuristic = [
                {
                    "Testing Type": "Functional",
                    "Test Scenario": f"Verify core behavior for {ac_label}",
                    "Test Steps": "1. Setup 2. Perform action 3. Verify outcome",
                    "Expected Result": "System behaves as per acceptance criterion"
                },
                {
                    "Testing Type": "Validation/UI",
                    "Test Scenario": f"Verify UI/validations for {ac_label}",
                    "Test Steps": "1. Provide invalid/edge input 2. Observe UI/validation",
                    "Expected Result": "Appropriate validation or UI constraint is enforced"
                },
                {
                    "Testing Type": "Database",
                    "Test Scenario": f"Verify DB persistence for {ac_label}",
                    "Test Steps": "1. Perform action 2. Query DB",
                    "Expected Result": "Data stored and consistent"
                }
            ]
            df_ai = pd.DataFrame(heuristic)

        df_ai = assign_tc_ids(df_ai, ac_label)
        df_ai["Role"] = "Super User"
        df_ai["Acceptance Criteria"] = ac_text
        df_ai = df_ai[["Test Case ID", "Testing Type", "Role", "Acceptance Criteria", "Test Scenario", "Test Steps", "Expected Result"]]
        final_rows.append(df_ai)

        # update progress
        progress_bar.progress(idx / total)

    # combine into one dataframe and persist to session state
    if final_rows:
        st.session_state.result_df = pd.concat(final_rows, ignore_index=True)
    else:
        st.session_state.result_df = pd.DataFrame(columns=["Test Case ID", "Testing Type", "Role", "Acceptance Criteria", "Test Scenario", "Test Steps", "Expected Result"])

    # prepare Excel binary for immediate download
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        st.session_state.result_df.to_excel(writer, index=False, sheet_name="TestCases")
    buffer.seek(0)
    # get bytes once (avoid passing BytesIO instances into st.download_button)
    st.session_state.excel_bytes = buffer.getvalue()

    # show completion and immediate download button
    st.success(f"✅ Generated {len(st.session_state.result_df)} test cases across {len(ac_blocks)} acceptance criteria.")
    # pass raw bytes to download_button (not nested BytesIO)
    st.download_button(
        label="📥 Download Test Cases (Excel)",
        data=st.session_state.excel_bytes,
        file_name="Generated_TestCases.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    progress_bar.empty()

# ----------------------------
# VIEW TABLE (uses session state)
# ----------------------------
if st.button("👁️ View Test Cases"):
    if st.session_state.result_df is not None:
        st.subheader("📊 Generated Test Cases Preview")
        # interactive dataframe - replaced deprecated option with width='stretch'
        st.dataframe(
            st.session_state.result_df,
            width='stretch',
            hide_index=True
        )

        # ----------------------------
        # COPY TO CLIPBOARD BUTTON (as CSV)
        # ----------------------------
        try:
            csv_text = st.session_state.result_df.to_csv(index=False)
            # JSON-encode the CSV so it's safe to embed in JS
            js_csv = json.dumps(csv_text)
            # Create a unique id for the button to avoid collisions
            button_id = "copy-btn"
            copy_html = f"""
            <div>
              <button id="{button_id}">📋 Copy Test Cases to Clipboard</button>
              <span id="copy-status" style="margin-left:10px;"></span>
            </div>
            <script>
            const csv = {js_csv};
            const btn = document.getElementById("{button_id}");
            const status = document.getElementById("copy-status");
            btn.onclick = async () => {{
                try {{
                    await navigator.clipboard.writeText(csv);
                    status.textContent = "Copied ✅";
                    setTimeout(()=>{{ status.textContent = ""; }}, 2000);
                }} catch (err) {{
                    status.textContent = "Copy failed — your browser may block clipboard operations.";
                    console.error(err);
                }}
            }};
            </script>
            """
            components.html(copy_html, height=60)
        except Exception as e:
            st.warning(f"Copy button could not be created: {e}")

        # Download button below the table (recreate bytes in case df changed)
        out_buffer = BytesIO()
        with pd.ExcelWriter(out_buffer, engine="openpyxl") as writer:
            st.session_state.result_df.to_excel(writer, index=False, sheet_name="TestCases")
        out_buffer.seek(0)
        out_bytes = out_buffer.getvalue()

        st.download_button(
            label="📥 Download Test Cases (Excel)",
            data=out_bytes,
            file_name="Generated_TestCases.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No test cases to display. Please generate test cases first.")
