import streamlit as st
import os
import sys
import subprocess
from pathlib import Path

st.set_page_config(page_title="News-to-PPT Generator", layout="centered")
st.title("📊 Automated News-to-PPT Generator")
st.markdown("Generate a business news PowerPoint report for any company in seconds.")

# --- Input Form ---
with st.form("input_form"):
    company = st.text_input("Company Name", "")
    company_type = st.selectbox("Type", ["competitor", "potential customer"])
    article_count = st.number_input("Number of Articles", min_value=1, max_value=50, value=10, step=1)
    submitted = st.form_submit_button("Run Pipeline 🚀")

ppt_path = None
status_messages = []

# --- Run Pipeline Logic ---
def run_pipeline_cli(company, company_type, article_count):
    """Run the CLI pipeline script with parameters, capture output."""
    script = "run_full_ppt_pipeline.py"
    if not Path(script).exists():
        return None, [f"❌ Script not found: {script}"]
    # Prepare simulated input for the script
    inputs = f"{company}\n{company_type}\n{article_count}\n"
    # Run as subprocess
    try:
        result = subprocess.run(
            [sys.executable, script],
            input=inputs,
            capture_output=True,
            text=True,
            timeout=300
        )
        stdout = result.stdout
        stderr = result.stderr
        # Try to find ppt path in output (robust: look for .pptx anywhere in line)
        ppt_path = None
        for line in stdout.splitlines():
            if ".pptx" in line.lower():
                # Extract the path after the last ':' or after 'PPT generated:'
                if ":" in line:
                    ppt_path_candidate = line.split(":", 1)[-1].strip()
                else:
                    ppt_path_candidate = line.strip()
                # If relative, resolve to absolute
                if ppt_path_candidate and os.path.exists(ppt_path_candidate):
                    ppt_path = os.path.abspath(ppt_path_candidate)
                    break
                # Try to resolve if path is printed but not existing
                if ppt_path_candidate and ppt_path_candidate.endswith(".pptx"):
                    # Check typical output directory
                    candidate = os.path.join("outputs", "full_pipeline", company.replace(" ", "_"), "ppt", os.path.basename(ppt_path_candidate))
                    if os.path.exists(candidate):
                        ppt_path = os.path.abspath(candidate)
                        break
        # Fallback: search expected output directories if not found in stdout
        if not ppt_path:
            # Try both 'ppt' and direct company folder for .pptx
            search_dirs = [
                os.path.join("outputs", "full_pipeline", company.replace(" ", "_"), "ppt"),
                os.path.join("outputs", "full_pipeline", company.replace(" ", "_"))
            ]
            for ppt_dir in search_dirs:
                if os.path.isdir(ppt_dir):
                    for file in os.listdir(ppt_dir):
                        if file.lower().endswith(".pptx") and company.lower().replace(" ", "_") in file.lower():
                            ppt_path = os.path.abspath(os.path.join(ppt_dir, file))
                            break
                if ppt_path:
                    break
        # Add clear error if not found
        if not ppt_path:
            stderr += f"\n[ERROR] PPT file not found in expected locations: {search_dirs}"
        return ppt_path, [stdout, stderr, f"[DEBUG] Resolved PPT path: {ppt_path if ppt_path else 'None'}"]
    except Exception as e:
        return None, [f"❌ Error running pipeline: {e}"]

# --- Main App Logic ---
if submitted:
    if not company.strip():
        st.error("Please enter a company name.")
    else:
        with st.spinner("Running pipeline. This may take up to a minute..."):
            ppt_path, status_messages = run_pipeline_cli(company.strip(), company_type, article_count)
        st.subheader("Status & Output")
        for msg in status_messages:
            if msg:
                st.code(msg, language="text")
        if ppt_path and os.path.exists(ppt_path):
            st.success(f"PPT generated: {ppt_path}")
            with open(ppt_path, "rb") as f:
                st.download_button("Download PowerPoint Report", f, file_name=os.path.basename(ppt_path), mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        else:
            st.warning("No PPT file found. Please check the status messages above.")

st.markdown("---")
st.caption("Made with ❤️ using Streamlit. Contact your developer for support.")
