import streamlit as st
import tempfile
import os
from openai import OpenAI
from oletools.olevba import VBA_Parser

# Set page configuration to wide layout
st.set_page_config(layout="wide", page_title="Excel VBA to C# Converter", initial_sidebar_state="expanded")

def extract_vba_from_excel(file_bytes, original_filename):
    """
    Extract VBA macro code from an Excel file using oletools' VBA_Parser.
    """
    ext = os.path.splitext(original_filename)[1]
    vba_code = ""

    # Save the uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    # Extract VBA macros
    try:
        vba_parser = VBA_Parser(tmp_path)
        if vba_parser.detect_vba_macros():
            for (filename, stream_path, vba_filename, code) in vba_parser.extract_all_macros():
                vba_code += f"' Macro from {vba_filename} in {filename}\n" + code + "\n\n"
        else:
            vba_code = "No VBA macros found in the uploaded file."
        vba_parser.close()
    except Exception as e:
        vba_code = f"Error extracting VBA code: {e}"
    finally:
        os.remove(tmp_path)

    return vba_code

def convert_vba_to_csharp(vba_code, api_key):
    """
    Use OpenAI GPT-4-turbo (o3) to convert VBA macro code into C#.
    """
    if not vba_code.strip() or vba_code.startswith("Error"):
        return "No valid VBA code found for conversion."
    
    # Set the API key dynamically
    # openai.api_key = api_key
    client = OpenAI(api_key=api_key)

    prompt = f"""
    You are an expert in converting VBA (Visual Basic for Applications) macros to C#.
    Convert the following VBA code into C# with appropriate syntax and best practices:

    VBA CODE:
    {vba_code}

    OUTPUT C# CODE:
    """
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a highly skilled C# developer with expertise in VBA conversion."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1
        )
        csharp_code = response.choices[0].message.content
    except Exception as e:
        csharp_code = f"Error converting VBA to C#: {e}"

    return csharp_code

def main():
    st.title("Excel VBA to C# Converter (Powered by OpenAI)")
    st.write(
        """
        **How it Works:**
        1. Enter your OpenAI API Key in the sidebar.
        2. Upload an Excel file (.xlsm, .xlsb, or .xls) with VBA macros.
        3. The app extracts VBA code from the file.
        4. The VBA code is sent to OpenAI GPT-4-turbo for conversion.
        5. Both the VBA and converted C# code are displayed side-by-side.
        """
    )

    # Sidebar: API key input
    api_key = st.sidebar.text_input("Enter your OpenAI API Key", type="password")
    if not api_key:
        st.sidebar.warning("Please enter your OpenAI API key to enable conversion.")

    # File uploader in main UI
    uploaded_file = st.file_uploader("Upload an Excel file with VBA macros", type=["xlsm", "xlsb", "xls"])

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        original_filename = uploaded_file.name

        with st.spinner("Extracting VBA code..."):
            vba_code = extract_vba_from_excel(file_bytes, original_filename)

        # Only attempt conversion if an API key is provided
        if api_key:
            with st.spinner("Converting VBA to C#..."):
                csharp_code = convert_vba_to_csharp(vba_code, api_key)
        else:
            csharp_code = "OpenAI API key not provided. Please enter your API key in the sidebar."

        # Display results in two columns
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Extracted VBA Code")
            st.code(vba_code, language="vb")
        with col2:
            st.subheader("Converted C# Code")
            st.code(csharp_code, language="csharp")

if __name__ == "__main__":
    main()
