

import streamlit as st
import tempfile
import os
from openai import AzureOpenAI  # Changed from OpenAI to AzureOpenAI
from oletools.olevba import VBA_Parser
from azure.identity import DefaultAzureCredential, get_bearer_token_provider

# Set page configuration to wide layout


def extract_vba_from_excel(file_bytes, original_filename):
    """
    Extract VBA macro code from an Excel file using oletools' VBA_Parser.
    """
    # This function remains unchanged
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

def convert_vba_to_csharp(vba_code,prompt_ ,api_key=st.secrets['api_key'], api_endpoint=st.secrets['api_endpoint'], deployment_name=st.secrets['deployment_name']):
    """
    Use Azure OpenAI to convert VBA macro code into C#.
    """
    if not vba_code.strip() or vba_code.startswith("Error"):
        return "No valid VBA code found for conversion."
    # print(api_key)
    # print(api_endpoint)
    # print(deployment_name)
     
    deployment = os.getenv("DEPLOYMENT_NAME", "gpt-4o")  
      
# Initialize Azure OpenAI Service client with Entra ID authentication
   
    # Initialize Azure OpenAI client
    client = AzureOpenAI(
        api_key= os.getenv("AZURE_OPENAI_API_KEY", api_key),  
        # api_version="2024-11-20",  # Using a standard Azure OpenAI API version
        azure_endpoint=os.getenv("ENDPOINT_URL",api_endpoint),
        # azure_ad_token_provider=token_provider,  
        api_version="2024-05-01-preview",
    )

    # prompt = f"""
    # You are an expert in converting VBA (Visual Basic for Applications) macros to C#.
    # Convert the following VBA code into C# with appropriate syntax and best practices:

    # VBA CODE:
    # {vba_code}

    # """
    
    try:
        response = client.chat.completions.create(
            model=deployment,  # Use the deployment name instead of model name
            messages=[
                {"role": "system", "content": "You are a highly skilled C# developer with expertise in VBA conversion."},
                {"role": "user", "content": prompt_}
            ],
            temperature=0.1
        )
        csharp_code = response.choices[0].message.content
    except Exception as e:
        csharp_code = f"Error converting : {e}"

    return csharp_code

def main_vba_code_converter():
    st.set_page_config(layout="wide", page_title="Excel VBA to C# Converter", initial_sidebar_state="expanded")
    st.title("Excel VBA to C# Converter (Powered by Azure OpenAI)")
    st.write(
        """
        **How it Works:**
        1. Enter your Azure OpenAI configuration in the sidebar.
        2. Upload an Excel file (.xlsm, .xlsb, or .xls) with VBA macros.
        3. The app extracts VBA code from the file.
        4. The VBA code is sent to Azure OpenAI for conversion.
        5. Both the VBA and converted C# code are displayed side-by-side.
        """
    )


    
    # Add a hint about one of the provided keys
    st.sidebar.markdown("**Hint:** You can use one of the provided API keys.")
    
    # if not api_key or not api_endpoint or not deployment_name:
    #     st.sidebar.warning("Please fill in all Azure OpenAI configuration fields.")

    # File uploader in main UI
    uploaded_file = st.file_uploader("Upload an Excel file with VBA macros", type=["xlsm", "xlsb", "xls"])

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        original_filename = uploaded_file.name
        

# Let the user edit the prompt
        prompt_text = st.text_area(
            "Prompt for Conversion",
            value=(
    "You are an expert in converting VBA (Visual Basic for Applications) macros to C#.\n"
    "Convert the following VBA code into C# with appropriate syntax and best practices:\n"
),
            height=120
        )           
        if st.button("Convert VBA"):
            with st.spinner("Extracting VBA code..."):
                vba_code = extract_vba_from_excel(file_bytes, original_filename)
                csharp_code = convert_vba_to_csharp(vba_code, prompt_=f"{prompt_text} VBA Code:{vba_code}")

                # Display results in two columns
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Extracted VBA Code")
                    st.code(vba_code, language="vb")
                with col2:
                    st.subheader("Converted Code")
                    st.code(csharp_code)







if __name__ == "__main__":
    main_vba_code_converter()