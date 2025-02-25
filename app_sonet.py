import os
import streamlit as st
import pandas as pd
import win32com.client
from io import BytesIO
import tempfile
import anthropic
import time
import re

st.set_page_config(layout="wide", page_title="VBA Macros to C# Converter")

st.title("Excel VBA Macros to C# Converter")
st.markdown("Upload an Excel file to extract VBA macros and convert them to C#.")

# Function to extract VBA code from Excel file
def extract_vba_code(excel_file):
    # Save the uploaded file to a temporary location
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
        tmp.write(excel_file.getvalue())
        temp_path = tmp.name
    
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        
        # Open the workbook
        workbook = excel.Workbooks.Open(temp_path)
        
        vba_modules = {}
        
        # Extract VBA code from each component
        for i in range(1, workbook.VBProject.VBComponents.Count + 1):
            component = workbook.VBProject.VBComponents(i)
            component_name = component.Name
            component_type = component.Type  # 1=Standard, 2=Class, 3=Form
            
            # Get the code from the component
            code_module = component.CodeModule
            line_count = code_module.CountOfLines
            
            if line_count > 0:
                code = code_module.Lines(1, line_count)
                type_name = {1: "Module", 2: "Class", 3: "Form", 11: "ActiveX", 100: "Document"}
                component_type_name = type_name.get(component_type, "Unknown")
                vba_modules[f"{component_name} ({component_type_name})"] = code
        
        # Close Excel
        workbook.Close(False)
        excel.Quit()
        
        # Delete the temporary file
        os.unlink(temp_path)
        
        return vba_modules
    except Exception as e:
        # Close Excel if it's open
        try:
            workbook.Close(False)
            excel.Quit()
        except:
            pass
        
        # Delete the temporary file
        try:
            os.unlink(temp_path)
        except:
            pass
        
        raise Exception(f"Error extracting VBA code: {str(e)}")

# Function to convert VBA code to C# using Claude API
def convert_vba_to_csharp(vba_code, api_key):
    client = anthropic.Anthropic(
        api_key=api_key
    )
    
    prompt = f"""
    You are an expert programmer. I need you to convert the following VBA macro code to C#.
    Please provide only the C# code without any explanation or comments about the conversion.
    
    Here's the VBA code:
    ```vba
    {vba_code}
    ```
    
    Convert this to C# code:
    """
    
    try:
        message = client.messages.create(
            model="claude-3-7-sonnet-20240229",
            max_tokens=4000,
            temperature=0,
            system="You are an expert programmer that specializes in converting Excel VBA macros to C#. Your task is to provide clean, properly formatted C# code that replicates the functionality of the provided VBA code. Include appropriate namespace imports.",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        
        response_content = message.content[0].text
        
        # Extract C# code from response if wrapped in code blocks
        code_pattern = r"```(?:csharp|cs)?\s*([\s\S]*?)```"
        match = re.search(code_pattern, response_content)
        if match:
            return match.group(1).strip()
        else:
            return response_content.strip()
    except Exception as e:
        return f"Error converting code: {str(e)}"

# Create a two-column layout
col1, col2 = st.columns(2)

with col1:
    st.header("Upload Excel File")
    uploaded_file = st.file_uploader("Choose an Excel file (.xlsm, .xls, .xlsx)", type=["xlsm", "xls", "xlsx"])
    
    # API Key input in the sidebar
    with st.sidebar:
        st.header("API Settings")
        api_key = st.text_input("Enter Claude API Key", type="password")
        st.info("Your API key is stored only during this session and not saved anywhere.")
        
        # Display current date/time and user
        st.markdown("---")
        st.subheader("Session Information")
        st.write(f"Current Date and Time: 2025-02-25 07:15:19")
        st.write(f"Current User: alokyadav2020")
    
    if uploaded_file and st.button("Extract VBA Macros"):
        if not api_key:
            st.error("Please enter your Claude API Key in the sidebar before converting.")
        else:
            with st.spinner("Extracting VBA macros..."):
                try:
                    vba_modules = extract_vba_code(uploaded_file)
                    if not vba_modules:
                        st.error("No VBA macros found in the uploaded file.")
                    else:
                        st.session_state.vba_modules = vba_modules
                        st.success(f"Successfully extracted {len(vba_modules)} VBA modules!")
                except Exception as e:
                    st.error(f"Error: {str(e)}")

    # Display VBA modules if available
    if 'vba_modules' in st.session_state and st.session_state.vba_modules:
        st.subheader("VBA Macros")
        selected_vba_module = st.selectbox(
            "Select a VBA module",
            options=list(st.session_state.vba_modules.keys()),
            key="vba_module_selector"
        )
        
        st.code(st.session_state.vba_modules[selected_vba_module], language="vba")
        
        if st.button("Convert Selected Module to C#"):
            if not api_key:
                st.error("Please enter your Claude API Key in the sidebar before converting.")
            else:
                with st.spinner("Converting VBA to C#..."):
                    if 'converted_modules' not in st.session_state:
                        st.session_state.converted_modules = {}
                    
                    module_name = st.session_state.vba_module_selector
                    vba_code = st.session_state.vba_modules[module_name]
                    
                    st.session_state.converted_modules[module_name] = convert_vba_to_csharp(vba_code, api_key)
                    st.success("Conversion completed!")

with col2:
    st.header("C# Code")
    if 'converted_modules' in st.session_state and st.session_state.converted_modules:
        if st.session_state.get("vba_module_selector") in st.session_state.converted_modules:
            selected_module = st.session_state.vba_module_selector
            st.subheader(f"Converted from: {selected_module}")
            st.code(st.session_state.converted_modules[selected_module], language="csharp")
            
            # Download button for C# code
            st.download_button(
                label="Download C# Code",
                data=st.session_state.converted_modules[selected_module],
                file_name=f"{selected_module.split(' ')[0]}.cs",
                mime="text/plain"
            )
    else:
        st.info("Select a VBA module and convert it to see the C# version here.")

# Add footer with instructions
st.markdown("---")
st.markdown("""
### Instructions:
1. Enter your Claude API Key in the sidebar
2. Upload an Excel file containing VBA macros (.xlsm files are recommended)
3. Click 'Extract VBA Macros'
4. Select a module from the dropdown list
5. Click 'Convert Selected Module to C#' to translate the VBA macro to C#
6. Use the download button to save the C# code
""")import os
import streamlit as st
import pandas as pd
import tempfile
import anthropic
import time
import re
from datetime import datetime, timezone
import platform
from io import BytesIO

# Import oletools for VBA extraction
try:
    from oletools.olevba import VBA_Parser
except ImportError:
    st.error("Required package 'oletools' not found. Please install it with: pip install oletools")
    st.stop()

st.set_page_config(layout="wide", page_title="VBA Macros to C# Converter")

st.title("Excel VBA Macros to C# Converter")
st.markdown("Upload an Excel file to extract VBA macros and convert them to C#.")

# Function to extract VBA code from Excel file using oletools
def extract_vba_code(excel_file):
    # Save the uploaded file to a temporary location
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
        tmp.write(excel_file.getvalue())
        temp_path = tmp.name
    
    try:
        vba_modules = {}
        vba_parser = VBA_Parser(temp_path)
        
        if vba_parser.detect_vba_macros():
            # Extract all macros
            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                if vba_code:
                    # Clean up the module name for better display
                    if vba_filename:
                        # Remove path info if present
                        module_name = os.path.basename(vba_filename)
                    else:
                        # Use stream path if no filename
                        module_name = stream_path if stream_path else "Unknown"
                    
                    # Try to determine the type of module
                    module_type = "Module"
                    if "class" in vba_code.lower():
                        module_type = "Class"
                    elif "userform" in vba_code.lower():
                        module_type = "Form"
                    
                    vba_modules[f"{module_name} ({module_type})"] = vba_code
        
        vba_parser.close()
        
        # Delete the temporary file
        try:
            os.unlink(temp_path)
        except:
            pass
        
        return vba_modules
    except Exception as e:
        # Delete the temporary file
        try:
            os.unlink(temp_path)
        except:
            pass
        
        raise Exception(f"Error extracting VBA code: {str(e)}")

# Function to convert VBA code to C# using Claude API
def convert_vba_to_csharp(vba_code, api_key):
    try:
        client = anthropic.Anthropic(
            api_key=api_key
        )
        
        prompt = f"""
        You are an expert programmer. I need you to convert the following VBA macro code to C#.
        Please provide only the C# code without any explanation or comments about the conversion.
        
        Here's the VBA code:
        ```vba
        {vba_code}
        ```
        
        Convert this to C# code:
        """
        
        message = client.messages.create(
            model="claude-3-7-sonnet-20240229",
            max_tokens=4000,
            temperature=0,
            system="You are an expert programmer that specializes in converting Excel VBA macros to C#. Your task is to provide clean, properly formatted C# code that replicates the functionality of the provided VBA code. Include appropriate namespace imports.",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        
        response_content = message.content[0].text
        
        # Extract C# code from response if wrapped in code blocks
        code_pattern = r"```(?:csharp|cs)?\s*([\s\S]*?)```"
        match = re.search(code_pattern, response_content)
        if match:
            return match.group(1).strip()
        else:
            return response_content.strip()
    except Exception as e:
        return f"Error converting code: {str(e)}"

# Create a two-column layout
col1, col2 = st.columns(2)

# Get current time in UTC
current_time = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
user_info = os.environ.get("USERNAME") or os.environ.get("USER") or "Unknown"
system_info = platform.system()

with col1:
    st.header("Upload Excel File")
    uploaded_file = st.file_uploader("Choose an Excel file (.xlsm, .xls, .xlsx)", type=["xlsm", "xls", "xlsx"])
    
    # API Key input directly in the main interface for better visibility
    api_key = st.text_input("Enter Claude API Key", type="password", 
                          help="Your API key is stored only in memory during this session and not saved anywhere.")
    
    # Display system info
    st.sidebar.header("Session Information")
    st.sidebar.write(f"Current Date and Time (UTC): {current_time}")
    st.sidebar.write(f"Current User: {user_info}")
    st.sidebar.write(f"System: {system_info}")
    
    if uploaded_file and st.button("Extract VBA Macros"):
        with st.spinner("Extracting VBA macros..."):
            try:
                vba_modules = extract_vba_code(uploaded_file)
                if not vba_modules:
                    st.error("No VBA macros found in the uploaded file.")
                else:
                    st.session_state.vba_modules = vba_modules
                    st.success(f"Successfully extracted {len(vba_modules)} VBA modules!")
            except Exception as e:
                st.error(f"Error: {str(e)}")
                st.info("If you're having trouble extracting VBA code, make sure your Excel file contains macros. Files with the .xlsm extension typically contain macros.")

    # Display VBA modules if available
    if 'vba_modules' in st.session_state and st.session_state.vba_modules:
        st.subheader("VBA Macros")
        selected_vba_module = st.selectbox(
            "Select a VBA module",
            options=list(st.session_state.vba_modules.keys()),
            key="vba_module_selector"
        )
        
        st.code(st.session_state.vba_modules[selected_vba_module], language="vba")
        
        if st.button("Convert Selected Module to C#"):
            if not api_key:
                st.error("Please enter your Claude API Key before converting.")
            else:
                with st.spinner("Converting VBA to C#..."):
                    if 'converted_modules' not in st.session_state:
                        st.session_state.converted_modules = {}
                    
                    module_name = st.session_state.vba_module_selector
                    vba_code = st.session_state.vba_modules[module_name]
                    
                    st.session_state.converted_modules[module_name] = convert_vba_to_csharp(vba_code, api_key)
                    st.success("Conversion completed!")

with col2:
    st.header("C# Code")
    if 'converted_modules' in st.session_state and st.session_state.converted_modules:
        if st.session_state.get("vba_module_selector") in st.session_state.converted_modules:
            selected_module = st.session_state.vba_module_selector
            st.subheader(f"Converted from: {selected_module}")
            st.code(st.session_state.converted_modules[selected_module], language="csharp")
            
            # Download button for C# code
            st.download_button(
                label="Download C# Code",
                data=st.session_state.converted_modules[selected_module],
                file_name=f"{selected_module.split(' ')[0]}.cs",
                mime="text/plain"
            )
            
            # Option to convert all modules at once
            if len(st.session_state.vba_modules) > 1:
                st.markdown("---")
                if st.button("Convert All Remaining Modules"):
                    if not api_key:
                        st.error("Please enter your Claude API Key before converting.")
                    else:
                        with st.spinner(f"Converting all modules... This may take some time."):
                            for module_name, vba_code in st.session_state.vba_modules.items():
                                if module_name not in st.session_state.converted_modules:
                                    st.session_state.converted_modules[module_name] = convert_vba_to_csharp(vba_code, api_key)
                            st.success(f"Successfully converted all {len(st.session_state.vba_modules)} modules!")
    else:
        st.info("Select a VBA module and convert it to see the C# version here.")

# Add footer with instructions
st.markdown("---")
st.markdown("""
### Instructions:
1. Upload an Excel file containing VBA macros (.xlsm files are recommended)
2. Click 'Extract VBA Macros' to extract all macro code
3. Enter your Claude API Key in the input field
4. Select a module from the dropdown list
5. Click 'Convert Selected Module to C#' to translate the VBA macro to C#
6. Use the download button to save the C# code

### Troubleshooting:
- If extraction fails, ensure your Excel file contains VBA macros
- For large files, the extraction process may take longer
- If conversion fails, check your API key and internet connection
""")