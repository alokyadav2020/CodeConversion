# import streamlit as st
# import tempfile
# import os
# import pandas as pd
# import json
# import time
# from oletools.olevba import VBA_Parser
# import re
# import openpyxl
# import sys
# import pythoncom
# import traceback

# # Import the COM library conditionally
# try:
#     import win32com.client
#     HAS_WIN32COM = True
# except ImportError:
#     HAS_WIN32COM = False
#     st.warning("win32com is not available. Install it using 'pip install pywin32' for better control detection.")

# def extract_control_properties_from_vba(control_name, all_code):
#     """Extract properties of a control from VBA code using regex patterns."""
#     properties = {"name": control_name}
    
#     # Common properties to look for
#     property_patterns = [
#         (r'{}\.Caption\s*=\s*["\'](.*?)["\']'.format(re.escape(control_name)), "Caption"),
#         (r'{}\.Value\s*=\s*["\'](.*?)["\']'.format(re.escape(control_name)), "Value"),
#         (r'{}\.Text\s*=\s*["\'](.*?)["\']'.format(re.escape(control_name)), "Text"),
#         (r'{}\.Height\s*=\s*(\d+)'.format(re.escape(control_name)), "Height"),
#         (r'{}\.Width\s*=\s*(\d+)'.format(re.escape(control_name)), "Width"),
#         (r'{}\.Top\s*=\s*(\d+)'.format(re.escape(control_name)), "Top"),
#         (r'{}\.Left\s*=\s*(\d+)'.format(re.escape(control_name)), "Left"),
#         (r'{}\.Visible\s*=\s*(\w+)'.format(re.escape(control_name)), "Visible"),
#         (r'{}\.Enabled\s*=\s*(\w+)'.format(re.escape(control_name)), "Enabled"),
#     ]
    
#     # Find all event handlers for this control
#     event_handlers = re.findall(r'Sub\s+{}_(\w+)\(.*?\)'.format(re.escape(control_name)), all_code, re.DOTALL)
#     if event_handlers:
#         properties["EventHandlers"] = event_handlers
    
#     # Extract other properties
#     for pattern, prop_name in property_patterns:
#         match = re.search(pattern, all_code, re.IGNORECASE)
#         if match:
#             properties[prop_name] = match.group(1)
    
#     return properties

# def extract_com_properties(obj):
#     """Extract properties from a COM object safely"""
#     properties = {}
    
#     # Common properties to try to extract
#     common_props = [
#         'Name', 'Caption', 'Text', 'Value', 'Height', 'Width', 'Top', 'Left',
#         'Visible', 'Enabled', 'ControlType', 'Type'
#     ]
    
#     for prop in common_props:
#         try:
#             value = getattr(obj, prop, None)
#             if value is not None:
#                 properties[prop] = str(value)
#         except:
#             pass
            
#     return properties

# def extract_controls_from_excel(file_bytes, original_filename):
#     """
#     Extracts information about controls from an Excel file using multiple methods:
#     1. COM automation (if on Windows and pywin32 installed)
#     2. oletools for VBA code analysis
#     3. openpyxl for sheet information
    
#     Returns: tuple of (controls_list, all_code)
#     """
#     ext = os.path.splitext(original_filename)[1]
    
#     # Save the uploaded file temporarily
#     with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
#         tmp.write(file_bytes)
#         temp_path = tmp.name
    
#     controls_list = []
#     all_code = ""
#     sheet_names = []
    
#     # Method 0: Use COM automation if on Windows (best for control detection)
#     if HAS_WIN32COM and sys.platform.startswith('win'):
#         excel_app = None
#         try:
#             # Initialize COM
#             pythoncom.CoInitialize()
            
#             # Create Excel application
#             excel_app = win32com.client.Dispatch("Excel.Application")
#             excel_app.Visible = False
#             excel_app.DisplayAlerts = False
            
#             # Absolute path required
#             abs_path = os.path.abspath(temp_path)
#             workbook = excel_app.Workbooks.Open(abs_path)
            
#             # Get sheet count and iterate through sheets
#             sheet_count = workbook.Sheets.Count
            
#             for sheet_idx in range(1, sheet_count + 1):
#                 sheet = workbook.Sheets(sheet_idx)
#                 sheet_name = sheet.Name
#                 sheet_names.append(sheet_name)
                
#                 # Try to get ActiveX controls (OLEObjects)
#                 try:
#                     if hasattr(sheet, 'OLEObjects'):
#                         ole_count = sheet.OLEObjects.Count
#                         for i in range(1, ole_count + 1):
#                             ole_obj = sheet.OLEObjects(i)
                            
#                             # Get control properties
#                             try:
#                                 properties = extract_com_properties(ole_obj)
                                
#                                 # Try to get the underlying control object
#                                 try:
#                                     control_obj = ole_obj.Object
#                                     control_props = extract_com_properties(control_obj)
#                                     properties.update(control_props)
#                                 except:
#                                     pass
                                
#                                 control_name = properties.get('Name', f"OLEObject_{i}")
#                                 control_type = properties.get('ControlType', 'ActiveX Control')
                                
#                                 control_info = {
#                                     "Control Name": control_name,
#                                     "Type": control_type,
#                                     "Sheet": sheet_name,
#                                     "Source": "COM-OLEObjects",
#                                     "Properties": json.dumps(properties, indent=2)
#                                 }
#                                 controls_list.append(control_info)
#                             except Exception as e:
#                                 st.warning(f"Error getting OLEObject {i} properties: {e}")
#                 except Exception as e:
#                     st.warning(f"Error accessing OLEObjects in {sheet_name}: {e}")
                
#                 # Try to get Form controls (Shapes)
#                 try:
#                     if hasattr(sheet, 'Shapes'):
#                         shapes_count = sheet.Shapes.Count
#                         for i in range(1, shapes_count + 1):
#                             shape = sheet.Shapes(i)
                            
#                             # Check if it's a form control
#                             try:
#                                 # Form controls have a FormControlType property
#                                 control_type = shape.FormControlType
                                
#                                 # Map control type numbers to names
#                                 control_type_names = {
#                                     0: "Button",
#                                     1: "CheckBox", 
#                                     2: "DropDown",
#                                     3: "EditBox",
#                                     4: "GroupBox",
#                                     5: "Label",
#                                     6: "ListBox",
#                                     7: "OptionButton",
#                                     8: "ScrollBar",
#                                     9: "Spinner"
#                                 }
                                
#                                 control_type_name = control_type_names.get(control_type, f"FormControl({control_type})")
                                
#                                 properties = extract_com_properties(shape)
#                                 properties['FormControlType'] = control_type_name
                                
#                                 control_name = properties.get('Name', f"FormControl_{i}")
                                
#                                 control_info = {
#                                     "Control Name": control_name,
#                                     "Type": f"Form Control - {control_type_name}",
#                                     "Sheet": sheet_name,
#                                     "Source": "COM-Shapes",
#                                     "Properties": json.dumps(properties, indent=2)
#                                 }
#                                 controls_list.append(control_info)
#                             except:
#                                 # Not a form control, might be a regular shape
#                                 pass
#                 except Exception as e:
#                     st.warning(f"Error accessing Shapes in {sheet_name}: {e}")
            
#             workbook.Close(SaveChanges=False)
#             excel_app.Quit()
            
#         except Exception as e:
#             st.error(f"COM extraction error: {e}\n{traceback.format_exc()}")
#         finally:
#             # Cleanup
#             if excel_app:
#                 try:
#                     excel_app.Quit()
#                 except:
#                     pass
                    
#             # Release COM
#             pythoncom.CoUninitialize()
    
#     # Method 1: Extract controls from VBA code using oletools
#     try:
#         vba_parser = VBA_Parser(temp_path)
        
#         if vba_parser.detect_vba_macros():
#             for (filename, stream_path, vba_filename, code) in vba_parser.extract_all_macros():
#                 all_code += code + "\n\n"
                
#             # More comprehensive regex patterns for control detection
#             # Form controls
#             form_control_pattern = re.compile(r'\b((?:CommandButton|CheckBox|OptionButton|ListBox|ComboBox|TextBox|ToggleButton|ScrollBar|SpinButton|Label|Image|Frame|OptionGroup)\d+)\b', re.IGNORECASE)
            
#             # ActiveX controls with more types
#             activex_pattern = re.compile(r'\b((?:MSForms\.|Microsoft\.Office\.Tools\.Excel\.)?(?:CommandButton|CheckBox|OptionButton|ListBox|ComboBox|TextBox|ToggleButton|ScrollBar|SpinButton|Label|Image|Frame|TabStrip|MultiPage|RefEdit|DropDown)[^\s\.\(\),]*)\b', re.IGNORECASE)
            
#             # Control references in With blocks
#             with_pattern = re.compile(r'With\s+([^\n]+?)(?:\s+As\s+\w+)?\s*\n(?:.*?\n)*?End With', re.DOTALL)
            
#             # Match control names
#             form_controls = set(form_control_pattern.findall(all_code))
#             activex_controls = set(activex_pattern.findall(all_code))
            
#             # Process With blocks to find more controls
#             with_blocks = with_pattern.findall(all_code)
#             potential_controls = set()
#             for control_ref in with_blocks:
#                 if '.' not in control_ref and not control_ref.startswith('Sheet') and not control_ref.strip().startswith('('):
#                     potential_controls.add(control_ref.strip())
            
#             # Process form controls
#             for control in form_controls:
#                 # Extract properties from VBA code
#                 properties = extract_control_properties_from_vba(control, all_code)
                
#                 # Only add if not already found by COM
#                 if not any(c['Control Name'] == control for c in controls_list):
#                     control_info = {
#                         "Control Name": control,
#                         "Type": "Form Control",
#                         "Source": "VBA Code",
#                         "Properties": json.dumps(properties, indent=2)
#                     }
#                     controls_list.append(control_info)
                
#             # Process ActiveX controls
#             for control in activex_controls:
#                 control_type = "ActiveX"
#                 if "MSForms." in control:
#                     control_type = "MSForms Control"
#                 elif "Microsoft.Office.Tools" in control:
#                     control_type = "VSTO Control"
                    
#                 control_name = control.split('.')[-1] if '.' in control else control
                
#                 # Extract properties from VBA code
#                 properties = extract_control_properties_from_vba(control, all_code)
#                 properties["fullName"] = control
                
#                 # Only add if not already found by COM
#                 if not any(c['Control Name'] == control_name for c in controls_list):
#                     control_info = {
#                         "Control Name": control_name,
#                         "Type": control_type,
#                         "Source": "VBA Code",
#                         "Properties": json.dumps(properties, indent=2)
#                     }
#                     controls_list.append(control_info)
            
#             # Process potential controls from With blocks
#             for control in potential_controls:
#                 if control not in form_controls and not any(control in c for c in activex_controls):
#                     if not any(c['Control Name'] == control for c in controls_list):
#                         properties = extract_control_properties_from_vba(control, all_code)
                        
#                         control_info = {
#                             "Control Name": control,
#                             "Type": "Unknown Control (from With block)",
#                             "Source": "VBA Code",
#                             "Properties": json.dumps(properties, indent=2)
#                         }
#                         controls_list.append(control_info)
        
#         vba_parser.close()
#     except Exception as e:
#         st.warning(f"VBA extraction error: {e}")
    
#     # Method 2: Try using openpyxl to extract sheet information (works for newer Excel formats)
#     if ext.lower() == '.xlsx' or ext.lower() == '.xlsm':
#         try:
#             wb = openpyxl.load_workbook(temp_path, keep_vba=True)
#             if not sheet_names:  # Only append if not already got from COM
#                 sheet_names = wb.sheetnames
#             wb.close()
#         except Exception as e:
#             st.warning(f"openpyxl extraction error: {e}")
    
#     # Method 3: Get sheet names using pandas if needed
#     if not sheet_names:
#         try:
#             excel_file = pd.ExcelFile(temp_path)
#             sheet_names = excel_file.sheet_names
#         except Exception as e:
#             st.warning(f"pandas sheet reading error: {e}")
    
#     # Add sheet information if we haven't added sheets yet
#     if not any(c['Type'] == "Worksheet" for c in controls_list):
#         for sheet_name in sheet_names:
#             sheet_info = {
#                 "Control Name": sheet_name,
#                 "Type": "Worksheet",
#                 "Sheet": sheet_name,
#                 "Source": "Sheet List",
#                 "Properties": json.dumps({"name": sheet_name}, indent=2)
#             }
#             controls_list.append(sheet_info)
            
#     # Clean up
#     try:
#         if os.path.exists(temp_path):
#             os.remove(temp_path)
#     except Exception:
#         pass
    
#     return controls_list, all_code  # Return both the controls list and VBA code

# def main():
#     st.set_page_config(page_title="Excel Controls Extractor", layout="wide")
#     st.title("Excel Controls Extractor")
    
#     # Show platform info
#     platform_info = f"Running on {sys.platform}"
#     if sys.platform.startswith('win'):
#         if HAS_WIN32COM:
#             platform_info += " with COM support (best control detection)"
#         else:
#             platform_info += " without COM support (limited control detection)"
#     else:
#         platform_info += " (limited control detection, COM only works on Windows)"
        
#     st.write(platform_info)
    
#     st.write("""
#     ## Extract All Controls from Excel Files
    
#     This tool analyzes Excel files and extracts information about:
#     - Form controls (buttons, checkboxes, etc.)
#     - ActiveX controls
#     - Named shapes that might be controls
#     - Control properties found in VBA code
#     - Event handlers associated with controls
    
#     Upload your Excel file to begin analysis.
#     """)

#     uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xlsm", "xlsb", "xls"])

#     if uploaded_file is not None:
#         file_bytes = uploaded_file.read()
#         original_filename = uploaded_file.name

#         with st.spinner("Extracting controls from the Excel file..."):
#             controls, all_code = extract_controls_from_excel(file_bytes, original_filename)

#         if controls:
#             st.success(f"Found {len(controls)} controls/objects in the Excel file!")
            
#             # Create DataFrame for display
#             df = pd.DataFrame(controls)
            
#             # Group controls by sheet
#             if "Sheet" in df.columns:
#                 sheet_names = sorted(df["Sheet"].dropna().unique())
#                 if sheet_names:
#                     selected_sheet = st.selectbox("Select Sheet", ["All Sheets"] + list(sheet_names))
                    
#                     if selected_sheet != "All Sheets":
#                         df = df[df["Sheet"] == selected_sheet]
            
#             # Add filter options
#             st.subheader("Filter Controls")
#             filter_type = st.multiselect("Filter by Type", options=sorted(df["Type"].unique()))
            
#             if filter_type:
#                 filtered_df = df[df["Type"].isin(filter_type)]
#             else:
#                 filtered_df = df
                
#             # Display the filtered results
#             st.subheader("Controls Found")
#             st.dataframe(filtered_df)
            
#             # Option to download results as CSV
#             csv = filtered_df.to_csv(index=False)
#             st.download_button(
#                 "Download Results as CSV",
#                 csv,
#                 f"{os.path.splitext(original_filename)[0]}_controls.csv",
#                 "text/csv",
#                 key='download-csv'
#             )
            
#             # Detailed view of selected control
#             st.subheader("Control Details")
#             selected_control = st.selectbox("Select control to see details:", 
#                                           options=filtered_df["Control Name"].tolist())
            
#             if selected_control:
#                 control_data = filtered_df[filtered_df["Control Name"] == selected_control].iloc[0]
#                 st.json(json.loads(control_data["Properties"]))
                
#                 # Show VBA code snippets related to the selected control if it's a form control
#                 try:
#                     if re.match(r'^[A-Za-z]+\d+$', selected_control) and all_code:
#                         vba_snippets = re.findall(r'[^\n]*\b{}\b[^\n]*'.format(re.escape(selected_control)), 
#                                                  all_code, re.IGNORECASE)
#                         if vba_snippets:
#                             st.subheader("Related VBA Code Snippets")
#                             for snippet in vba_snippets[:10]:  # Show max 10 snippets
#                                 st.code(snippet.strip(), language="vb")
#                             if len(vba_snippets) > 10:
#                                 st.info(f"Showing 10 of {len(vba_snippets)} code snippets")
#                 except Exception as e:
#                     st.warning(f"Error displaying code snippets: {e}")
#         else:
#             st.info("No controls found in the uploaded Excel file.")

# if __name__ == "__main__":
#     main()

import streamlit as st
import tempfile
import os
import pandas as pd
import json
import time
from oletools.olevba import VBA_Parser
import re
import zipfile
import xml.etree.ElementTree as ET
import io

def extract_controls_from_excel(file_bytes, original_filename):
    """
    Extract Excel controls using multiple approaches:
    1. Direct XML parsing for .xlsx/.xlsm files
    2. VBA code analysis for potential control references
    3. Sheet structure analysis
    """
    ext = os.path.splitext(original_filename)[1].lower()
    
    # Save the uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(file_bytes)
        temp_path = tmp.name
    
    controls_list = []
    all_code = ""
    
    # APPROACH 1: Direct XML parsing for modern Excel files
    if ext in ['.xlsx', '.xlsm']:
        try:
            # Office files are ZIP archives with XML content
            with zipfile.ZipFile(temp_path, 'r') as z:
                # Extract sheet information
                sheet_names = []
                sheet_ids = {}
                
                # Parse workbook.xml to get sheet names and IDs
                if 'xl/workbook.xml' in z.namelist():
                    with z.open('xl/workbook.xml') as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        
                        # Excel uses namespaces
                        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                        
                        # Get all sheet elements
                        for sheet in root.findall('.//main:sheet', ns) or root.findall('.//sheet'):
                            sheet_name = sheet.get('name', '')
                            sheet_id = sheet.get('sheetId', '')
                            r_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
                            sheet_names.append(sheet_name)
                            sheet_ids[r_id] = {'name': sheet_name, 'id': sheet_id}
                            
                            # Add sheet to control list
                            sheet_info = {
                                "Control Name": sheet_name,
                                "Type": "Worksheet",
                                "Sheet": sheet_name,
                                "Properties": json.dumps({"name": sheet_name, "id": sheet_id}, indent=2)
                            }
                            controls_list.append(sheet_info)
                
                # Look for control-related files
                drawing_sheets = {}
                
                # Check for relationship files that might point to controls
                for filename in z.namelist():
                    # Look for drawing relationships
                    if filename.startswith('xl/drawings/_rels/'):
                        with z.open(filename) as f:
                            tree = ET.parse(f)
                            root = tree.getroot()
                            
                            # Find control references
                            for rel in root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                                rel_type = rel.get('Type', '')
                                if 'control' in rel_type.lower():
                                    target = rel.get('Target', '')
                                    r_id = rel.get('Id', '')
                                    
                                    control_info = {
                                        "Control Name": f"Control_{r_id}",
                                        "Type": "Form Control (XML)",
                                        "Properties": json.dumps({
                                            "RelType": rel_type,
                                            "Target": target,
                                            "Id": r_id
                                        }, indent=2)
                                    }
                                    controls_list.append(control_info)
                
                # Look for actual control files
                for filename in z.namelist():
                    if 'activeX' in filename or 'ctrlProp' in filename:
                        try:
                            with z.open(filename) as f:
                                control_data = f.read()
                                
                                # For binary data, just note its existence
                                control_name = os.path.basename(filename)
                                
                                control_info = {
                                    "Control Name": control_name,
                                    "Type": "ActiveX Control (XML)",
                                    "Properties": json.dumps({
                                        "Path": filename,
                                        "Size": len(control_data)
                                    }, indent=2)
                                }
                                controls_list.append(control_info)
                        except Exception as e:
                            st.warning(f"Error reading control file {filename}: {e}")
                
                # Look for VBA project if it exists (for macro-enabled workbooks)
                vba_found = False
                for filename in z.namelist():
                    if filename == 'xl/vbaProject.bin':
                        vba_found = True
                        control_info = {
                            "Control Name": "VBA Project",
                            "Type": "VBA Container",
                            "Properties": json.dumps({
                                "Path": filename
                            }, indent=2)
                        }
                        controls_list.append(control_info)
                
                # If we found VBA, let's analyze it with oletools
                if vba_found:
                    # oletools will handle this later
                    pass
        
        except Exception as e:
            st.error(f"Error analyzing Excel XML structure: {e}")
    
    # APPROACH 2: VBA Code Analysis
    try:
        vba_parser = VBA_Parser(temp_path)
        
        if vba_parser.detect_vba_macros():
            for (filename, stream_path, vba_filename, code) in vba_parser.extract_all_macros():
                all_code += code + "\n\n"
            
            # Look for forms in the VBA project
            forms = re.findall(r'Begin VB\.Form\s+(\w+)', all_code)
            for form in forms:
                control_info = {
                    "Control Name": form,
                    "Type": "VBA Form",
                    "Properties": json.dumps({
                        "name": form,
                        "source": "VBA Code"
                    }, indent=2)
                }
                controls_list.append(control_info)
            
            # Look for UserForms
            userforms = re.findall(r'Begin VB\.UserForm\s+(\w+)', all_code)
            for form in userforms:
                control_info = {
                    "Control Name": form,
                    "Type": "VBA UserForm",
                    "Properties": json.dumps({
                        "name": form,
                        "source": "VBA Code"
                    }, indent=2)
                }
                controls_list.append(control_info)
            
            # Very thorough pattern to detect control declarations in VBA
            control_pattern = re.compile(r'(?:Dim|Private|Public)\s+(\w+)\s+As\s+(?:MSForms\.)?(CommandButton|CheckBox|OptionButton|ListBox|ComboBox|TextBox|ToggleButton|ScrollBar|SpinButton|Label|Image|Frame|TabStrip|MultiPage|RefEdit|DropDown)', re.IGNORECASE)
            
            control_matches = control_pattern.findall(all_code)
            for name, control_type in control_matches:
                control_info = {
                    "Control Name": name,
                    "Type": f"VBA {control_type}",
                    "Properties": json.dumps({
                        "name": name,
                        "controlType": control_type,
                        "source": "VBA Declaration"
                    }, indent=2)
                }
                controls_list.append(control_info)
            
            # Look for WithEvents declarations (often used for controls)
            withevents_pattern = re.compile(r'(?:Dim|Private|Public)\s+WithEvents\s+(\w+)\s+As\s+(\w+)', re.IGNORECASE)
            
            withevents_matches = withevents_pattern.findall(all_code)
            for name, control_type in withevents_matches:
                control_info = {
                    "Control Name": name,
                    "Type": f"WithEvents {control_type}",
                    "Properties": json.dumps({
                        "name": name,
                        "controlType": control_type,
                        "source": "VBA WithEvents"
                    }, indent=2)
                }
                controls_list.append(control_info)
            
            # Look for controls in specific sheets through code references
            sheet_control_pattern = re.compile(r'(?:Sheet\d+|Worksheets\(\d+\)|Worksheets\(["\']([^"\']+)["\']\))\.Shapes\(["\']([^"\']+)["\']\)', re.IGNORECASE)
            
            sheet_control_matches = sheet_control_pattern.findall(all_code)
            for sheet_name, control_name in sheet_control_matches:
                sheet_name = sheet_name or "Unknown Sheet"
                control_info = {
                    "Control Name": control_name,
                    "Type": "Sheet Shape/Control",
                    "Sheet": sheet_name,
                    "Properties": json.dumps({
                        "name": control_name,
                        "sheet": sheet_name,
                        "source": "VBA Code Reference"
                    }, indent=2)
                }
                controls_list.append(control_info)
        
        vba_parser.close()
    except Exception as e:
        st.warning(f"VBA analysis error: {e}")
    
    # APPROACH 3: Try to read sheets with pandas for form field detection
    try:
        xls = pd.ExcelFile(temp_path)
        for sheet_name in xls.sheet_names:
            try:
                # Just check if we can read the sheet - some sheets with controls might fail
                df = pd.read_excel(xls, sheet_name)
                
                # We successfully read the sheet, but we'll only add it if it's not already in the list
                if not any(c.get("Control Name") == sheet_name and c.get("Type") == "Worksheet" for c in controls_list):
                    sheet_info = {
                        "Control Name": sheet_name,
                        "Type": "Worksheet",
                        "Sheet": sheet_name,
                        "Properties": json.dumps({"name": sheet_name}, indent=2)
                    }
                    controls_list.append(sheet_info)
            except Exception as e:
                # If reading fails, it might be due to controls/objects in the sheet
                if not any(c.get("Control Name") == sheet_name for c in controls_list):
                    sheet_info = {
                        "Control Name": sheet_name,
                        "Type": "Worksheet (with possible form controls)",
                        "Sheet": sheet_name,
                        "Properties": json.dumps({
                            "name": sheet_name, 
                            "error": str(e),
                            "note": "Failed to read with pandas, might contain controls"
                        }, indent=2)
                    }
                    controls_list.append(sheet_info)
    except Exception as e:
        st.warning(f"Pandas analysis error: {e}")
    
    # Clean up the temporary file
    try:
        os.remove(temp_path)
    except:
        pass
    
    # If we found no controls but have sheets, note this in the UI
    if not any(c.get("Type") != "Worksheet" for c in controls_list) and controls_list:
        st.info("No controls were found in this workbook, only worksheets.")
    
    return controls_list, all_code

def main():
    st.set_page_config(page_title="Excel Controls Extractor", layout="wide")
    st.title("Excel Controls Extractor")
    
    st.write("""
    ## Extract Controls from Excel Files
    
    This tool analyzes Excel files and extracts information about controls such as:
    - Form controls (buttons, checkboxes, etc.)
    - ActiveX controls
    - VBA Forms and UserForms
    - Control references in VBA code
    
    Upload your Excel file to begin analysis.
    """)
    
    # Debug mode checkbox
    debug_mode = st.checkbox("Enable Debug Mode")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xlsm", "xlsb", "xls"])

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        original_filename = uploaded_file.name
        
        if debug_mode:
            st.write(f"File name: {original_filename}")
            st.write(f"File size: {len(file_bytes)} bytes")

        with st.spinner("Analyzing Excel file for controls..."):
            controls, all_code = extract_controls_from_excel(file_bytes, original_filename)

        if controls:
            st.success(f"Found {len(controls)} items in the Excel file.")
            
            # Create DataFrame for display
            df = pd.DataFrame(controls)
            
            # Filter controls (not worksheets) for quick analysis
            real_controls = df[df["Type"] != "Worksheet"]
            if not real_controls.empty:
                st.subheader(f"Controls found: {len(real_controls)}")
                st.dataframe(real_controls)
            
            # Add filter options
            st.subheader("All Items (with filters)")
            filter_type = st.multiselect("Filter by Type", options=sorted(df["Type"].unique()))
            
            if filter_type:
                filtered_df = df[df["Type"].isin(filter_type)]
            else:
                filtered_df = df
                
            # Display the filtered results
            st.dataframe(filtered_df,use_container_width=True)
            
            # Option to download results as CSV
            csv = filtered_df.to_csv(index=False)
            st.download_button(
                "Download Results as CSV",
                csv,
                f"{os.path.splitext(original_filename)[0]}_controls.csv",
                "text/csv",
                key='download-csv'
            )
            
            # Detailed view of selected control
            st.subheader("Control Details")
            selected_control = st.selectbox("Select item to see details:", 
                                          options=filtered_df["Control Name"].tolist())
            
            if selected_control:
                control_data = filtered_df[filtered_df["Control Name"] == selected_control].iloc[0]
                st.json(json.loads(control_data["Properties"]))
                
                # Show VBA code snippets related to the selected control
                if all_code and selected_control != "VBA Project" and not selected_control.startswith("Sheet"):
                    try:
                        vba_snippets = re.findall(r'[^\n]*\b{}\b[^\n]*'.format(re.escape(selected_control)), 
                                                 all_code, re.IGNORECASE)
                        if vba_snippets:
                            st.subheader("Related VBA Code Snippets")
                            for snippet in vba_snippets[:10]:  # Show max 10 snippets
                                st.code(snippet.strip(), language="vb")
                            if len(vba_snippets) > 10:
                                st.info(f"Showing 10 of {len(vba_snippets)} code snippets")
                    except Exception as e:
                        if debug_mode:
                            st.warning(f"Error displaying code snippets: {e}")
            
            # Debug output
            if debug_mode and all_code:
                st.subheader("Debug: All VBA Code")
                st.code(all_code[:10000] + ("..." if len(all_code) > 10000 else ""), language="vb")
        else:
            st.info("No controls or worksheets found in the uploaded Excel file.")

if __name__ == "__main__":
    main()