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
import traceback # Import for detailed error logging

# Define common XML namespaces (no changes here)
NAMESPACES = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'ole': 'http://schemas.openxmlformats.org/oleObject/2006',
    'ax': 'http://schemas.microsoft.com/office/drawing/2016/11/diagram', 
    'wne': 'http://schemas.microsoft.com/office/excel/2006/main'
}

def _get_ns_tag(tag_string, ns_map): # No changes here
    if ':' in tag_string:
        parts = tag_string.split(':')
        prefix = parts[0]
        name = parts[1]
        if prefix in ns_map:
            return f"{{{ns_map[prefix]}}}{name}"
    return tag_string


def extract_controls_from_excel(file_bytes, original_filename, debug_mode): # Added debug_mode parameter
    ext = os.path.splitext(original_filename)[1].lower()
    
    temp_path = None # Initialize temp_path
    try:
        # Create the temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(file_bytes)
            temp_path = tmp.name
        
        controls_list = []
        all_vba_code = "" 
        
        # APPROACH 1: Direct XML parsing for modern Excel files (.xlsx, .xlsm)
        if ext in ['.xlsx', '.xlsm']:
            try:
                with zipfile.ZipFile(temp_path, 'r') as z: # zipfile is correctly handled with 'with'
                    archive_files = z.namelist()
                    
                    # --- 1.1 Parse workbook.xml --- (Code remains the same)
                    sheet_rId_to_name = {}
                    sheet_target_to_display_name = {} 
                    workbook_rels_path = 'xl/_rels/workbook.xml.rels'
                    sheet_rId_to_target_path = {}

                    if workbook_rels_path in archive_files:
                        with z.open(workbook_rels_path) as f_rels:
                            rels_tree = ET.parse(f_rels)
                            for rel_elem in rels_tree.getroot().findall(_get_ns_tag('Relationship', NAMESPACES), NAMESPACES):
                                if "worksheet" in rel_elem.get('Type', ''):
                                    sheet_rId_to_target_path[rel_elem.get('Id')] = rel_elem.get('Target')
                    
                    if 'xl/workbook.xml' in archive_files:
                        with z.open('xl/workbook.xml') as f:
                            tree = ET.parse(f)
                            for sheet_elem in tree.getroot().findall(_get_ns_tag('main:sheets/main:sheet', NAMESPACES), NAMESPACES):
                                name = sheet_elem.get('name')
                                r_id = sheet_elem.get(_get_ns_tag('r:id', NAMESPACES))
                                if name and r_id:
                                    sheet_rId_to_name[r_id] = name
                                    target_path = sheet_rId_to_target_path.get(r_id)
                                    if target_path:
                                        full_target_path = os.path.join('xl', target_path).replace('\\', '/')
                                        sheet_target_to_display_name[full_target_path] = name

                                    controls_list.append({
                                        "Control Name": name, "Type": "Worksheet", "Sheet": name,
                                        "Properties": json.dumps({"name": name, "r:id": r_id, "target": target_path or "N/A"}, indent=2),
                                        "Source": "Workbook XML"
                                    })
                    
                    # --- 1.2 Map sheets to their drawing files --- (Code remains the same)
                    sheet_display_name_to_drawing_path = {}
                    for sheet_file_path_in_zip, display_name in sheet_target_to_display_name.items():
                        sheet_rels_path = os.path.join(os.path.dirname(sheet_file_path_in_zip), '_rels', os.path.basename(sheet_file_path_in_zip) + '.rels').replace('\\', '/')
                        if sheet_rels_path in archive_files:
                            with z.open(sheet_rels_path) as f_srels:
                                srels_tree = ET.parse(f_srels)
                                for rel_elem in srels_tree.getroot().findall(_get_ns_tag('Relationship', NAMESPACES), NAMESPACES):
                                    if "drawing" in rel_elem.get('Type', ''):
                                        drawing_target = rel_elem.get('Target')
                                        drawing_path_abs = os.path.normpath(os.path.join(os.path.dirname(sheet_rels_path), drawing_target)).replace('\\','/')
                                        sheet_display_name_to_drawing_path[display_name] = drawing_path_abs
                                        break
                    
                    # --- 1.3 Parse drawing files for controls --- (Code remains the same)
                    for sheet_name, drawing_path in sheet_display_name_to_drawing_path.items():
                        if drawing_path in archive_files:
                            with z.open(drawing_path) as f_drawing:
                                drawing_tree = ET.parse(f_drawing)
                                drawing_root = drawing_tree.getroot()
                                for anchor_elem in drawing_root.findall('.//xdr:twoCellAnchor', NAMESPACES) + \
                                                   drawing_root.findall('.//xdr:oneCellAnchor', NAMESPACES) + \
                                                   drawing_root.findall('.//xdr:absoluteAnchor', NAMESPACES):
                                    control_name_xml = None
                                    control_type_detail = "Shape/Object"
                                    properties = {"sheet": sheet_name, "xml_source_path": drawing_path}
                                    cNvPr_elem = anchor_elem.find('.//xdr:cNvPr', NAMESPACES)
                                    if cNvPr_elem is not None:
                                        control_name_xml = cNvPr_elem.get('name')
                                        if cNvPr_elem.get('descr'): properties['description'] = cNvPr_elem.get('descr')

                                    sp_elem = anchor_elem.find('./xdr:sp', NAMESPACES)
                                    if sp_elem is not None:
                                        control_type_detail = "Shape"
                                        if sp_elem.get('macro'):
                                            properties['macro_assigned'] = sp_elem.get('macro')
                                            control_type_detail = "Shape with Macro"
                                        clientData_elem = sp_elem.find('./xdr:clientData', NAMESPACES)
                                        if clientData_elem is not None:
                                            control_type_detail = "Form Control (Legacy)"
                                            ctrlPr_elem = clientData_elem.find('.//wne:ctrlPr', NAMESPACES)
                                            if ctrlPr_elem is not None:
                                                ctrl_prop_rid = ctrlPr_elem.get(_get_ns_tag('r:id', NAMESPACES))
                                                if ctrl_prop_rid: properties['ctrlProp_rId'] = ctrl_prop_rid
                                    
                                    control_tag_elem = anchor_elem.find('./xdr:control', NAMESPACES)
                                    if control_tag_elem is not None:
                                        if control_tag_elem.get('name'): control_name_xml = control_tag_elem.get('name')
                                        properties['shapeId_xml'] = control_tag_elem.get('shapeId')
                                        activex_rid = control_tag_elem.get(_get_ns_tag('r:id', NAMESPACES))
                                        if activex_rid:
                                            drawing_rels_path = os.path.join(os.path.dirname(drawing_path), '_rels', os.path.basename(drawing_path) + '.rels').replace('\\', '/')
                                            if drawing_rels_path in archive_files:
                                                with z.open(drawing_rels_path) as f_drels:
                                                    drels_tree = ET.parse(f_drels)
                                                    for rel in drels_tree.getroot().findall(_get_ns_tag('Relationship', NAMESPACES), NAMESPACES):
                                                        if rel.get('Id') == activex_rid and "control" in rel.get('Type', '').lower():
                                                            activex_target = rel.get('Target') 
                                                            activex_path_abs = os.path.normpath(os.path.join(os.path.dirname(drawing_rels_path), activex_target)).replace('\\','/')
                                                            properties['activeX_definition_path'] = activex_path_abs
                                                            if activex_path_abs in archive_files and activex_path_abs.endswith('.xml'):
                                                                with z.open(activex_path_abs) as f_ax_prop:
                                                                    ax_prop_tree = ET.parse(f_ax_prop)
                                                                    ax_control_elem = ax_prop_tree.find('.//ax:axControl', NAMESPACES)
                                                                    if ax_control_elem is not None:
                                                                        prog_id = ax_control_elem.get('progId')
                                                                        properties['progId'] = prog_id
                                                                        control_type_detail = f"ActiveX ({prog_id})" if prog_id else "ActiveX Control"
                                                                        ax_props_data = {}
                                                                        for prop_elem in ax_control_elem.findall('./ax:prop', NAMESPACES):
                                                                            ax_props_data[prop_elem.get('name')] = prop_elem.get('val')
                                                                        if ax_props_data: properties['activeX_properties_xml'] = ax_props_data
                                                            break
                                    
                                    graphicFrame_elem = anchor_elem.find('./xdr:graphicFrame', NAMESPACES)
                                    if graphicFrame_elem is not None:
                                        oleObj_elem = graphicFrame_elem.find('.//mc:Fallback/oleObj', NAMESPACES)
                                        if oleObj_elem is not None:
                                            prog_id = oleObj_elem.get('progId')
                                            if prog_id:
                                                control_type_detail = f"ActiveX ({prog_id})"
                                                properties['progId'] = prog_id
                                                if not control_name_xml: control_name_xml = oleObj_elem.get('name', f"Unnamed ActiveX {prog_id}")
                                            ole_rid = oleObj_elem.get(_get_ns_tag('r:id', NAMESPACES))
                                            if ole_rid: properties['ole_object_rId'] = ole_rid

                                    if control_name_xml:
                                        controls_list.append({
                                            "Control Name": control_name_xml, "Type": control_type_detail, "Sheet": sheet_name,
                                            "Properties": json.dumps(properties, indent=2, ensure_ascii=False), "Source": "XML Drawing"
                                        })
                    
                    if 'xl/vbaProject.bin' in archive_files:
                        controls_list.append({
                            "Control Name": "VBA Project", "Type": "VBA Container", "Sheet": None,
                            "Properties": json.dumps({"Path": "xl/vbaProject.bin"}, indent=2), "Source": "File Structure"
                        })
            
            except Exception as e_xml:
                st.error(f"Error analyzing Excel XML structure: {e_xml}", icon="⚠️")
                if debug_mode:
                    st.error(f"XML Traceback: {traceback.format_exc()}", icon="🐛")


        # APPROACH 2: VBA Code Analysis (using oletools)
        vba_parser = None # Initialize vba_parser to None
        try:
            if os.path.exists(temp_path) and os.path.getsize(temp_path) > 0:
                 # VBA_Parser can fail for non-OLE files or corrupted files
                vba_parser = VBA_Parser(temp_path)
                if vba_parser.detect_vba_macros():
                    macros_data = []
                    try:
                        macros_data = list(vba_parser.extract_all_macros())
                    except Exception as e_olevba_extract:
                        st.warning(f"oletools: Could not extract all macros: {e_olevba_extract}", icon="⚠️")

                    for (vba_filename, stream_path, ole_filename, code) in macros_data:
                        all_vba_code += f"'--- Source: {vba_filename} (Stream: {stream_path}, OLE File: {ole_filename}) ---\n{code}\n\n"
                    
                    if all_vba_code:
                        # Form and control detection regex remains the same
                        form_pattern = re.compile(r'Begin\s+VB\.(UserForm|Form)\s+(\w+)', re.IGNORECASE)
                        for match in form_pattern.finditer(all_vba_code):
                            form_type, form_name = match.groups()
                            controls_list.append({
                                "Control Name": form_name, "Type": f"VBA {form_type}", "Sheet": None,
                                "Properties": json.dumps({"name": form_name, "source_type": form_type}, indent=2),
                                "Source": "VBA Code (Form Definition)"
                            })

                        control_types_vba = [
                            "CommandButton", "CheckBox", "OptionButton", "ListBox", "ComboBox",
                            "TextBox", "Label", "Image", "ToggleButton", "Frame", "MultiPage",
                            "TabStrip", "ScrollBar", "SpinButton", "RefEdit", "DropDown"
                        ]
                        control_decl_pattern_str = r'(?:Dim|Private|Public)(\s+WithEvents)?\s+(\w+)\s+As\s+(?:MSForms\.)?(' + '|'.join(control_types_vba) + r')'
                        control_decl_pattern = re.compile(control_decl_pattern_str, re.IGNORECASE)
                        
                        for match in control_decl_pattern.finditer(all_vba_code):
                            withevents, ctrl_name, ctrl_type = match.groups()
                            type_prefix = "WithEvents " if withevents else ""
                            controls_list.append({
                                "Control Name": ctrl_name, "Type": f"VBA {type_prefix}{ctrl_type}", "Sheet": None,
                                "Properties": json.dumps({"name": ctrl_name, "controlType": ctrl_type, "declaration": match.group(0).strip()}, indent=2),
                                "Source": "VBA Code (Declaration)"
                            })
                        
                        simple_sheet_control_pattern = re.compile(
                            r'(?P<sheet_id>Worksheets\s*\(\s*["\']?([^"\')]+)["\']?\s*\)|Sheets\s*\(\s*["\']?([^"\')]+)["\']?\s*\)|Sheet\d+|ThisWorkbook)\s*\.'
                            r'(?:(?P<collection>OLEObjects|Shapes|Controls)\s*\(\s*["\']?(?P<control_name_in_collection>[^"\')]+)["\']?\s*\)|(?P<control_name_direct>[A-Za-z_][\w]*))',
                            re.IGNORECASE
                        )
                        for match in simple_sheet_control_pattern.finditer(all_vba_code):
                            data = match.groupdict()
                            sheet_identifier = data.get('sheet_id','Unknown Sheet').replace('Worksheets(','').replace('Sheets(','').replace(')','').replace('"','').replace("'", "")
                            control_name_vba = data.get('control_name_in_collection') or data.get('control_name_direct')
                            collection_type = data.get('collection')
                            
                            if control_name_vba:
                                type_suffix = f" ({collection_type} Ref)" if collection_type else " (Direct Ref)"
                                controls_list.append({
                                    "Control Name": control_name_vba, "Type": f"VBA Sheet Control Ref{type_suffix}", "Sheet": sheet_identifier,
                                    "Properties": json.dumps({"name": control_name_vba, "sheet_ref_vba": sheet_identifier, "full_match": match.group(0).strip()}, indent=2),
                                    "Source": "VBA Code (Sheet Reference)"
                                })
                elif ext in ['.xlsm', '.xls', '.xlsb']: # If detect_vba_macros is False
                    st.info("No VBA macros detected by oletools.")
            elif ext in ['.xlsm', '.xls', '.xlsb']: # If file is empty or doesn't exist but is macro-enabled type
                st.info(f"File '{original_filename}' (type: {ext}) appears unsuitable for VBA analysis (e.g., empty or not an OLE container).")
        
        except Exception as e_vba: # Catch errors from VBA_Parser() or other issues
            st.warning(f"VBA analysis error: {e_vba}", icon="⚠️")
            if debug_mode:
                st.warning(f"VBA Traceback: {traceback.format_exc()}", icon="🐛")
        finally:
            if vba_parser: # Ensure vba_parser was instantiated
                try:
                    vba_parser.close()
                except Exception as e_close_vba:
                    # Log this error, but don't let it prevent file deletion
                    st.warning(f"Error closing VBA_Parser: {e_close_vba}", icon="🐛")


        # APPROACH 3: Pandas for basic sheet listing
        if ext == '.xls' or not any(c.get("Type") == "Worksheet" for c in controls_list):
            try:
                # Use 'with' statement for pd.ExcelFile to ensure it's closed
                with pd.ExcelFile(temp_path) as xls_file:
                    for sheet_name_pd in xls_file.sheet_names:
                        if not any(c.get("Control Name") == sheet_name_pd and c.get("Type") == "Worksheet" for c in controls_list):
                            sheet_type_pd = "Worksheet"
                            props_pd = {"name": sheet_name_pd}
                            try:
                                pd.read_excel(xls_file, sheet_name_pd, nrows=1) 
                            except Exception as e_read_pd:
                                sheet_type_pd = "Worksheet (potential objects)"
                                props_pd["read_error"] = str(e_read_pd)
                                props_pd["note"] = "Pandas failed to read sheet, may contain complex objects/controls."
                            
                            controls_list.append({
                                "Control Name": sheet_name_pd, "Type": sheet_type_pd, "Sheet": sheet_name_pd,
                                "Properties": json.dumps(props_pd, indent=2), "Source": "Pandas Sheet Scan"
                            })
            except Exception as e_pandas:
                st.warning(f"Pandas analysis error for sheet listing: {e_pandas}", icon="⚠️")
                if debug_mode:
                    st.warning(f"Pandas Traceback: {traceback.format_exc()}", icon="🐛")
        
        # Deduplication logic (remains the same)
        final_controls_list = []
        seen_keys = set()
        for item in controls_list:
            control_name_key = (item.get("Control Name") or "").strip().lower()
            type_key = (item.get("Type") or "").strip().lower()
            sheet_key = (item.get("Sheet") or "").strip().lower()
            if "vba" in type_key and "activex" not in type_key :
                if "commandbutton" in type_key: type_key = "activex (forms.commandbutton.1)" 
            if "xml" in (item.get("Source") or "").lower() and "activex" in type_key:
                try:
                    props = json.loads(item.get("Properties", "{}"))
                    prog_id = props.get("progId", "").lower()
                    if prog_id: type_key = f"activex ({prog_id})"
                except: pass
            item_key = (control_name_key, type_key, sheet_key)
            if item_key not in seen_keys:
                final_controls_list.append(item)
                seen_keys.add(item_key)
                
        if not final_controls_list and ext not in ['.xlsx', '.xlsm', '.xls', '.xlsb']:
            st.warning(f"Unsupported file type: {ext}. Analysis may be limited or incomplete.", icon="⚠️")
        elif not final_controls_list:
            st.info("No controls or identifiable worksheet structures found.")

        return final_controls_list, all_vba_code

    finally: # This 'finally' ensures temp_path is cleaned up if it was created
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception as e_remove_tmp:
                # This is where the WinError 32 might still show if a handle wasn't released above
                st.warning(f"Could not remove temporary file {temp_path}: {e_remove_tmp}", icon="🔥")
                if debug_mode:
                     st.warning(f"Temp File Removal Traceback: {traceback.format_exc()}", icon="🐛")

def format_properties_for_display(df_column): # No changes here
    if pd.api.types.is_string_dtype(df_column):
        return df_column.apply(lambda x: (x[:100] + '...') if isinstance(x, str) and len(x) > 100 else x)
    return df_column

def main(): # No changes here other than passing debug_mode
    st.set_page_config(page_title="Excel Controls & Macros Extractor", layout="wide")
    st.title("🔬 Excel Controls & Macros Extractor")
    
    st.markdown("""
    Upload an Excel file (`.xlsx`, `.xlsm`, `.xls`, `.xlsb`) to extract information about:
    - **Worksheets** (from XML or Pandas scan)
    - **Form Controls & Shapes with Macros** (from XML structure in `.xlsx`/`.xlsm`)
    - **ActiveX Controls** (from XML structure, including `progId` and properties if found)
    - **VBA UserForms & Declared Controls** (from VBA code analysis via `oletools`)
    - **VBA Code References** to sheet controls/objects
    - **VBA Project** presence and full code (in debug mode)
    
    *Note: Extraction from `.xls` and `.xlsb` files primarily focuses on VBA content and basic sheet listing. Detailed XML-based control parsing applies mainly to `.xlsx` and `.xlsm` files.*
    """)
    
    debug_mode = st.sidebar.checkbox("Enable Debug Mode", value=False, help="Show full VBA code and other diagnostic info like XML parsing tracebacks.")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xlsm", "xls", "xlsb"])

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        original_filename = uploaded_file.name
        
        if debug_mode:
            st.sidebar.info(f"**File:** `{original_filename}`\n\n**Size:** `{len(file_bytes)} bytes`")

        with st.spinner(f"Analyzing `{original_filename}`... This might take a moment."):
            start_time = time.time()
            # Pass debug_mode to the extraction function
            controls, all_vba_code_extracted = extract_controls_from_excel(file_bytes, original_filename, debug_mode)
            end_time = time.time()
            st.info(f"Analysis completed in {end_time - start_time:.2f} seconds.")

        # Rest of the main function remains the same
        if controls:
            st.success(f"Found {len(controls)} items (controls, sheets, VBA elements).")
            
            df = pd.DataFrame(controls)
            cols_order = ["Control Name", "Type", "Sheet", "Source", "Properties"]
            df_cols = [col for col in cols_order if col in df.columns] + \
                      [col for col in df.columns if col not in cols_order]
            df = df[df_cols]

            if "Properties" in df.columns:
                 df["Properties_Summary"] = format_properties_for_display(df["Properties"])
                 display_cols = ["Control Name", "Type", "Sheet", "Source", "Properties_Summary"]
            else:
                 display_cols = ["Control Name", "Type", "Sheet", "Source"]

            control_types_for_summary = [
                t for t in df["Type"].unique() if "Worksheet" not in t and "VBA Container" not in t and "potential objects" not in t
            ] if "Type" in df.columns else []

            if control_types_for_summary:
                real_controls_df = df[df["Type"].isin(control_types_for_summary)]
                if not real_controls_df.empty:
                    st.subheader(f"Summary: {len(real_controls_df)} Controls/Code Objects Found")
                    st.dataframe(real_controls_df[display_cols].copy(), use_container_width=True)
            
            st.subheader("All Extracted Items")
            col1, col2 = st.columns(2)
            with col1:
                unique_types = sorted(df["Type"].unique()) if "Type" in df.columns else []
                filter_type = st.multiselect("Filter by Type", options=unique_types, default=[])
            with col2:
                unique_sources = sorted(df["Source"].unique()) if "Source" in df.columns else []
                filter_source = st.multiselect("Filter by Source", options=unique_sources, default=[])

            filtered_df = df.copy()
            if filter_type: filtered_df = filtered_df[filtered_df["Type"].isin(filter_type)]
            if filter_source: filtered_df = filtered_df[filtered_df["Source"].isin(filter_source)]
                
            st.dataframe(filtered_df[display_cols].copy(), use_container_width=True)
            
            csv_export = filtered_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="Download Results as CSV", data=csv_export,
                file_name=f"{os.path.splitext(original_filename)[0]}_extracted_items.csv", mime="text/csv",
            )
            
            if not filtered_df.empty:
                st.subheader("Inspect Item Details")
                filtered_df["_display_name_select"] = filtered_df.apply(
                    lambda r: f"{r.get('Control Name', 'N/A')} ({r.get('Type', 'N/A')}) - Sheet: {r.get('Sheet', 'N/A')} - Src: {r.get('Source', 'N/A')}", axis=1
                )
                
                selected_display_name_unique = st.selectbox(
                    "Select item to view its full properties:",
                    options=filtered_df["_display_name_select"].tolist()
                )
                
                if selected_display_name_unique:
                    selected_row = filtered_df[filtered_df["_display_name_select"] == selected_display_name_unique].iloc[0]
                    
                    st.markdown(f"""
                    - **Name:** `{selected_row.get('Control Name', 'N/A')}`
                    - **Type:** `{selected_row.get('Type', 'N/A')}`
                    - **Sheet:** `{selected_row.get('Sheet', 'N/A')}`
                    - **Source:** `{selected_row.get('Source', 'N/A')}`
                    """)
                    
                    st.write("**Full Properties (JSON):**")
                    try:
                        props_data = json.loads(selected_row["Properties"])
                        st.json(props_data)
                    except (json.JSONDecodeError, TypeError):
                        st.text(selected_row.get("Properties", "No properties or invalid JSON."))
                    
                    selected_control_name_for_vba = selected_row.get("Control Name")
                    if all_vba_code_extracted and selected_control_name_for_vba and \
                       selected_row.get("Type") not in ["Worksheet", "VBA Container", "Worksheet (potential objects)"]:
                        try:
                            escaped_name = re.escape(selected_control_name_for_vba)
                            vba_snippets = re.findall(r'^(?:[^\n]*\b{}\b[^\n]*)$'.format(escaped_name), 
                                                     all_vba_code_extracted, re.IGNORECASE | re.MULTILINE)
                            if vba_snippets:
                                st.subheader(f"VBA Code Snippets mentioning '{selected_control_name_for_vba}'")
                                unique_snippets = sorted(list(set(s.strip() for s in vba_snippets)))
                                for snippet in unique_snippets[:15]:
                                    st.code(snippet, language="vbnet")
                                if len(unique_snippets) > 15:
                                    st.caption(f"Showing 15 of {len(unique_snippets)} unique code snippets.")
                        except Exception as e_snip:
                            if debug_mode: st.warning(f"Error finding/displaying VBA snippets: {e_snip}", icon="🐛")
            else:
                st.info("No items match current filter criteria.")
        
        elif not controls and (ext in ['.xlsx', '.xlsm', '.xls', '.xlsb']):
            st.info("No controls, VBA, or distinct sheet structures were programmatically identified in this file.")
        
        if debug_mode:
            st.sidebar.subheader("Debug: Full VBA Code")
            if all_vba_code_extracted:
                with st.sidebar.expander("View/Hide Full VBA Code", expanded=False):
                    st.code(all_vba_code_extracted, language="vbnet", line_numbers=True)
            else:
                st.sidebar.write("No VBA code was extracted.")
    else:
        st.info("Awaiting Excel file upload. Supported types: `.xlsx`, `.xlsm`, `.xls`, `.xlsb`")

if __name__ == "__main__":
    main()