import streamlit as st
from xlsx2ldf import ExcelToLDFConverter
import os
from datetime import datetime

st.set_page_config(
    page_title="Excel to LDF Converter",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        .stApp {
            background-color: #f5f5f5;
        }
        .stFileUploader {
            border: 2px dashed #4CAF50;
            border-radius: 5px;
            padding: 20px;
        }
        .stButton>button {
            background-color: #4CAF50;
            color: white;
            border-radius: 5px;
            padding: 10px 24px;
            font-size: 16px;
            font-weight: bold;
            border: none;
            transition: all 0.3s;
        }
        .stButton>button:hover {
            background-color: #45a049;
            transform: scale(1.05);
        }
        .success-box {
            background-color: #dff0d8;
            color: #3c763d;
            border: 1px solid #d6e9c6;
            border-radius: 4px;
            padding: 15px;
            margin: 20px 0;
        }
        .error-box {
            background-color: #f2dede;
            color: #a94442;
            border: 1px solid #ebccd1;
            border-radius: 4px;
            padding: 15px;
            margin: 20px 0;
        }
        .header {
            color: #4CAF50;
        }
    </style>
""", unsafe_allow_html=True)

def main():

    st.title("🚗 Excel to LDF Converter")
    st.markdown("""
    Convert your LIN network descriptions from Excel format to LDF (LIN Description File) format.
    Upload your Excel file below and download the converted LDF file with a single click!
    """)

    st.header("📤 Upload Your Excel File", divider='green')
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx or .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
        help="The Excel file should contain 'Matrix', 'Info', and 'LIN Schedule' sheets"
    )

    st.header("⚙️ Output Settings", divider='green')
    default_name = f"converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.ldf"
    output_filename = st.text_input(
        "Output LDF filename", 
        value=default_name,
        help="Name for the generated LDF file"
    )
    
    st.header("🔄 Convert", divider='green')
    convert_btn = st.button("Convert Excel to LDF", use_container_width=True)
    
    if convert_btn and uploaded_file is not None:
        with st.spinner('Converting... Please wait'):
            try:
                temp_path = f"temp_{uploaded_file.name}"
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                converter = ExcelToLDFConverter(temp_path)
                success = converter.convert(output_filename)

                os.remove(temp_path)
                
                if success:

                    st.markdown(f"""
                    <div class="success-box">
                        <h3>✅ Conversion Successful!</h3>
                        <p>Your LDF file has been generated successfully.</p>
                    </div>
                    """, unsafe_allow_html=True)

                    with open(output_filename, "rb") as f:
                        st.download_button(
                            label="⬇️ Download LDF File",
                            data=f,
                            file_name=output_filename,
                            mime="application/ldf",
                            use_container_width=True
                        )

                    os.remove(output_filename)
                else:
                    st.markdown("""
                    <div class="error-box">
                        <h3>❌ Conversion Failed</h3>
                        <p>There was an error converting your Excel file to LDF format.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
            except Exception as e:
                st.markdown(f"""
                <div class="error-box">
                    <h3>❌ Error During Conversion</h3>
                    <p>{str(e)}</p>
                </div>
                """, unsafe_allow_html=True)
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                if os.path.exists(output_filename):
                    os.remove(output_filename)
    
    elif convert_btn and uploaded_file is None:
        st.warning("⚠️ Please upload an Excel file first!")

    st.header("📝 Instructions", divider='green')
    with st.expander("How to use this converter"):
        st.markdown("""
        1. **Prepare your Excel file**:  
           - Ensure your Excel file has three sheets named exactly:
             - `Matrix` - Contains signal definitions
             - `Info` - Contains LIN network information
             - `LIN Schedule` - Contains scheduling information
           - Follow the standard format used by the ExcelToLDFConverter
        
        2. **Upload your file**:  
           - Click "Browse files" or drag and drop your Excel file
        
        3. **Set output filename**:  
           - You can keep the default name or enter your own
        
        4. **Convert**:  
           - Click the "Convert Excel to LDF" button
           - Wait for the conversion to complete
        
        5. **Download**:  
           - Once conversion is successful, click the download button to get your LDF file
        """)
    
    st.header("ℹ️ About", divider='green')
    st.markdown("""
    This tool converts LIN network descriptions from Excel format to LDF (LIN Description File) format.
    
    **Features**:
    - Converts signal definitions, frames, and scheduling information
    - Handles both .xlsx and .xls file formats
    - Generates standard-compliant LDF files
    
    **Requirements**:
    - Excel file must follow the expected format
    - Requires Python with ldfparser and pandas installed
    
    Developed with EEA team
    """)

if __name__ == "__main__":
    main()