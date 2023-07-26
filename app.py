import os
import pandas as pd
import streamlit as st

def convert_xlsx_to_csv(input_file, output_file):
    try:
        df = pd.read_excel(input_file, engine='openpyxl')
        df.to_csv(output_file, index=False)
        st.write(f"Conversion successful. CSV file saved as '{output_file}'.")

    except Exception as e:
        st.error(f"An error occurred: {e}")

def main():
    st.title("XLSX to CSV Converter")

    file = st.file_uploader("Upload XLSX file", type=["xlsx"])

    if file is not None:
        if file.name.lower().endswith('.xlsx'):
            # Save the uploaded file temporarily
            temp_path = 'temp.xlsx'
            with open(temp_path, 'wb') as f:
                f.write(file.getvalue())

            # Convert XLSX to CSV
            output_file_path = 'output.csv'
            convert_xlsx_to_csv(temp_path, output_file_path)

            # Remove the temporary XLSX file
            os.remove(temp_path)

            st.success("Conversion successful. CSV file ready for download.")
            st.markdown(get_download_link(output_file_path), unsafe_allow_html=True)
        else:
            st.error("Invalid file format. Only XLSX files are supported.")

def get_download_link(file_path):
    with open(file_path, 'rb') as f:
        file_data = f.read()
    st.download_button(label="Download CSV", data=file_data, file_name="output.csv")

if __name__ == "__main__":
    main()
