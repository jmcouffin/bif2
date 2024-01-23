import string
import os.path as op
import olefile
import re
import os
import streamlit as st

def get_rvt_file_version(rvt_file):
    if op.exists(rvt_file): 
        if olefile.isOleFile(rvt_file): 
            rvt_ole = olefile.OleFileIO(rvt_file) 
            bfi = rvt_ole.openstream("BasicFileInfo")
            file_info_bytes = bfi.read()  # read the stream once
            bfi_text = str(file_info_bytes.decode("ascii", "ignore"))
            # get the Central Model Path, Locale when saved Format, and the Build Number
            bfi_text = ''.join(filter(lambda x: x in string.printable, bfi_text))
            # print(bfi_text.splitlines())
            cmp = re.search(r"Central Model Path: (.*)", bfi_text)
            print(cmp.group(1))
            if ".rvt" in cmp.group(1):
                cmp_data = "Central Model Path: " + cmp.group(1)
            else:
                cmp = re.search(r"Last Save Path: (.*)", bfi_text)

                if ".rvt" in cmp.group(1):
                    cmp_data = "Central Model Path: " + cmp.group(1).split("\\")[-1]
                else:
                    cmp_data = "Central Model Path: None"
            locale = re.search(r"Locale when saved: (.*)", bfi_text)
            if locale:
                locale_data = "Locale when saved: " + locale.group(1)
            else:
                locale_data = "Locale when saved: None"
            build = re.search(r"Build: (.*)", bfi_text)
            if build:
                build_data = "Build Number: " + build.group(1)
            else:
                build_data = "Build Number: None"
            rvt_format = re.search(r"Format: (.*)", bfi_text)
            if rvt_format:
                rvt_format_data = "Format: " + rvt_format.group(1)
            else:
                rvt_format_data = "Format: None"
            return cmp_data + "\n" + locale_data + "\n" + build_data + "\n" + rvt_format_data

        else: 
            pass 
    else: 
        pass

if __name__ == "__main__":
    folder = st.text_input("Enter folder path to search for rvt files")
    # walk the folder and find rvt files
    st.write(folder)
    for root, dirs, files in os.walk(folder):
        for files in files:
            st.write(files)
            if files.endswith(".rvt"):
                st.write(files)
                rvt_file = os.path.join(folder, files)
                # test if file exists
                if op.exists(rvt_file):
                     st.write(files)
                     st.write(get_rvt_file_version(rvt_file))
