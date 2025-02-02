import PyInstaller.__main__
import sys
import os

additional_files = []

args = [
    'pdf_splitter.py', 
    '--name=PDF Splitter',  
    '--onefile',            
    '--windowed',           
    '--clean',             
]


for file in additional_files:
    args.append(f'--add-data={file};.')


PyInstaller.__main__.run(args)