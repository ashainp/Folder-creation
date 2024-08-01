# Folder-creation
Folder Creator from Word Document

Overview
The Folder Creator from Word Document script is a Python tool designed to automate the creation of folders based on a list of names provided in a Microsoft Word document (.docx). It streamlines the process of setting up multiple directories, which can be useful for organizing projects, data, or any other resources. The script reads each line of the document, sanitizes the folder names, and creates corresponding directories in a specified base directory.

Features
Automated Folder Creation: Quickly generate folders from a list of names in a Word document.
Error Handling: Includes robust error handling to manage issues with file access or folder creation.
Customizable: Users can specify the document path and base directory for flexibility.
Logging: Detailed logging of actions taken by the script, including folder creation success and error messages.

Requirements
Python 3.x
python-docx library

Example
If your Word document (Folders.docx) contains:
Project Alpha
Project Beta
Project Gamma

The script will create the following folders in the base directory:
1-Project Alpha
2-Project Beta
3-Project Gamma

Pretty simple code, but if you have hundreds of folders that need to be created, this can save you some time.


Contributions are welcome! If you have suggestions for improvements or encounter any issues, please feel free to fork the repository and submit a pull request. Ensure that your contributions are well-documented and tested.
