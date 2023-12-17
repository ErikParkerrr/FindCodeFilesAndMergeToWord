# About
This is a simple project that searches all sub directories for a specific file extension, grabs the contents, and merges them into a word document. It automatically creates the subheadings based on the file names and sets the font to be code style. 

This is useful for when you need all your code in the appendix section of a report or similar. 

# Usage
Put the main.py file in the highest level folder you want it to search. It will find all subdirectories. In the script, set the file extension type and then run it. It will generate the merged docx file in the same place. 

    FILE_EXTENSION = ".vhd"

This was originally set up to search specifically for VHDL files, but it will work with any plain text files. 