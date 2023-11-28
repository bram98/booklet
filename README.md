# Booklet
Files for processing the output from the Debye Booklet forms

# How to use
1. Install anaconda, preferably create a new environment
2. Run the commands in commands.txt to install the right packages (TODO currently there are a bit too many)
3. Unzip the responses from the OneDrive download file
4. Rename the folders until you have the following data structure together with your python files. Note: the response_data and subfolders should be literally the same
   '''
   ├── responses.xslx
   └── response_data
       ├── figures
           └── ... all figures ...
       ├── portraits
           └── ... all photos ... 
       └── project_descriptions
           └── ... all project descriptions and references ...
   '''
  5. Run `rename_files.py`. This should create some more folders containing nicely named files. Be sure to run `inspect_names()` to see any differences between the names provided in the form and that of uu.
