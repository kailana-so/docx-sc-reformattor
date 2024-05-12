# Superscript reformatter for docx documents

A weird little script for reformatting superscripts to in-text references.

### How to use

- docx document to reformat
- CSV file with intext references

Paste the path to your CSV file against the references_data variable:

```
references_data = "./mock_data/references.csv"
```

Paste the path to the docx document against the input_path variable:

```
input_path = "./mock_data/doc_to_fix.docx"
```

Write the path to the new file against the output_path variable:
Note: this will create the new file too.

````
output_path = './newtext/replaced_superscripts_text.docx'
````

### To run

Call the script in your terminal. 
Note: if your not in the folder the script is located, also inc. path to your script before running.

````
python3 superscript_reformatter.py
````

### Requirements

- python 
- docx

If you don't have either:

````
brew install python
````

````
pip install docx
````