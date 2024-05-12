import csv
import re
from docx import Document

print()
print("* Start *")
references_dict = {}

# Path from this file to the reference cvs - if not in cvs format just convert it
references_data = "./dummydata/references.csv"

#  Build a dictionary with the csv file
print("------------")
print("Reading references: ", references_data)
with open(references_data, 'r') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        key = row[0].strip()
        ref = row[1].strip()
        references_dict[key] = ref

#  Optional: write the dictionary to a new file
# new_dict = "./newtext/referencesShortDic.txt"
# print("------------")
# print("Writing dictionary: ", new_dict)
# with open(new_dict, 'w') as convert_file:
#     convert_file.write(json.dumps(references_dict))

# Path from this file to the Document you want to convert 
input_path = "./dummydata/NicolasWorkText.docx"
document = Document(input_path)
print("------------")
print("Reading text: ", input_path)

print("------------")
print("Starting replacements...")

#  replacing the superscript in document 
def replace_superscript_texts(paragraph):
    for run in paragraph.runs:
        if run.font.superscript:
            new_text = []
            # Extract the superscript text and split by commas to handle multiple references
            # also trim white space
            cleaned_text = run.text.replace(" ", "")
            cleaned = cleaned_text.replace('–', '-')
            intext_superscripts = cleaned.split(',')
            # Optional: for debugging
            # print(supers, "supers")

            # Handle each superscript and an superscript ranges ie. 29-31
            for intext_superscript in intext_superscripts:
                if intext_superscript != '':
                    # Optional: for debugging
                    # print(intext_superscript, "intext_superscript")

                    # loop through ranges
                    if "–" in intext_superscript or "-" in intext_superscript:
                        # Replace endash with normal dash for consistency when processing
                        intext_superscript = intext_superscript.replace("–", "-")
                        intext_superscript_ranges = intext_superscript.split("-")
                        for sc_range in intext_superscript_ranges:
                            if sc_range == "":
                                intext_superscript_ranges.remove(sc_range)
                            if sc_range == "-":
                                intext_superscript_ranges.remove(sc_range)
                        if len(intext_superscript_ranges) == 1:
                            # Handle individual numbers
                            # removing pre and suffix dashes
                            cleaned = intext_superscript.replace(intext_superscript[intext_superscript.index("-")],"",1)
                            ref = references_dict.get(cleaned.strip(), f"[ERROR not found: {cleaned}]")
                            new_text.append(ref)
                        else:
                            if intext_superscript_ranges != "+" and intext_superscript_ranges != "th":
                                start, end = map(int, intext_superscript_ranges)
                                # Generate the range of numbers and look up each one in the dictionary
                                expanded_range = range(start, end + 1)
                                for num in expanded_range:
                                    ref = references_dict.get(str(num), f"[ERROR not found: {num}]")
                                    new_text.append(ref)
                    else:
                        if intext_superscript != "+" and intext_superscript != "th" and intext_superscript != "-":
                            # Handle individual numbers
                            ref = references_dict.get(intext_superscript.strip(), f"[ERROR not found: {intext_superscript}]")
                            new_text.append(ref)
                    run.font.bold = False
                    run.font.superscript = False

                    if "ERROR" in "".join(new_text):
                        print("Reference Error: ", new_text)
            # Replace the original run text with the new references text
            run.font.bold = False
            run.font.superscript = False
            run.text = " ".join(new_text)

# enumerating over the paragraphs
for paragraph in document.paragraphs:
    replace_superscript_texts(paragraph)
print("------------")
print("Replacements completed")

output_path = './newtext/replaced_superscripts_text.docx'
document.save(output_path)
print("------------")
print(f"File successfully written to: {output_path}")
print("------------")
print("* End *")
print()