# docx-replace-attach
![License](https://img.shields.io/github/license/mintya/docx-replace-attach)
![GitHub release](https://img.shields.io/github/release/mintya/docx-replace-attach)
![GitHub stars](https://img.shields.io/github/stars/mintya/docx-replace-attach)
![GitHub forks](https://img.shields.io/github/forks/mintya/docx-replace-attach)  

This library was built on top of [python-docx](https://python-docx.readthedocs.io/en/latest/index.html) and the main purpose is to replace key words to attachments inside a document of MS Word.    
  
Just like this:   
![Example](images/img.png)  

## Getting started
You can define a key in your Word document This program requires the following key format: `{key_name}`
### How to install
```sh
pip3 install docx-replace-attach
```
### How to use
Example:
```py
from docx import Document
from docx_attachment import replace_xlsx
from docx_attachment import replace_docx


# The template file
doc = Document('template.docx')

# replace key_name to a word file, key_name is word_attachment, word file is word_attachment.docx
replace_docx(doc, 'word_attachment', 'word_attachment.docx')

# replace key_name to a excel file, key_name is excel_attachment, excel file is excel_attachment.xlsx
replace_xlsx(doc, 'excel_attachment', 'excel_attachment.xlsx')

# save as a new file
doc.save('new.docx')

```
You can also use:
```py
from docx_attachment import replace_xlsx_in_template
from docx_attachment import replace_docx_in_template


replace_docx_in_template('template.docx', 'new.docx', 'word_attachment', 'word_attachment.docx')

replace_xlsx_in_template('template.docx', 'new.docx', 'excel_attachment', 'excel_attachment.xlsx')

```
## Stargazers over time
[![Stargazers over time](https://starchart.cc/mintya/docx-replace-attach.svg?variant=adaptive)](https://starchart.cc/mintya/docx-replace-attach)
