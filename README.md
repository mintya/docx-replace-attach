# docx-replace-attach
This library was built on top of [python-docx](https://python-docx.readthedocs.io/en/latest/index.html) and the main purpose is to replace key words to attachments inside a document of MS Word.  
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
from docx_attachment import replace_word


# The template file
doc = Document('template.docx')

# replace key_name to a word file, key_name is word_attachment, word file is word_attachment.docx
replace_word(doc, 'word_attachment', 'word_attachment.docx')

# replace key_name to a excel file, key_name is excel_attachment, excel file is excel_attachment.xlsx
replace_xlsx(doc, 'excel_attachment', 'excel_attachment.xlsx')

# save as a new file
doc.save('new.docx')

```
You can also use:
```py
from docx_attachment import replace_xlsx_t
from docx_attachment import replace_word_t


replace_word_t('template.docx', 'new.docx', 'word_attachment', 'word_attachment.docx')

replace_xlsx_t('template.docx', 'new.docx', 'excel_attachment', 'excel_attachment.xlsx')

```
