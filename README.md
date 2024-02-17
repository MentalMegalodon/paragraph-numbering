# paragraph-numbering
This is a personal project to add paragraph numbers to .docx documents. I use these for beta reading books, so that all readers can reference a specific paragraph. Chapters should reset the numbering if they are titled as chapters. The project currently relies on the `python-docx` package, but hopefully I can rewrite it to parse the underlying xml once I figure out how to query for the right nodes.

## Installation
`python -m pip install -r requirements.txt`

## Use
`python addnumsboth.py <filename>.docx`
This will produce an output file `<filename> with inline numbers.docx` that should have the proper paragraph numbers.
