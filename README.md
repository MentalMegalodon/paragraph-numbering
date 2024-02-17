# paragraph-numbering
This is a personal project to add paragraph numbers to .docx documents. I use these for beta reading books, so that all readers can reference a specific paragraph. Chapters should reset the numbering if they are titled as chapters. The project currently relies on the `python-docx` package, but hopefully I can rewrite it to parse the underlying xml once I figure out how to query for the right nodes.

## Installation
Clone the git repository.

`python -m pip install -r requirements.txt`

## Use
`python addnumsboth.py <filename>.docx`

This will produce an output file `<filename> with inline numbers.docx` that should have the proper paragraph numbers.

## Goals
- Rewrite code to use the built-in python xml parser and insert the paragraph numbers without any other modification of the underlying text.
- Re-add code to insert both inline and margin numbers, to be more friendly to kindle formatting.
- Containerize the project so it can be run more easily on any machine.
