# Drugs Parser for .docx files


## Drugs Parser for .docx files
This  code parse file with requested drug names, sort them by frequency, and group drugs that are written with mistakes (Aspirin or Asperin)

## Before you start
- pip install argparse
- pip install python-docx

## How to use
- python parse.py --source_path ./ліки.docx --target_path ./res.docx


 The output consists of 2 files: res.docx - that contain detailed information about all requests, and res_short.docx - contains only names, and number of requests 
 
 In the case the drug title is divided with '/' (гидазепам/гидозепам/гідазепам), it means that those drugs were grouped together (because of misspelling or using different languages) 
## What if I dont want to use .docx file?
 In this case you should change read_doc function, so that it reads lines from your type of file. (it is really simple task) 
 
