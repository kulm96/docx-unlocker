# docx-unlocker
Simple Python script to unlock protected Word documents

# Usage
python docx-unlocker.py protected_word_doc.docx

This script unlocks .docx files by modifying the 'w:enforcement' attribute.
If a directory is provided, the script will unlock all .docx files within that directory.
Unlocked files will be created alongside the original with "_UNLOCKED" appended.


