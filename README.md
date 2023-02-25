# MsWordUpdater
Python script to update a batch of ms word documents (must end in .docx)

step 1:
Update 'data/update_list' to change any text you want to update from old '.docx' files

step 2:
Create venv from requirements.txt

step 3:
Run main.py to update documents according to the update_list.

step 4:
Enjoy the 8 hours you saved not having to do this manually



NOTES:
As of now, anytime the text of a phrase gets updated, that text's formatting (Bold, Font, FontSize, etc.) gets over-written back to the documents default. Will add the ability to retain formatting in the future.

Issues with Table of Contents of a document getting deleted if a section title gets updated. Will need ot investigate this...

footer_loop, header_loop, and table_loop are very similar. Will look into consolidating into 1 single method
