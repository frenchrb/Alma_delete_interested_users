# Alma_delete_interested_users

Script to delete interested users from Alma PO Lines


## Requirements
Created and tested with Python 3.6; see ```environment.yml``` for complete requirements.

Requires an Alma Acquisition API key with Read/write permissions. A config file (local_settings.ini) with this key should be located in the same directory as the script and input file. Example of local_settings.ini:

```
[Alma Acq R/W]
key:apikey
```


## Usage
```python delete_interested_users.py input.xlsx```
where ```input.xlsx``` is a spreadsheet listing POL IDs in the first column.


## Contact
Rebecca B. French - <https://github.com/frenchrb>
