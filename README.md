# pdf_to_excel
Convert an existing PDF file to Excel for saving account details

## Install
Installing necessary repositories:
```
pip3 install -r requirements.txt
```

## Usage
Open a new tab in the terminal and go into this directory and run the following command:
```
python3 pdf2text.py
```

Open a new tab in your browser and go to this URL:
```
http://127.0.0.1:5000
```

Choose file, and upload a PDF that you need to convert.
Then download it once it has been converted.


## FAQ:
1. In case the browser throws an "Access Denied to 127.0.0.1", close the tab, re-run the python code and try again. 
    (For google chrome) If it still doesn't work, go to `chrome://net-internals/#sockets` -> `[Flush socket pools]`
