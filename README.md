The PST extractor supported Python version 3.5 to 3.10. Here in this project we used 3.10.0
After you checkout the code, please make sure that you create a virtual environment for python version 3.10.0 using the command python -m venv /path/to/new/virtual/environment/venv
Next is to select the required python interpretor. For that, click Cntrl + shift + p --> Select Interpretor > Enter Interpretor path > Find > Scripts Folder(visible only if you create a virtual environment) > select python application
All the required code resides in server.py file
Now you need to install all the required pytho dependencies used in the code Run the code
pip install flask
pip install -U flask-cors
pip install Aspose.Email-for-Python-via-NET
pip install BeautifulSoup4
After you run the code, the server will start in port 5000
