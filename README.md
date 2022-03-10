# Marine-Litter-Manager

To create an exe (Linux/Windows):

https://pypi.org/project/auto-py-to-exe/
pip install auto-py-to-exe
auto-py-to-exe


To include the NODC logo use the following option inside auto-py-to-exe:
--hidden-import='PIL._tkinter_finder'


FOR WINDOWS ONLY with ANACONDA:
--exclude-module scikit-learn,PyQt5,PyQt4,2to3,IPython,Jinja2,pycparser,scipy
