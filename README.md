# Marine-Litter-Manager

<center>
<table border ="0">
 <tr>
  <td>
<img src="img/test2.png" width="300"> 
  </td>
  </tr>
 <tr>
  <td>
Marine Litter Manager (MLM) is a Python data formatting tool that can be used to generate:

   
<ul>
  <li>EMODnet beach litter format</li>
  <li>EMODnet seafloor trawlings litter format</li>
</ul>
This is done following the specifications of the official guidelines published by EMODnet Chemistry ( https://www.emodnet-chemistry.eu/ ). It is available for Linux and Windows. 
</td>
  </tr>
 </table>
</center>

# Screenshots
<center>
<table border ="0">
 <tr>
  <td colspan="2">
<img src="img/main.png" width="600">
  </td>
  </tr>
 <tr>
  <td>
<img src="img/10.png" width="300">
  </td>
  <td>
<img src="img/11.png" width="300">
  </td>
  </tr>
  <tr>
  <td>
<img src="img/12.png" width="300">
  </td>
  <td>
<img src="img/13.png" width="300">
  </td>
  </tr>
  <tr>
  <td>
<img src="img/14.png" width="300">
  </td>
  <td>
<img src="img/15.png" width="300">
  </td>
  </tr>
  <tr>
  <td>
<img src="img/16.png" width="300">
  </td>
  <td>
<img src="img/17.png" width="300">
  </td>
  </tr>
  <tr>
  <td>
<img src="img/18.png" width="300">
  </td>
  <td>
<img src="img/19.png" width="300">
  </td>
  </tr>
  <tr>
  <td>
<img src="img/20.png" width="300">
  </td>
  <td>
<img src="img/21.png" width="300">
  </td>
  </tr>
 </table>
</center>

# How To create an exe (Linux/Windows):
<ol>
<li>
https://pypi.org/project/auto-py-to-exe/
<br>
pip install auto-py-to-exe
<br>
auto-py-to-exe
</li>
<li>
add the following files: legenda.txt, logo.png
</li>
<li>
To include the MLM logo use the following option inside auto-py-to-exe:
<br>
--hidden-import='PIL._tkinter_finder'
</li>
<li>
FOR WINDOWS ONLY, with ANACONDA:
<br>
--exclude-module scikit-learn,PyQt5,PyQt4,2to3,IPython,Jinja2,pycparser,scipy
 </li>
</ol>

