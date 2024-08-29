# SMS-Fisher
SMS Fisher is a set of tools for automating certain tedious tasks at SMS Distribution. </br></br>
<b>billreader.py</b> - A script to reformat a PDF file with a Santini bill to a .xlsx file that can be imported directly into Fishbowl.</br></br>
<b>catalogreader.py</b> - A script to reformat a .xlsx file with a Santini collection catalog into an .xlsx file that can be imported directly into Fishbowl.</br></br>
<b>catalogtowholesale.py</b> - A script to convert catalogs made with catalogreader.py into wholesale orders that can be imported directly into Fishbowl.</br></br>
<b>wholesaletosalesorder.py</b> - A script to conver wholesale orders made with catalogtowgolesale.py to sales order that can be directly imported into Fishbowl.</br></br>
<b>l1tousd</b> - A script to convert L1 spreadsheets from EUR into USD at a fixed rate.</br></br>
<u>Installation steps:</u>
<ol>
    <li>Make sure you have Python installed</li>
    <li>Clone repository</li>
    <li>Open Terminal in folder where requirements.txt is located</li>
    <li>Run this command:
    <br><code>py -m pip install -r requirements.txt</code></li>
    <li>Run <code>BillReader.py</code> or <code>CatalogReader.py</code> and follow instructions given. Make sure the .xlsx file you're using is formatted correctly.</li>
</ol>
<br><br>
<b><i>This program was made for use by SMS Distribution employees. It contains no sensitive information, so you may use it or clone it whoever you are, it just likely won't be useful.</i></b>

