# po2pl
Create a packing list from different POs requested by my brother for construction business. 

### Input
Multiple excel files of POs are used to collate into a packing list
* PO number and Project name - needed to identify which PO the items are associated with
* Items purchased - each are placed under the right PO number and PJ name

![input file](https://raw.githubusercontent.com/mjocen/po2pl/master/img/sample_po.png)

### Desired Output
Table format packing list inside a word document ready for printing. *see image below for reference*

![desired output](https://raw.githubusercontent.com/mjocen/po2pl/master/img/output.png)

### Actual Output
Excel file containing the desired format but without the desired font styles. *see image below for reference*

![actual output](https://raw.githubusercontent.com/mjocen/po2pl/master/img/actual_output.png)

### Running the Program
The python file contains everything and can execute the file as is. There are still heaps of issues with this initial release but I'm working on it and listed them as a mental note.

For now, make sure you have the following libraries installed:
    * PyQt5
    * xlrd
    * xlwt
    

Make sure you have the following folder path `Sample_PO/pl` inside the folder where the python file is located.
