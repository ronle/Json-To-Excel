# JsonToExcel
This Excel powered by VBA is capable of taking Json file, present the user with a treeview structure of the file and allow user
selection of the key elements upon which a new line would be created.

In the repository is a simplified Json file to allow for simple conversion and representation in tabular view.  

To operate:  
1. Open the table and click on "Button 1"  
2. Acknowledge the welcome message  
3. Click on "Build TreeView"  
4. The script will create a treeview representing the Json file  
5. Expand the tree and select "colorName" as the key value (place checkmark next to it)  
6. Select the output tab & click the "Parse To Excel" button  
  
* Blue values represents end KeyNames that are valid entries for a new line break.
* None blue lines will are master keys that a new line cannot be broken by 
  
Note: 
Not all files are suitable for representation as tabular view.  
As some Objects might include different items, their representation in tabular view might not be consistent which makes it  
difficult to spot patterns.
