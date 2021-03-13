# Introduction
This is an inventory system that uses MS Excel VBA as the main program.
The system is used to track the stocks in the warehouse.

# Design

## Database
There are three tables for this system:
1. Items (Component Part Number, Description, System
Quantity, Stock Unit, Current 
Location, Picture, Vendor , cost)
2. Vendor(supplier name, items sold)
3. Transaction History (Part Number,	Description,	Reference Number,	Transaction Description,	Date of Transaction,	Transactor,	Old Stock Quantity,	New Stock Quantity,	Old Vendor,	New Vendor,	Old Location,	New Location,	Old Cost,	New Cost)
4. Rack (Location, Description)

## Interface
1. Search - 
For the search method, the user will type the part number, and the system will automatically search for the part number.
When matches found, it will display to the textboxes.

![Home](/interface/home.JPG)

2. Update Data - 
    1. The user will type the part number first then click the RETRIEVE button.
    2. The data will be displayed.
    3. The user will change the data
    4. Click the UPDATE button and the system will update the data,


![Update](/interface/errorcorrection.JPG)

3. Add Item - 
    1. Click the Add Item tab and click ADD TO DATABASE button.
    2. Type the part number and description
    3. Click the ADD CATEGORY/VENDOR to add vendor
    4. Upload picture
    5. Select the current location
   
![Add](/interface/additem.JPG)

![Add Supplier](/interface/addsupplier.JPG)

4. Create Purchase Order - 
    1. In the update database tab, select Purchase Order in the dropdown menu
    2. Click GO, and the interface below will be shown
    3. Type and retrieve the necessary information
    4. Then, click GO to add the new PO
    5. This will then be added to the Transaction History database.

![Create Purchase](/interface/po.JPG)

5. Withdraw Item - 
    1. In the update database tab, select Withdraw in the dropdown menu
    2. Click GO, and the interface below will be shown
    3. Type and retrieve the necessary information
    4. Then, click GO to add the withdrawed item
    5. This will then be added to the Transaction History database.

![Withdraw](/interface/withdraw.JPG)
