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
When matches found, it will display to the textboxes
![Home](/interface/home.JPG)

2. Update Data - 
    a. The user will type the part number frirst then click the RETRIEVE button.
    b. The data will be displayed.
    c. The user will change the data
    d. Click the UPDATE button and the system will update the data,


![Update](/interface/errorcorrection.JPG)

