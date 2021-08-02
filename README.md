# ADPExcelSheet
This Excel workbook was designed for automating many repetitive processes when I worked at this job. I taught myself Excel, Access, VBA, and SQL in order to automate my menial tasks. This project represents the beginning of my programming career. All the code is found in 'src.vba'.
# Features
The ADP Excel Sheet can tackle many frequently repeated actions such as formatting fitments, dropship reports, and bill of materials. It can also generate templates for use in uploading to Amazon, Ebay, Finale, Shipstation, GS1, or MyFitment. It features useful exporting functions to save the current worksheet without Macros or as a CSV or TXT file as well as simple email functions.
Buttons are context-aware and will deactivate when their function is not applicable in the current worksheet.
## Fitments
### ![To DB icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/To%20DB.png) To Database button
After formatting, use this button to add the fitments to the AD Find Sets Database.
### ![ToSixbit icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/ToSixbit.png) To Sixbit button
After formatting, use this button to add the fitments to Sixbit to use in Ebay listings.
### ![fixformat icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/fixformat.png) Format Fitments
Paste in fitments from Metro/NexPart or Sixbit. The button automatically detects where the fitments came from and renames the sheet and source appropriately. Fitments must be pasted as values into cell A1. Contains error checking but there's still a good possibility of it missing some things. In that case, manual repairs will be needed but this will still get most of the job done.
### ![MyFitment icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/MyFitment.png) MyFitment Inheritance
This button generates the template needed to upload inheritance information to MyFitment. This template should be exported as XLSX as that is what MyFitment asks for.
## Format
### ![OOS icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/OOS.png) Out of Stock
This button formats the Ebay Out of Stock page. Copy the table from the Out of Stock page in Ebay and paste as values.
### ![ManageInv icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/ManageInv.png) Manage Inventory
This button formats Amazon Manage Inventory page. I use this feature for identifying products which have had their detail page removed (dog picture error message). In order to find all listings, make sure to expand parent listings per page first.
### ![Dropship icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/Dropship.png) Dropship Report
This button will format a Dropship report from Herko or an order report from ShipStation.
### ![ImportShipstation icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/ImportShipstation.png) Import Shipstation
After formatting a Shipstation dropship report, this button will provide the appropriate information from a Shipstation dropship report of your choice to a Dropship Report formatting in the process.
### ![FormatInv icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/FormatInv.png) Inventory
This menu provides a few functions. It will format an "Export Products" report or a "Products sold in date range" report from Finale. Formatting "Export Products" will remove inactive products. Formatting "Products sold in date range" will prepare the sheet for inventory taking. It gives you such actions as removing any products with qty above or below a chosen value as well as removing all products except the rows chosen to keep. There's also a "Confirmed Inventory" button that adds products to a database and removes these products the next time this formatting is run in order to be more efficient when doing inventory.
### ![FormatBoM icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/FormatBoM.png) Bill of Materials
This button formats Bill of Materials reports from Finale.
### ![ReplaceComponent icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/ReplaceComponent.png) Replace Component(WIP)
This button will automatically replace one product for another inside a Bill of Materials.
## Finale Templates
### ![products icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/products.png) Products
This button will provide a menu to choose Finale fields to add to a template to create or update products in Finale.
### ![stocktake icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/stocktake.png) Stock Take
This button generates a template for doing a stock take in Finale.
### ![BoM icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/BoM.png) Bill of Materials
This button generates a template for adding or updating Bill of Materials in Finale.
### ![Lookups icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/Lookups.png) Lookups
This button generates a template for updating Lookups in Finale.
### ![ShippingMethods icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/ShippingMethods.png) Shipping Methods
This button generates a template for filling out a Finale product's shipping methods.
### ![PO icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/PO.png) Purchase Items
This button generates a template for adding products in bulk to a purchase order.
## Listing Templates
### ![Shipstation icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/shipstation.png) Shipstation Fields
This button generates a template for adding or updating products to Shipstation.
### ![amazon icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/amazon.png) List Amazon
The List Amazon button can generate a template for listing to Amazon if it doesn't exist in the workbook yet. Once the template exists, the List Amazon button provides an in-depth menu to provide listing information and will populate the worksheet with the appropriate formatting to upload the worksheet to Amazon.
### ![gs1 icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/gs1.png) GS1
This button generates a template for activating or updating GTINs in GS1 Data Hub.
### ![Volume Pricing icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/Volume%20Pricing.png) Ebay Volume Pricing
This button generates a template for creating volume pricing in Ebay (a feature that Ebay seems to have removed). May be deprecated in the future.
### ![walmart icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/walmart.png) List Walmart
## Export
### ![xlsx icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/xlsx.png) ![csv365](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/csv365.png) ![txt icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/txt.png) File Formats
Clicking on the export buttons will save the current worksheet as XLSX, CSV or TXT format.
### ![email as attachment icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/email%20as%20attachment.png) Email
This split button has the option to Email the current worksheet as an attachment or the option to attach multiple worksheets from the workbook.
## ![Database Status icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/Database%20Status.png) ![Variables icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/Variables.png) ![AD Version icon](https://github.com/Villhelm-E/ADP-Excel-Sheet/blob/master/Icons/AD%20Version.png) About
These buttons provide database connection information and version information.
