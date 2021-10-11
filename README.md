# BW3 ATS Tool
An Excel based tool written in VBA and SQL for CEVA Logistics' Dell campus. 
### Description
This tool provides a real time & organized data set of orders along with their attributes for CEVA to execute completion based on the expected ship day and carrier cut-time.

There are three **main catergories** that orders will fall into:
 - WIP (Work in Process)
 - Exception
 - Part Shortage
 
The data set from SQL is organized as an array within VBA to reduce run-time and memory usage and then sorted into these categories based of the below attributes:
- Order Holds
- Expeditor comment; the prefix used to denote the comment type
  - PS (Part Shortage)
  - EX (Exception)
- If the above arguments are not met then the order is listed as WIP and is able to ship.

Orders will then be sorted into **"buckets"** based on order status, ESD, RSD, download date and time, and carrier cut-time. 

### Files Listed
The files listed in this repository are as follows:
- Example Data Folder - A small screen shot of the main pivot data and an example data set from the tool is stored here.
- MAIN.bas - This is the main VBA module that is imported into the workbook, this module manipulates the data.
- MAIN.sql - The SQL query that pulls that data form our WMS (warehouse managemnet system)
- VERSION_CONTROLLER.bas - This module is what pulls in the MAIN.bas module. This helps with versionm controling the module after distrobution. 
