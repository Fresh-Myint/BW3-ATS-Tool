# BW3 ATS Tool
An Excel based tool written in VBA and SQL for CEVA Logistics' Dell campus. 
### Description
This tool provides a real time & organized data set of orders along with their attributes for CEVA to execute completion based on the expected ship day and carrier cut-time.

There are three **main catergories** that orders will fall into:
 - WIP (Work in Process
 - Exception
 - Part Shortage
 
The data set from SQL is organized as an array within VBA to reduce run-time and memory usage and then sorted into these categories based of the below attributes:
- Order Holds
- Expeditor comment; the prefix used to denote the comment type
  - PS (Part Shortage)
  - EX (Exception)
- 
