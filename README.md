# Client/Custodian Corporate Action Cross Reference
This repository showcases a project I developed working in asset management. The master workbook (NT VLOOKUP) cross references election data from all internal portfolio managers with our custodian input data in order to save time, mitigate risk and exposure, and to ensure accuracy with our largest custodian.

In order to mitigate risk, I created this workbook for all corporate actions partners to use. The custodian that we work for happens to hold the large majority of our accounts. Therefore using the technology that we have, I developed this workbook to not only save time, but also to implement automation where the risk of user error is high. You first import the saved election data from the portfolio managers (COMPANY _ AGTI INSTX), then you import the corresponding custodian elected input data (COMPANY _ CDR WEB). Once both files are loaded, the program will finish by cleaning and formatting the data, followed by an implementation of multiple VLOOKUP functions to cross reference both sheets. You will be notified when both sheets are loaded and complete. 

Note: For confidentiality purposes, I have completely changed the data in both: the election data from the portfolio managers, and the corresponding custodian elected input data. In a real-world application at work, we are using actual inputs with sensitive client information.

Enjoy! ^_^
