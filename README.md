# cashreport-macro
LibreOffice/OpenOffice macro for Hungarian cash report (időszaki pénztárjelentés)

This macro is written for use in LibreOffice or OpenOffice. Given a properly formatted .ods file (having the rigth columns with validated data), upon running it prepares a cash book report table which can then be saved as a pdf with pre-set format and document numbering. After printing it can be used as a document that tax authorities will accept.

The macro puts data into chronological order, sums incomes and expenses, formats values according to Hungarian standards, ands an identifier based on the period of the data and default settings. It also formats the pdf page ready for printing to have all eleents which are requred by the tax office.

The macro is installed in the attached penztarjelentes.ods (OpenOffice/LibreOffice Calc) file. It will install 3 menu options into the *Tools* menu named as:
  * *Pénztárjelentés - adatrendezés*: for sorting data in columns A...E
  * *Pénztárjelentés - készítése*: for making the report into columns H-M
  * *Pénztárjelentés - mentése pdf-be*: to format and save the report with a calculated serial ID as pdf

To use the macro for your needs update the Defaults datastructure elements in the GetDefault function. 

It has been tested with LibreOffice 5.4.3 and 5.4.6 on Windows 10, and with OpenOffice 1.4.3 on Windows 10. Should run however on newer versions and other platforms too. 

Disclaimer. Although the format is according to the official requirements, this is not an official document approved or audited by the Hungarian tax authorities, so use it at your own risk.
