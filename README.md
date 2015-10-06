# Excel-Press

## Summary ##
Excel-Press is a Python implementation of Microsoft Office's proprietary VBA compression and decompression algorithm.

Excel files with an '.xls' file extension (Excel 97-2003) are essentially zip files. There are already tools widely available to unzip and decompress the contents of these VBA macros, however; I was unable to find a Python implementation of Microsoft's VBA compression algorithm.

## Files ##
* [excel-press.py](https://github.com/coldfusion39/excel-press/blob/master/decompress.py): Python script to compress or decompress a VBA macro file
* [macro_raw](https://github.com/coldfusion39/excel-press/blob/master/examples/macro_raw): VBA macro in compressed format
* [macro_readable](https://github.com/coldfusion39/excel-press/blob/master/examples/macro_readable): VBA macro after decompression
* [test.xls](https://github.com/coldfusion39/excel-press/blob/master/examples/Excel/test.xls): Excel document containing VBA macro
* [test](https://github.com/coldfusion39/excel-press/tree/master/examples/Excel/test): Contents of [test.xls](https://github.com/coldfusion39/excel-press/blob/master/examples/Excel/test.xls) after the .xls extension is changed to .zip and unzipped
* [Module1](https://github.com/coldfusion39/excel-press/blob/master/examples/Excel/Module1): Full VBA macro file after unzipping [test.xls](https://github.com/coldfusion39/excel-press/blob/master/examples/Excel/test.xls)

## Requirements ##
The [olefile](https://bitbucket.org/decalage/olefileio_pl) Python library is required.

`pip install olefile`

## Examples ##
Decompress an already compressed VBA macro file

`python excel-press.py -d ./examples/Module1`

Output just the VBA macro portion of the compressed VBA file 

`python excel-press.py -d ./examples/Module1 --raw`

Compress the specified VBA macro

`python excel-press.py -c ./examples/macro_readable`

## Credits ##
The decompress function of excel-press.py is largely adapted from Didier Stevens' [oledump.py](http://blog.didierstevens.com/programs/oledump-py/) script.

Information about Microsoft Office's proprietary VBA compression and decompress scheme was found in the following MSDN documentation [MS-OVBA.pdf](https://msdn.microsoft.com/en-us/library/office/cc313094%28v=office.12%29.aspx).
