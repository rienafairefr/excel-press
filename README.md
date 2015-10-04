# Excel-Compress

## Summary
Excel-Compress is a Python implementation of Microsoft Office's proprietary VBA compression file format.
Excel files with an '.xls' file extension are essentially zip files containing, among other things, VBA macros.
There are already tools widely available to unzip and decompress the contents of these VBA macros, however; I was unable to find a Python implementation of Microsoft's proprietary VBA compression algorithm.

## Files
* [decompress.py](https://github.com/coldfusion39/Excel-Compress/blob/master/decompress.py)
  * Python script to decompress an already compressed VBA macro file.
* [compress.py](https://github.com/coldfusion39/Excel-Compress/blob/master/compress.py)
  * Python script to compress a plain text VBA macro file.

* [examples](https://github.com/coldfusion39/Excel-Compress/tree/master/examples)
  * [macro_raw](https://github.com/coldfusion39/Excel-Compress/blob/master/examples/macro_raw)
    * VBA macro in compressed format
  * [macro_readable](https://github.com/coldfusion39/Excel-Compress/blob/master/examples/macro_readable)
    * VBA macro after decompression
  * [test.xls](https://github.com/coldfusion39/Excel-Compress/blob/master/examples/Excel/test.xls)
    * Excel document containing VBA macro
  * [test](https://github.com/coldfusion39/Excel-Compress/tree/master/examples/Excel/test)
    * Contents of [test.xls](https://github.com/coldfusion39/Excel-Compress/blob/master/examples/Excel/test.xls) after the .xls extension is changed to .zip and unzipped
  * [Module1](https://github.com/coldfusion39/Excel-Compress/blob/master/examples/Excel/Module1)
    * Full VBA macro file after unzipping [test.xls](https://github.com/coldfusion39/Excel-Compress/blob/master/examples/Excel/test.xls)

## Dependencies
[olefile](https://bitbucket.org/decalage/olefileio_pl) is required to run both decompress.py and compress.py.

`pip install olefile`

## Examples
Decompress an already compressed VBA macro file

`python decompress.py -f ./examples/Module1`

Output just the VBA macro portion of the compressed VBA file 

`python decompress.py -f ./examples/Module1 --raw`

Compress the specified VBA macro

`python compress.py -f ./examples/macro_readable`

## Credits
[decompress.py](https://github.com/coldfusion39/Excel-Compress/blob/master/decompress.py) is largely adapted from Didier Stevens' [oledump.py](http://blog.didierstevens.com/programs/oledump-py/) script.

Information about Microsoft Office's proprietary VBA compression and decompress scheme was found in the following MSDN documentation [MS-OVBA.pdf](https://msdn.microsoft.com/en-us/library/office/cc313094%28v=office.12%29.aspx).
