This component converts excel file to json and json to excel.
For this conversion, we use third party JavaScript (Sheet JavaScript) in lwc. By using this javascript method we can easily read excel and convert it. 

we can use the ```Sheet js``` for conversation and ```codemirror js``` for show the converted json in code block. we are using the ```apex controller (ExcelController)``` for fetching the ```Accounts``` and ```Contacts``` Records. we do not download the formatted excel in lwc so we using the ```visual force page (jsonToExcelParserPage)```.

 <h2>Parsing functions</h2>

* ```XLSX.read(data, read_opts)``` attempts to parse data.
* ```XLSX.readFile(filename, read_opts)``` attempts to read filename and parse.

<h2>Writing functions</h2>

* ```XLSX.write(wb, write_opts)``` attempts to write the workbook wb
* ```XLSX.writeFile(wb, filename, write_opts)``` attempts to write wb to filename. In browser-based environments, it will attempt to force a client-side download.
* ```XLSX.writeFileAsync(wb, filename, o, cb)``` attempts to write wb to filename. If o is omitted, the writer will use the third argument as the callback.

<h2>Utilities</h2>

Utilities are available in the XLSX.utils object and are described in the Utility Functions section:

**Importing**:

* ```aoa_to_sheet``` converts an array of arrays of JS data to a worksheet.
* ```json_to_sheet``` converts an array of JS objects to a worksheet.
* ```table_to_sheet``` converts a DOM TABLE element to a worksheet.
* ```sheet_add_aoa``` adds an array of arrays of JS data to an existing worksheet.
* ```sheet_add_json``` adds an array of JS objects to an existing worksheet.
* ```sheet_to_row_object_array```  return a array of the row objects
* ```book_new()```  for creating new excel book
* ```book_append_sheet()``` Add worksheet to Excel

**Exporting**:

* ```sheet_to_json``` converts a worksheet object to an array of JSON objects.
* ```sheet_to_csv``` generates delimiter-separated-values output.
* ```sheet_to_txt``` generates UTF16 formatted text.
* ```sheet_to_html``` generates HTML output.
* ```sheet_to_formulae``` generates a list of the formulae (with value fallbacks).

**Cell and cell address manipulation:**

* ```format_cell``` generates the text value for a cell (using number formats).
* ```encode_row / decode_row``` converts between 0-indexed rows and 1-indexed rows.
* ```encode_col / decode_col``` converts between 0-indexed columns and column names.
* ```encode_cell / decode_cell``` converts cell addresses.
* ```encode_range / decode_range``` converts cell ranges.

<h2>Cell Object</h2>

Cell objects are plain JS objects with keys and values following the convention:

| Key   |  Description                                                           |
|:------|:-----------------------------------------------------------------------|
| `v`   | raw value (see Data Types section for more info)                       |
| `w`   | formatted text (if applicable)                                         |
| `t`   | type: `b` Boolean, `e` Error, `n` Number, `d` Date, `s` Text, `z` Stub |
| `f`   | cell formula encoded as an A1-style string (if applicable)             |
| `F`   | range of enclosing array if formula is array formula (if applicable)   |
| `r`   | rich text encoding (if applicable)                                     |
| `h`   | HTML rendering of the rich text (if applicable)                        |
| `c`   | comments associated with the cell                                      |
| `z`   | number format string associated with the cell (if requested)           |
| `l`   | cell hyperlink object (`.Target` holds link, `.Tooltip` is tooltip)    |
| `s`   | the style/theme of the cell (if applicable)                            |

<h2>Data Types</h2>

The raw value is stored in the v value property, interpreted based on the t type property. This separation allows for representation of numbers as well as numeric text. There are 6 valid cell types:

| Type  |  Description                                                           |
|:------|:-----------------------------------------------------------------------|
| `b`   | Boolean: value interpreted as JS boolean                               |
| `e`   | Error: value is a numeric code and `w` property stores common name **  |
| `n`   | Number: value is a JS `number` **                                      |
| `d`   | Date: value is a JS `Date` object or string to be parsed as Date **    |
| `s`   |cText: value interpreted as JS `string` and written as text **          |
| `z`   | Stub: blank stub cell that is ignored by data processing utilities **  |

<h2>Workbook Object</h2>

* ```workbook.SheetNames``` is an ordered list of the sheets in the workbook
* ```wb.Sheets[sheetname]``` returns an object representing the worksheet.
* ```wb.Props``` is an object storing the standard properties. 
* ```wb.Custprops``` stores custom properties. Since the XLS standard properties deviate from the XLSX standard, XLS parsing stores core properties in both places.
* ```wb.Workbook``` stores workbook-level attributes.
* ```wb.Workbook.Names``` is an array of defined name objects which have the keys:

<h2>Cell Styles</h2>

Cell styles are specified by a style object that roughly parallels the OpenXML structure. The style object has five top-level attributes: fill, font, numFmt, alignment, and border.

| Style Attribute  |  Sub Attributes | Values 																					   |
|:-----------------|:----------------|:--------------------------------------------------------------------------------------------|
| fill             | patternType     | `solid` or `none` 																		   |
|                  | fgColor         | COLOR_SPEC 																				   |
|                  | bgColor         | COLOR_SPEC 																				   |
| font             | name            | "Calibri" // default 																	   |
|                  | sz				 | "11" // font size in points  															   |
|                  | color			 | COLOR_SPEC 																				   |
|                  | bold			 | true or false 																			   |
|                  | underline		 | true or false 																			   |
|   			   | italic			 | true or false 																			   |
|   			   | strike			 | true or false 																			   |
|   			   | outline		 | true or false 																			   |
|   			   | shadow			 | true or false 																			   |
|   			   | vertAlign		 | true or false 																			   |
| numFmt		   | 				 | "0" // integer index to built in formats, see StyleBuilder.SSF property 					   |
| 				   | 				 | "0.00%" // string matching a built-in format, see StyleBuilder.SSF 						   |
| 				   | 				 | "0.0%" // string specifying a custom format 												   |
| 				   | 				 | "0.00%;\\(0.00%\\);\\-;@" // string specifying a custom format, escaping special characters |
| 				   | 				 | "m/dd/yy" // string a date format using Excel's format notation 							   |
| alignment 	   | vertical	     | "bottom" or "center" or "top" 															   |
| 				   | horizontal	     | "left" or "center" or "right" 															   |
| 				   | wrapText	     | true or false 																			   |
| 				   | readingOrder	 | 2 // for right-to-left  															   		   |
| 				   | textRotation    | Number from 0 to 180 or 255 (default is 0)  												   |
|  				   | 				 | 90 is rotated up 90 degrees 																   |
|  				   | 				 | 45 is rotated up 45 degrees 																   |
|  				   | 				 | 135 is rotated down 45 degrees 															   |
|  				   | 				 | 180 is rotated down 180 degrees 															   |
|  				   | 				 | 255 is special, aligned vertically 														   |
| border 		   | top	         | { style: BORDER_STYLE, color: COLOR_SPEC } 												   |
| 				   | bottom			 | { style: BORDER_STYLE, color: COLOR_SPEC } 												   |
| 				   | left			 | { style: BORDER_STYLE, color: COLOR_SPEC } 												   |
| 				   | right			 | { style: BORDER_STYLE, color: COLOR_SPEC } 												   |
| 				   | diagonal		 | { style: BORDER_STYLE, color: COLOR_SPEC } 												   |
| 				   | diagonalUp		 | true or false 																			   |
| 				   | diagonalDown	 | true or false 																			   |
