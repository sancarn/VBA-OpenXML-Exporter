# cRawXLData
Excel workbooks these days (xlsx/xlsm etc.) are based on OpenXML format. This is a format created by Microsoft to allow Excel spreadsheets to be opened in other applications and on other operating systems.

Some data stored in the OpenXML format is not accessible via VBA. For example whether an image is linked or embedded.

This class has been created to export this data into a searchable format.

## Usage example:

```vb
Sub USAGE_EXAMPLE()
  'GET DATA FOR WHOLE WORKBOOK:
  Dim raw_data As New RawXLData
  raw_data.Init ActiveWorkbook
  set Sheet1Pictures = raw_data.Data("xl\drawings\drawing1.xml")
  '...

  'ALTERNATIVELY YOU CAN GET DATA FROM SPECIFIC PATHS. THIS IS OFTEN FASTER THAN OBTAINING THE WHOLE WORKBOOK:
  Dim raw_data As New RawXLData
  raw_data.Init ActiveWorkbook, "xl\drawings"
  set Sheet1Pictures     = raw_data.Data("drawing1.xml")
  set Sheet1PicsRels = raw_data.Data("_rels\drawing1.xml.rels")

  'IF YOU WANT TO LOOP OVER THE KEYS OF A WORKBOOK USE THIS:
  Dim raw_data As New RawXLData
  raw_data.Init ActiveWorkbook, "xl\drawings"
  Dim key as Variant
  for each key in raw_data.Data.Keys()
    if key like "_rels*" then
      Dim xml as Object
      xml = raw_data.Data(key)

      '...
      'DO STUFF WITH XML
      '...
    end if
  next key
End Sub
```

# Reference

## Properties

### **`Long`** `ErrorLevel` as integer

If initialisation fails, a non-zero `ErrorLevel` is raised. This helps you prevent custom error handling.

### **`String`** `ErrorText`

If initialisation fails, `ErrorText` will contain information about what went wrong.

### **`String`** `sFilterPath` *(Read Only)*

If a filter path is provided to `Init` then this property will contain the filtered path.

### **`Workbook`** `wbTarget` *(Read Only)*

`wbTarget` will contain the workbook who's data has been exported.

### **`String`** `sDataDir` *(Read Only)*

The data directory created by the cRawXLData. Currently it only uses this directory for parsing, however in the future this may also be used to create modify workbook XML data.

## Methods

### `Init(wb As Workbook, Optional ByVal sFilterPath As String = "") As cRawXLData`

Initialises the class based on the workbook and filter paths you have given it. The return value is a `cRawXLData` object, as this can also be called like a static method.
