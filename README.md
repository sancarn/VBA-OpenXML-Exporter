# cRawXLData
Excel workbooks these days (xlsx/xlsm etc.) are based on OpenXML format. This is a format created by Microsoft to allow Excel spreadsheets to be opened in other applications and on other operating systems.

Some data stored in the OpenXML format is not accessible via VBA. For example whether an image is linked or embedded.

This class has been created to export this data into a searchable format.

## Usage example:

```vb
Sub USAGE_EXAMPLE()
  Dim raw_data As New RawXLData
  raw_data.Init ActiveWorkbook
  set Sheet1Pictures = raw_data.Data("xl\drawings\drawing1.xml")
  '...

  'ALTERNATIVE:
  Dim raw_data As New RawXLData
  raw_data.Init ActiveWorkbook, "xl\drawings"
  set Sheet1Pictures     = raw_data.Data("drawing1.xml")
  set Sheet1PicsRels = raw_data.Data("_rels\drawing1.xml.rels")

  'ALTERNATIVE
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
