# Advanced Usage

## Multiple worksheets
By default, this library creates an Excel file with only one worksheet. But you can also use this library to create Excel files with multiple worksheets. 
To to this you have to extend the base `Exporter` class, and add logic to the `createWorksheets` function:
 
```php
<?php

use Prezent\ExcelExporter\Exporter;

class MultipleSheetExporter extends Exporter
{
   /**
    * {@inheritdoc}
    */
   protected function createWorksheets()
   {
       $this->getFile()->createSheet(0)->setTitle('First Worksheet');
       $this->getFile()->createSheet(1)->setTitle('Second Worksheet');

       return true;
   } 
}
```

You can now write data to the different sheets, by passing the sheet index to the `writeRow` call:

```php
<?php

$multipleSheetExporter = new MultipleSheetExporter();
$multipleSheetExporter->writeRow(['data', 'for', 'first', 'sheet'], 0);
$multipleSheetExporter->writeRow(['second', 'sheet', 'gets', 'this'], 1);
```