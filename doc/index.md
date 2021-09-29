# prezent/excel-exporter

A wrapper around [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet), to allow for easy export of data to Excel files.

## Index

1. [Installation](installation.md)
2. [Advanced Usage](advanced-usage.md)
3. [Styling your exports](styling.md)

## Quick example

```php
<?php

use Prezent\ExcelExporter\Exporter;

$exporter = new Exporter($tempDir);

$data = ['foo', 'bar'];
$exporter->writeRow($data);

// generate the file
list($path, $filename) = $exporter->generateFile('export.xlsx');

// stream to browser
$exporter->outputFile($filename);
```

## Documentation

The complete documentation can be found in the [doc directory](doc/index.md).
