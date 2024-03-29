# prezent/excel-exporter

A wrapper around [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet), to allow for easy export of data to Excel files.

## Installation

This extension can be installed using Composer. Tell composer to install the extension:

```bash
$ php composer.phar require prezent/excel-exporter
```

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
