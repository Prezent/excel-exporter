# prezent/excel-exporter

A wrapper around [PHPExcel](https://github.com/PHPOffice/PHPExcel), to allow for easy export of data to Excel files.

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
list($path, $filename) = $exporter->generateFile(sprintf('%s-ProjectStatus.xlsx', date('Y-m-d')));

// stream to browser
$exporter->outputFile($path, $filename);
```