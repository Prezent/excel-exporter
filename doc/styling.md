# Styling your exports

Upon calling `generateFile()`, the `formatFile()` function is called. In the base `Exporter` class, this function does not contain any logic.
To add some styling to your export files, extend the base `Exporter`, and add logic to the `formatFile()` function:

```php
<?php

use Prezent\ExcelExporter\Exporter;

class StyledExporter extends Exporter
{
    /**
     * {@inheritdoc}
     */
    public function formatFile()
    {
        $sheet = $this->getSheet()->getWorksheet();
        $maxRow = $sheet->getMaxRow();
        $maxColumn = $sheet->getMaxColumn();
    
        // autosize all columns
        foreach(range('A', $maxColumn) as $columnID) {
            $sheet->getColumnDimension($columnID)->setAutoSize(true);
        }
    
        // borders
        $borderStyle = array(
            'borders' => array(
                'outline' => array(
                    'style' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
                    'color' => array('rgb' => '000000'),
                ),
            )
        );
    
        // border around the entire table
        $sheet->getStyle(sprintf('A1:%s%d', $maxColumn, $maxRow))->applyFromArray($borderStyle);
        //border around the hearder
    
        // alignment
        $sheet->getStyle(sprintf('A1:C%d', $maxRow))
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT)
        ;
        $sheet->getStyle(sprintf('D1:%s%d', $maxColumn, $maxRow))
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT)
        ;
    
        // format a specific column as percentage
        $sheet->getStyle(sprintf('G1:G%s', $maxRow))->getNumberFormat()->applyFromArray([
            'code' => \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE
        ]);
    
        return $this;
    }
}
```

Basically, the  `$this->getFile()` function returns an instance of `\PhpOffice\PhpSpreadsheet\Spreadsheet`. From here on, you can use all functions that are defined on it.
For more info, [see the documentation of PhpSpreadsheet](https://phpspreadsheet.readthedocs.io/en/latest/)

## Styling with multiple worksheets
You can also access the worksheets directly, this can come in handy if you have an Exporter with multiple worksheets (see [Advanced Usage](advanced-usage.md)).

```php
<?php

use Prezent\ExcelExporter\Exporter;

class MultipleSheetExporter extends Exporter
{
    /**
     * {@inheritdoc}
     */
    public function formatFile()
    {   
       // formats all sheets in the same way
       foreach ($this->getSheets() as $sheet) {
           $this->formatWorksheet($sheet);
       }
       
       return $this;
    }

    /**
     * Format a signle worksheet
     * 
     * @param Sheet $sheet
     */
    public function formatWorksheet(Sheet $sheet)
    {
        $worksheet = $sheet->getWorksheet();
        
        // format the worksheet here in the same way as above
    }
}
```
