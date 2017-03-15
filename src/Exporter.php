<?php

namespace Prezent\ExcelExporter;

/**
 * ExcelExporter
 *
 * @author      Robert-Jan Bijl <robert-jan@prezent.nl>
 */
class Exporter
{
    /**
     * @var \PHPExcel
     */
    private $file;

    /**
     * @var string
     */
    private $currentColumn;

    /**
     * @var string
     */
    protected $maxColumn = 'A';

    /**
     * @var int
     */
    protected $maxRow = 1;

    /**
     * @var int
     */
    private $currentRow;

    /**
     * @var array
     */
    private $columns;

    /**
     * @var string
     */
    private $tempPath;

    /**
     * @var string
     */
    private $fileName;

    /**
     * @var string
     */
    private $filePath;

    /**
     * @param string $tempPath
     */
    public function __construct($tempPath)
    {
        $this->tempPath = $tempPath;
        $this->init();
    }

    /**
     * Initialize the exporter
     *
     * @return bool
     */
    protected function init()
    {
        $this->columns = range('A', 'Z');
        foreach (range('A', 'K') as $first) {
            foreach (range('A', 'Z') as $second) {
                $this->columns[] = sprintf('%s%s', $first, $second);
            }
        }

        $this->currentRow = 1;
        $this->currentColumn = reset($this->columns);
        $this->file = $this->createFile();

        return true;
    }

    public function resetCoordinates()
    {
        $this->currentColumn = reset($this->columns);
        $this->currentRow = 1;
    }

    /**
     * Write data to a row
     *
     * @param array $data
     * @param bool  $finalize
     * @param null  $sheetIndex
     * @throws \Exception
     */
    public function writeRow(array $data = array(), $finalize = true, $sheetIndex = null)
    {
        // get the sheet to write in
        if (null === $sheetIndex) {
            $sheet = $this->file->getActiveSheet();
        } else {
            $sheet = $this->file->getSheet($sheetIndex);
            $this->file->setActiveSheetIndex($sheetIndex);
        }

        $lastDataKey = $this->getLastArrayKey($data);
        foreach ($data as $key => $value) {
            $coordinate = sprintf('%s%d', $this->currentColumn, $this->currentRow);
            $sheet->setCellValue($coordinate, $value);
            if ($key !== $lastDataKey) {
                $this->nextColumn();
            }
        }

        if ($finalize) {
            $this->nextRow();
        }
    }

    /**
     * Generate the file, return its location
     *
     * @param string $filename
     * @param string $format
     * @param bool $disconnect
     * @return array
     */
    public function generateFile($filename, $format = 'Excel2007', $disconnect = true)
    {
        $this->formatFile();
        return $this->writeFileToTmp($filename, $format, $disconnect);
    }

    /**
     * Format the file
     *
     * @return boolean
     */
    protected function formatFile()
    {
        // this base class does not do any formatting. Extend this class if you need specific formatting
        return true;
    }

    /**
     * Output a file to the browser
     *
     * @param string $filePath
     * @param string $fileName
     * @return bool
     */
    public function outputFile($filePath = null, $fileName = null)
    {
        $fileName = null === $fileName ? $this->fileName : $fileName;
        $filePath = null === $filePath ? $this->filePath : $filePath;

        $handler = fopen($filePath, 'r');

        header('Content-Description: File Transfer');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header(sprintf('Content-Disposition: attachment; filename=%s', $fileName));
        header('Content-Transfer-Encoding: chunked');
        header('Expires: 0');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Pragma: public');
        header(sprintf('Content-Length: %s', filesize($filePath)));

        //Send the content in chunks
        while (false !== ($chunk = fread($handler, 4096))) {
            echo $chunk;
        }

        return true;
    }

    /**
     * Set the pointer to the next column
     *
     * @return bool
     */
    private function nextColumn()
    {
        $this->currentColumn = next($this->columns);
        $this->updateMaxColumn($this->currentColumn);

        return true;
    }

    /**
     * Set the pointer to the next row, by default reset to first column
     *
     * @param bool $reset
     * @return bool
     */
    private function nextRow($reset = true)
    {
        $this->currentRow += 1;
        $this->updateMaxRow($this->currentRow);
        if ($reset) {
            $this->currentColumn = reset($this->columns);
        }

        return true;
    }

    /**
     * Update the max colum
     *
     * @param string $currentColumn
     * @return bool
     */
    private function updateMaxColumn($currentColumn)
    {
        $max = $this->maxColumn;

        if (strlen($currentColumn) > strlen($max)) {
            return $this->maxColumn = $currentColumn;
        } elseif (strlen($currentColumn) < strlen($max)) {
            // do nothing
        } elseif (strlen($currentColumn) == 1 && strlen($max) == 1) {
            if ($currentColumn > $max) {
                $this->maxColumn = $currentColumn;
            }
        } elseif ($currentColumn[0] > $max[0]) {
            $this->maxColumn = $currentColumn;
        } elseif (($currentColumn[0] == $max[0]) && $currentColumn[1] > $max[1]) {
            $this->maxColumn = $currentColumn;
        }

        return true;
    }

    /**
     * Update the max row
     *
     * @param $currentRow
     * @return bool
     */
    private function updateMaxRow($currentRow)
    {
        $this->maxRow = max($currentRow, $this->maxRow);

        return true;
    }

    /**
     * Create the PHPExcel instance to work in
     *
     * @return \PHPExcel
     */
    private function createFile()
    {
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '128MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        $file = new \PHPExcel();
        $this->createWorksheets($file);

        return $file;
    }

    /**
     * Create worksheets
     *
     * @param \PHPExcel $file
     * @return \PHPExcel
     */
    protected function createWorksheets(\PHPExcel $file)
    {
        // create one default sheet
        $file->createSheet();

        return $file;
    }


    /**
     * Create excel file and store in tmp dir
     *
     * @param string $filename
     * @param string $format
     * @param bool   $disconnect
     * @throws \Exception
     * @return array
     */
    private function writeFileToTmp($filename, $format = 'Excel2007', $disconnect = true)
    {
        $path = sprintf('%s/%s', $this->tempPath, $filename);
        $objWriter = \PHPExcel_IOFactory::createWriter($this->file, $format);
        $objWriter->save($path);

        if ($disconnect) {
            $this->file->disconnectWorksheets();
            unset($this->file);
        }

        $this->fileName = $filename;
        $this->filePath = $path;

        return array($path, $filename);
    }

    /**
     * Getter for file
     *
     * @return \PHPExcel
     */
    public function getFile()
    {
        return $this->file;
    }

    /**
     * Getter for maxColumn
     *
     * @return string
     */
    public function getMaxColumn()
    {
        return $this->maxColumn;
    }

    /**
     * Getter for maxRow
     *
     * @return int
     */
    public function getMaxRow()
    {
        return $this->maxRow;
    }

    /**
     * Reset the max column back to 'A'
     *
     * @return $this
     */
    public function resetMaxColumn()
    {
        $this->maxColumn = 'A';
        return $this;
    }

    /**
     * Get all used columns, based on the maxColumn
     *
     * @return array
     */
    protected function getUsedColumns()
    {
        if (strlen($this->maxColumn) == 2) {
            $usedColumns = range('A', 'Z');
            $max = $this->maxColumn[0];
            $upperRange = range('A', $max);
            foreach ($upperRange as $x) {
                $end = $max == $x ? $this->maxColumn[1] : 'Z';

                foreach (range('A', $end) as $y) {
                    $usedColumns[] = sprintf('%s%s', $x, $y);
                };
            }
        } else {
            $usedColumns = range('A', $this->maxColumn);
        }

        return $usedColumns;
    }

    /**
     * Getter for currentRow
     *
     * @return int
     */
    public function getCurrentRow()
    {
        return $this->currentRow;
    }

    /**
     * Get the last key of an array
     *
     * @param array $array
     * @return mixed
     */
    private function getLastArrayKey(array $array)
    {
        $arrayKeys = array_keys($array);

        return array_pop($arrayKeys);
    }
}
