<?php

namespace Prezent\ExcelExporter;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Prezent\ExcelExporter\ExcelExporter
 *
 * @author      Robert-Jan Bijl <robert-jan@prezent.nl>
 */
class Exporter
{
    /**
     * Mapping from PHPExcel formats to PHPSpreadsheet formats, for BC
     * @var array
     */
    private $formatMapping = [
        'CSV' => 'Csv',
        'Excel2003XML' => 'Xml',
        'Excel2007' => 'Xlsx',
        'Excel5' => 'Xls',
        'Gnumeric' => 'Gnumeric',
        'HTML' => 'Html',
        'OOCalc' => 'Ods',
        'OpenDocument' => 'Ods',
        'PDF' => 'Pdf',
        'SYLK' => 'Slk',
    ];

    /**
     * @var Spreadsheet
     */
    private $file;

    /**
     * @var Sheet[]
     */
    private $sheets = [];

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
     * @var bool
     */
    private $generated = false;

    /**
     * @param string $tempPath
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function __construct($tempPath)
    {
        $this->tempPath = $tempPath;
        $this->init();
    }

    /**
     * Initialize the exporter
     *
     * @return self
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    final protected function init()
    {
        $this->file = $this->createFile();
        $this->createWorksheets()
            ->initWorksheets()
        ;

        return $this;
    }

    /**
     * Create the PHPExcel instance to work in
     *
     * @return Spreadsheet
     */
    private function createFile()
    {
        // todo: caching
        return new Spreadsheet();
    }

    /**
     * Create worksheets
     *
     * @return self
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function createWorksheets()
    {
        // create one default sheet
        $this->file->createSheet();

        return $this;
    }

    /**
     * Initialize the worksheets, by creating a Sheet object for all of them
     *
     * @return self
     */
    private function initWorksheets()
    {
        foreach ($this->file->getAllSheets() as $index => $worksheet) {
            $this->sheets[$index] = new Sheet($worksheet);
        }

        return $this;
    }

    /**
     * Write data to a row
     *
     * @param array $data
     * @param int $sheetIndex
     * @param bool $finalize
     * @return Sheet
     * @throws \Exception
     */
    public function writeRow(array $data = [], $sheetIndex = 0, $finalize = true)
    {
        $sheet = $this->sheets[$sheetIndex];
        return $sheet->writeRow($data, $finalize);
    }

    /**
     * Generate the file, return its location
     *
     * @param string $filename
     * @param string $format
     * @param bool $disconnect
     * @return array
     * @throws \Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    final public function generateFile($filename, $format = 'Xlsx', $disconnect = true)
    {
        // perform the formatting
        $this->formatFile();
        // set the first sheet active, to make sure that is the sheet people see when they open the file
        $this->file->setActiveSheetIndex(0);

        $format = $this->convertFormat($format);
        list($path, $filename) = $this->writeFileToTmp($filename, $format, $disconnect);
        $this->setGenerated(true);

        return array($path, $filename);
    }

    /**
     * Format the file
     *
     * @return self
     */
    protected function formatFile()
    {
        // this base class does not do any formatting. Extend this class if you need specific formatting
        return $this;
    }

    /**
     * Output a file to the browser
     *
     * @param string $fileName
     * @param string $format
     * @param bool $disconnect
     * @return bool
     * @throws \Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function outputFile($fileName = null, $format = 'Xlsx', $disconnect = true)
    {
        $format = $this->convertFormat($format);
        if (!$this->generated) {
            $this->generateFile($fileName, $format, $disconnect);
        }

        if (null === $fileName) {
            $fileName = $this->fileName;
        }

        $handler = fopen($this->filePath, 'r');

        header('Content-Description: File Transfer');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header(sprintf('Content-Disposition: attachment; filename=%s', $fileName));
        header('Content-Transfer-Encoding: chunked');
        header('Expires: 0');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Pragma: public');
        header(sprintf('Content-Length: %s', filesize($this->filePath)));

        // Send the content in chunks
        while (!feof($handler)) {
            echo fread($handler, 4096);
        }

        return true;
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
    private function writeFileToTmp($filename, $format = 'Xlsx', $disconnect = true)
    {
        $format = $this->convertFormat($format);
        $path = sprintf('%s/%s', $this->tempPath, $filename);

        $objWriter = IOFactory::createWriter($this->file, $format);
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
     * Setter for generated
     *
     * @param bool $generated
     * @return self
     */
    public function setGenerated($generated)
    {
        $this->generated = $generated;
        return $this;
    }

    /**
     * Getter for file
     *
     * @return Spreadsheet
     */
    public function getFile()
    {
        return $this->file;
    }

    /**
     * Getter for sheets
     *
     * @return Sheet[]
     */
    public function getSheets()
    {
        return $this->sheets;
    }

    /**
     * Get a specific sheet, by sheetIndex
     *
     * @param $sheetIndex
     * @return Sheet
     */
    public function getSheet($sheetIndex = 0)
    {
        if (!isset($this->sheets[$sheetIndex])) {
            throw new \InvalidArgumentException(sprintf('No sheet with index %d defined', $sheetIndex));
        }

        return $this->sheets[$sheetIndex];
    }

    /**
     * Set the title for a certain worksheet. Defaults to the first sheet
     *
     * @param string $title
     * @param int $sheetIndex
     * @return self
     */
    public function setWorksheetTitle($title, $sheetIndex = 0)
    {
        if (!isset($this->sheets[$sheetIndex])) {
            throw new \InvalidArgumentException(sprintf('No sheet with index %d defined', $sheetIndex));
        }

        $this->sheets[$sheetIndex]->getWorksheet()->setTitle($title);
        return $this;
    }

    /**
     * Getter for tempPath
     *
     * @return string
     */
    public function getTempPath()
    {
        return $this->tempPath;
    }

    /**
     * Getter for fileName
     *
     * @return string
     */
    public function getFileName()
    {
        return $this->fileName;
    }

    /**
     * Getter for filePath
     *
     * @return string
     */
    public function getFilePath()
    {
        return $this->filePath;
    }

    /**
     * Getter for generated
     *
     * @return bool
     */
    public function isGenerated()
    {
        return $this->generated;
    }
}
