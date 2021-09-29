<?php

declare(strict_types=1);

namespace Prezent\ExcelExporter;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * @author Robert-Jan Bijl <robert-jan@prezent.nl>
 */
class Exporter
{
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
     * Array containing the data per sheetIndex
     *
     * @var array
     */
    private $data = [];

    /**
     * @param string $tempPath
     * @throws Exception
     */
    public function __construct(string $tempPath)
    {
        $this->tempPath = $tempPath;
        $this->init();
    }

    /**
     * Initialize the exporter
     *
     * @return self
     * @throws Exception
     */
    final protected function init(): self
    {
        $this->file = $this->createFile();
        $this->createWorksheets()
            ->initWorksheets()
        ;

        return $this;
    }

    /**
     * Create the Spreadsheet instance to work in
     *
     * @return Spreadsheet
     */
    private function createFile(): Spreadsheet
    {
        // todo: caching
        return new Spreadsheet();
    }

    /**
     * Create worksheets
     *
     * @return self
     * @throws Exception
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
    private function initWorksheets(): self
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
     * @return self
     * @throws \Exception
     */
    public function writeRow(array $data = [], int $sheetIndex = 0): self
    {
        $this->data[$sheetIndex][] = $data;

        return $this;
    }

    /**
     * Generate the file, return its location
     *
     * @param string $filename
     * @param string $format
     * @param bool $disconnect
     * @return array
     * @throws Exception
     */
    final public function generateFile(string $filename, string $format = 'Xlsx', bool $disconnect = true): array
    {
        foreach ($this->data as $sheetIndex => $sheetData) {
            $this->getSheet($sheetIndex)->writeData($sheetData);
        }

        // perform the formatting
        $this->formatFile();
        // set the first sheet active, to make sure that is the sheet people see when they open the file
        $this->file->setActiveSheetIndex(0);

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
     * @return void
     * @throws \Exception
     * @throws Exception
     */
    public function outputFile(string $fileName = '', string $format = 'Xlsx', bool $disconnect = true)
    {
        if (!$this->generated) {
            $this->generateFile($fileName, $format, $disconnect);
        }

        if (empty($fileName)) {
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
    }

    /**
     * Create Excel file and store in tmp dir
     *
     * @param string $filename
     * @param string $format
     * @param bool $disconnect
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeFileToTmp(string $filename, string $format = 'Xlsx', bool $disconnect = true): array
    {
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
    public function setGenerated(bool $generated): self
    {
        $this->generated = $generated;

        return $this;
    }

    /**
     * Getter for file
     *
     * @return Spreadsheet
     */
    public function getFile(): Spreadsheet
    {
        return $this->file;
    }

    /**
     * Getter for sheets
     *
     * @return Sheet[]
     */
    public function getSheets(): array
    {
        return $this->sheets;
    }

    /**
     * Get a specific sheet, by sheetIndex
     *
     * @param int $sheetIndex
     * @return Sheet
     */
    public function getSheet(int $sheetIndex = 0): Sheet
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
    public function setWorksheetTitle(string $title, int $sheetIndex = 0): self
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
    public function getTempPath(): string
    {
        return $this->tempPath;
    }

    /**
     * Getter for fileName
     *
     * @return string
     */
    public function getFileName(): string
    {
        return $this->fileName;
    }

    /**
     * Getter for filePath
     *
     * @return string
     */
    public function getFilePath(): string
    {
        return $this->filePath;
    }

    /**
     * Getter for generated
     *
     * @return bool
     */
    public function isGenerated(): bool
    {
        return $this->generated;
    }
}
