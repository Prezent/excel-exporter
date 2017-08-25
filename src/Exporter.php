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
     * @return \PHPExcel
     */
    private function createFile()
    {
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '128MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        return new \PHPExcel();
    }

    /**
     * Create worksheets
     *
     * @return self
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
     */
    final public function generateFile($filename, $format = 'Excel2007', $disconnect = true)
    {
        list($path, $filename) = $this
            ->formatFile()
            ->writeFileToTmp($filename, $format, $disconnect)
        ;

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
     */
    public function outputFile($fileName = null, $format = 'Excel2007', $disconnect = true)
    {
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
     * @return \PHPExcel
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
}
