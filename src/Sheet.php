<?php

namespace Prezent\ExcelExporter;

/**
 * Prezent\ExcelExporter\Sheet
 *
 * @author Robert-Jan Bijl <robert-jan@prezent.nl>
 */
class Sheet
{
    /**
     * @var string
     */
    private $currentColumn = 'A';

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
    private $currentRow = 1;

    /**
     * @var array
     */
    private $columns = 1;

    /**
     * @var \PHPExcel_Worksheet
     */
    private $worksheet;

    /**
     * Sheet constructor.
     * @param \PHPExcel_Worksheet $worksheet
     */
    public function __construct(\PHPExcel_Worksheet $worksheet)
    {
        $this->worksheet = $worksheet;
        $this->columns = range('A', 'Z');
        foreach (range('A', 'K') as $first) {
            foreach (range('A', 'Z') as $second) {
                $this->columns[] = sprintf('%s%s', $first, $second);
            }
        }

        $this->resetCoordinates(true);
    }

    /**
     * Reset the coordinates back to initial values
     *
     * @param bool $resetMax
     * @return Sheet
     */
    public function resetCoordinates($resetMax = false)
    {
        $this->currentColumn = reset($this->columns);
        $this->currentRow = 1;

        if ($resetMax) {
            $this->maxColumn = reset($this->columns);
            $this->maxRow = 1;
        }

        return $this;
    }

    /**
     * Write a row in the sheet
     *
     * @param array $data
     * @param bool $finalize
     * @return self
     */
    public function writeRow(array $data = [], $finalize = true)
    {
        $lastDataKey = $this->getLastArrayKey($data);
        foreach ($data as $key => $value) {
            $coordinate = sprintf('%s%d', $this->getCurrentColumn(), $this->getCurrentRow());
            $this->worksheet->setCellValue($coordinate, $value);
            if ($key !== $lastDataKey) {
                $this->nextColumn();
            }
        }

        if ($finalize) {
            $this->nextRow();
        }

        return $this;
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
     * @param bool $offsetByOne
     * @return int
     */
    public function getMaxRow($offsetByOne = true)
    {
        if ($offsetByOne) {
            return $this->maxRow -1;
        }

        return $this->maxRow;
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
     * Getter for currentColumn
     *
     * @return string
     */
    public function getCurrentColumn()
    {
        return $this->currentColumn;
    }

    /**
     * Getter for worksheet
     *
     * @return \PHPExcel_Worksheet
     */
    public function getWorksheet()
    {
        return $this->worksheet;
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
     * Set the pointer to the next column
     *
     * @return bool
     */
    public function nextColumn()
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
    public function nextRow($reset = true)
    {
        $this->currentRow += 1;
        $this->updateMaxRow($this->currentRow);
        if ($reset) {
            $this->currentColumn = reset($this->columns);
        }

        return true;
    }

    /**
     * Update the max column
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