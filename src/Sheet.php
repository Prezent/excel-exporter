<?php

declare(strict_types=1);

namespace Prezent\ExcelExporter;

use PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
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
     * @var Worksheet
     */
    private $worksheet;

    /**
     * Sheet constructor.
     * @param Worksheet $worksheet
     */
    public function __construct(Worksheet $worksheet)
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
     *f
     * @param bool $resetMax
     * @return self
     */
    public function resetCoordinates(bool $resetMax = false): self
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
     * Write all data for a sheet at once
     *
     * @param array $data
     * @param bool $advancedValueBinder
     * @return self
     * @throws Exception
     */
    public function writeData(array $data, bool $advancedValueBinder = true): self
    {
        if ($advancedValueBinder) {
            Cell::setValueBinder(new AdvancedValueBinder());
        }

        $this->worksheet->fromArray($data);

        // update the sheet dimension

        // The max row is simply the size of the array
        $this->setMaxRow(count($data));

        // The max column is the max size of one of the rows in the array
        $maxColumn = array_reduce($data, function(int $maxColumn, array $rowData) {
            $maxColumn = max($maxColumn, count($rowData));

            return $maxColumn;
        }, 0);
        $this->setMaxColumn(Coordinate::stringFromColumnIndex($maxColumn));

        return $this;
    }

    /**
     * Write a row in the sheet
     *
     * @param array $data
     * @param bool $finalize
     * @return self
     */
    public function writeRow(array $data = [], bool $finalize = true): self
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
    public function getMaxColumn(): string
    {
        return $this->maxColumn;
    }

    /**
     * Getter for maxRow
     *
     * @param bool $offsetByOne
     * @return int
     */
    public function getMaxRow(bool $offsetByOne = true): int
    {
        if ($offsetByOne) {
            return $this->maxRow - 1;
        }

        return $this->maxRow;
    }

    /**
     * Getter for currentRow
     *
     * @return int
     */
    public function getCurrentRow(): int
    {
        return $this->currentRow;
    }

    /**
     * Getter for currentColumn
     *
     * @return string
     */
    public function getCurrentColumn(): string
    {
        return $this->currentColumn;
    }

    /**
     * Getter for worksheet
     *
     * @return Worksheet
     */
    public function getWorksheet(): Worksheet
    {
        return $this->worksheet;
    }

    /**
     * Get all used columns, based on the maxColumn
     *
     * @return array
     */
    protected function getUsedColumns(): array
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
    public function nextColumn(): bool
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
    public function nextRow(bool $reset = true): bool
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
     * @return mixed
     */
    private function updateMaxColumn(string $currentColumn)
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
     * Setter for maxRow
     *
     * @param int $maxRow
     * @return self
     */
    private function setMaxRow(int $maxRow): self
    {
        $this->maxRow = $maxRow;

        return $this;
    }

    /**
     * Setter for maxColumn
     *
     * @param string $maxColumn
     * @return self
     */
    public function setMaxColumn(string $maxColumn): self
    {
        $this->maxColumn = $maxColumn;

        return $this;
    }

    /**
     * Update the max row
     *
     * @param int $currentRow
     * @return self
     */
    private function updateMaxRow(int $currentRow): self
    {
        return $this->setMaxRow(max($currentRow, $this->maxRow));
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
