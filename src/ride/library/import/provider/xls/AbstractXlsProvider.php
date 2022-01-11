<?php

namespace ride\library\import\provider\xls;

use ride\library\import\exception\ImportException;
use ride\library\import\provider\FileProvider;
use ride\library\system\file\File;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Abstract import provider for the XLS file format
 */
abstract class AbstractXlsProvider implements FileProvider {

    /**
     * Instance of the file to read or write
     * @var \ride\library\system\file\File
     */
    protected $file;

    /**
     * Instance of the PHPExcel Object
     * @ var \PHPExcel
     */
    protected $excel;

    /**
     * Number of the row
     * @var integer
     */
    protected $rowNumber = 1;

    /**
     * Column names in the file
     * @var array
     */
    protected $columnNames;

    /**
     * Constructs a new XLS Provider
     * @param \ride\library\system\file\File $file
     * @return null
     */
    public function __construct(File $file = null) {
        if ($file) {
            $this->setFile($file);
        }
    }

    /**
     * Sets the file to read/write
     * @param \ride\library\system\file\File $file
     * @return null
     */
    public function setFile(File $file) {
        $this->file = $file;
    }

    /**
     * Gets the file to read/write
     * @return \ride\library\system\file\File
     */
    public function getFile() {
        if (!$this->file) {
            throw new ImportException('Could not get the file: no file set');
        }

        return $this->file;
    }

    /**
     * Sets the instance of PHPExcel
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $excel
     */
    public function setExcel(Spreadsheet $excel) {
        $this->excel = $excel;
    }

    /**
     * Gets the instance of the PHPExcel
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function getExcel() {
        if (!$this->excel) {
            $this->excel = new Spreadsheet();
        }

        return $this->excel;
    }

    /**
     * Sets the row number
     * @param integer $row Number of the current row
     * @return null
     */
    public function setRowNumber($rowNumber) {
        $this->rowNumber = $rowNumber;
    }

    /**
     * Gets the rowNumber
     * @return Int
     */
    public function getRowNumber() {
        return $this->rowNumber;
    }

    /**
     * Maps a column number to a name
     * @var integer $columnIndex Index of the column, starting from 0
     * @var string $columnName Name for the column
     * @return null
     */
    public function setColumnName($columnIndex, $columnName) {
        $this->columnNames[$columnIndex] = $columnName;
    }

    /**
     * Sets the column names for the first row of the output
     * @param array $columnNames Value of getColumnNames of the source provider
     * @return null
     */
    public function setColumnNames(array $columnNames) {
        $this->columnNames = $columnNames;
    }

    /**
     * Gets the available columns for this provider
     * @return array Array with the name of the column as key and as value
     */
    public function getColumnNames() {
        $columns = array();

        foreach ($this->columnNames as $columnIndex => $columnName) {
            $columns[$columnName] = $columnName;
        }

        return $columns;
    }

}