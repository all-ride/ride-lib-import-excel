<?php

namespace ride\library\import\provider\xls;

use PhpOffice\PhpSpreadsheet\IOFactory;
use ride\library\import\exception\ImportException;
use ride\library\import\provider\SourceProvider;
use ride\library\import\Importer;
use ride\library\system\file\File;


class XlsSourceProvider extends AbstractXlsProvider implements SourceProvider {
    /**
     * Constructs a new source provider
     *
     * @param File $file
     * @param int $worksheetNumber
     */
    public function __construct(File $file, $worksheetNumber = 0) {
        $this->setFile($file);
        $this->setWorksheetNumber($worksheetNumber);
        $this->firstRowAreColumnNames = false;
    }

    /**
     * Sets the worksheet number
     * @param integer $worksheetNumber
     * @return null
     */
    public function setWorksheetNumber($worksheetNumber) {
        $this->worksheetNumber = $worksheetNumber;
        $this->columnNames = null;
    }

    /**
     * Sets whether the first row are the column names
     * @param boolean $firstRowAreColumnNames
     * @return null
     */
    public function setFirstRowAreColumnNames($firstRowAreColumnNames) {
        $this->firstRowAreColumnNames = $firstRowAreColumnNames;
    }

    /**
     * Gets whether the first row are the column names
     * @return boolean
     */
    public function getFirstRowAreColumnNames() {
        return $this->firstRowAreColumnNames;
    }

    /**
     * Gets the available column names for this provider
     * @return array Array with the name of the column as key and as value
     */
    public function getColumnNames() {
        if ($this->columnNames !== null) {
            return $this->columnNames;
        }

        $this->readFile();

        if ($this->firstRowAreColumnNames) {
            $this->columnNames = $this->getRow();
            foreach ($this->columnNames as $index => $columnName) {
                unset($this->columnNames[$index]);
                $this->columnNames[$columnName] = $columnName;
            }
        } else {
            $this->columnNames = array();

            for ($column = 'A'; $column != $this->highestColumnNumber; $column++) {
                $this->columnNames[$column] = $column;
            }
        }

        return $this->columnNames;
    }

    /**
     * Reads the necessairy data from the file to initialize this provider
     * @return null
     */
    private function readFile() {
        try {
            $inputFileType = IOFactory::identify($this->file);
            $objReader = IOFactory::createReader($inputFileType);
            $objPHPExcel = $objReader->load($this->file);
        } catch(\Exception $exception) {
            throw new ImportException('Could not read file: ' . $this->file, 0, $exception);
        }

        $this->worksheet = $objPHPExcel->getSheet($this->worksheetNumber);
        $this->highestRowNumber = $this->worksheet->getHighestRow();
        $this->highestColumnNumber = $this->worksheet->getHighestColumn();
        $this->rowNumber = 1;
    }

    /**
     * Performs preparation tasks of the import
     * @return null
     */
    public function preImport(Importer $importer) {

    }

    /**
     * Performs finishing tasks of the import
     * @return null
     */
    public function postImport() {

    }

    /**
     * Gets the next row from this destination
     * @return array|null $data Array with the name of the column as key and the
     * value to import as value. Null is returned when all rows are processed.
     */
    public function getRow() {
        if ($this->rowNumber > $this->highestRowNumber) {
            return null;
        }

        $row = $this->worksheet->rangeToArray('A' . $this->rowNumber . ':' . $this->highestColumnNumber . $this->rowNumber, null, true, false);
        $row = $row[0];

        if ($this->columnNames) {
            $index = 0;
            foreach ($this->columnNames as $columnName) {
                $row[$columnName] = $row[$index];
                unset($row[$index]);

                $index++;
            }
        }

        $this->rowNumber++;

        return $row;
    }

}
