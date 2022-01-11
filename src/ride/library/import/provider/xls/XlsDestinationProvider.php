<?php

namespace ride\library\import\provider\xls;

use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use ride\library\import\provider\DestinationProvider;
use ride\library\import\Importer;

/*
 * Destination provider for the XLS file type.
 * This Provider uses the PHPExcel library to produce an XLS file.
 */
class XlsDestinationProvider extends AbstractXlsProvider implements DestinationProvider {

    /**
     * Performs preparation tasks of the import
     * @param \ride\library\import\Importer $importer
     * @return null
     */
    public function preImport(Importer $importer) {
        if ($this->columnNames) {
            $this->setRow($this->columnNames);
        }
    }

    /**
     * Imports a row into this destination
     * @param array $row Array with the name of the column as key and the value
     * to import as value
     * @return null
     */
    public function setRow(array $row) {
        $excel = $this->getExcel();
        $sheet = $excel->getSheet(0);

        $rowNumber = $this->getRowNumber();
        $rowDiff = 1;
        $colNumber = 0;
        $singleColumns = array();

        foreach ($row as $value) {
            if (is_array($value)) {
                // array value column
                $rowIndex = 0;
                do {
                    $sheet->setCellValueByColumnAndRow($colNumber, $rowNumber + $rowIndex, array_shift($value));
                    $rowIndex++;
                } while ($value);

                $rowDiff = max($rowDiff, $rowIndex);
            } else {
                // single value column
                $sheet->setCellValueByColumnAndRow($colNumber, $rowNumber, $value);

                $singleColumns[$colNumber] = $value;
            }

            $colNumber++;
        }

        $newRowNumber = $rowNumber + $rowDiff;

        // fill single columns up until the maximum of array columns
        if ($rowDiff > 1) {
            for ($i = $rowNumber; $i <= $newRowNumber; $i++) {
                foreach ($singleColumns as $colNumber => $value) {
                    $sheet->setCellValueByColumnAndRow($colNumber, $i, $value);
                }
            }
        }

        $this->setRowNumber($newRowNumber);
    }

    /**
     * Performs finishing tasks of the import
     * Writes the PHPExcel object to a file.
     * @return null
     */
    public function postImport() {
        $file = $this->getFile();

        $writer = new Xlsx($this->getExcel());
        $writer->save($file->getAbsolutePath());
    }

}
