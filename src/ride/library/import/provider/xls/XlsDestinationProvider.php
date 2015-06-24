<?php

namespace ride\library\import\provider\xls;

use ride\library\import\provider\DestinationProvider;
use ride\library\import\Importer;

use PHPExcel;
use PHPExcel_Writer_Excel2007;

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
        $sheet = $excel->getSheet();

        $rowNumber = $this->getRowNumber();
        $rowDiff = 1;
        $colNumber = 0;

        foreach ($row as $value) {
            if (is_array($value)) {
                $rowIndex = 0;
                do {
                    $sheet->setCellValueByColumnAndRow($colNumber, $rowNumber + $rowIndex, array_shift($value));
                    $rowIndex++;
                } while ($value);

                $rowDiff = max($rowDiff, $rowIndex);
            } else {
                $sheet->setCellValueByColumnAndRow($colNumber, $rowNumber, $value);
            }

            $colNumber++;
        }

        $this->setRowNumber($rowNumber + $rowDiff);
    }

    /**
     * Performs finishing tasks of the import
     * Writes the PHPExcel object to a file.
     * @return null
     */
    public function postImport() {
        $file = $this->getFile();

        $writer = new PHPExcel_Writer_Excel2007($this->getExcel());
        $writer->save($file->getAbsolutePath());
    }

}
