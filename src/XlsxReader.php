<?php

/*
 * This file is part of the Atico/SpreadsheetTranslator package.
 *
 * (c) Samuel Vicent <samuelvicent@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Atico\SpreadsheetTranslator\Reader\Xlsx;

use Atico\SpreadsheetTranslator\Core\Exception\SheetNameNotFound;
use Atico\SpreadsheetTranslator\Core\Reader\AbstractArrayReader;
use Atico\SpreadsheetTranslator\Core\Reader\ReaderInterface;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Worksheet;

class XlsxReader extends AbstractArrayReader implements ReaderInterface
{
    /** @var \PHPExcel_Worksheet[] $sheets */
    protected $sheets;

    /** @var PHPExcel $excel */
    protected $excel;

    /**
     * @throws \PHPExcel_Reader_Exception
     */
    function __construct($filePath)
    {
        $this->excel = PHPExcel_IOFactory::load($filePath);
        $this->sheets = null;
    }

    public function getSheetNames()
    {
        return $this->excel->getSheetNames();
    }

    public function getTitle($index)
    {
        return $this->getSheetIndex($index)->getTitle();
    }

    /**
     * @param PHPExcel_Worksheet $sheet
     * @return mixed
     */
    public function getData($sheet)
    {
        return $sheet->toArray();
    }

    /**
     * @throws SheetNameNotFound
     */
    public function getDataBySheetName($name)
    {
        $sheetName = $this->excel->getSheetByName($name);

        if (empty($sheetName)) {
            throw SheetNameNotFound::create($name);
        }

        return $this->getData($sheetName);
    }

    public function getSheets()
    {
        if (null === $this->sheets) {
            $this->sheets = $this->excel->getAllSheets();
        }

        return $this->sheets;
    }

    public function getSheetIndex($index)
    {
        return $this->getSheets()[$index]; 
    }
}
