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

use Atico\SpreadsheetTranslator\Core\Exception\SheetNameNotFoundException;
use Atico\SpreadsheetTranslator\Core\Reader\AbstractArrayReader;
use Atico\SpreadsheetTranslator\Core\Reader\ReaderInterface;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReaderBase;

class XlsxReader extends AbstractArrayReader implements ReaderInterface
{
    /** @var Worksheet[] $sheets */
    protected $sheets;

    /** @var Spreadsheet $excel */
    protected $excel;

    /**
     * @throws \Exception
     */
    function __construct($filePath)
    {
        $reader = new XlsxReaderBase();
        $reader->setReadDataOnly(true);
        $this->excel = $reader->load($filePath);

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
     * @param Worksheet $sheet
     * @return mixed
     */
    public function getData($sheet)
    {
        return $sheet->toArray();
    }

    /**
     * @throws SheetNameNotFoundException
     */
    public function getDataBySheetName($name)
    {
        $sheetName = $this->excel->getSheetByName($name);

        if (empty($sheetName)) {
            throw SheetNameNotFoundException::create($name);
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
