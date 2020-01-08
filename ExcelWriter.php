<?php

defined("DS") OR define("DS", DIRECTORY_SEPARATOR);
defined("EXCEL_ROOT") OR define("EXCEL_ROOT", __DIR__);

ini_set("memory_limit", -1);
set_time_limit(0);


class ExcelWriter {
    static $sheetInitialString = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">
                                    <dimension ref=\"A1:A1\"/>
                                    <sheetViews>
                                        <sheetView tabSelected=\"1\" workbookViewId=\"0\">
                                        <selection activeCell=\"A1\" sqref=\"A1\"/>
                                    </sheetView>
                                    </sheetViews>
                                    <sheetFormatPr defaultRowHeight=\"16.5\" x14ac:dyDescent=\"0.3\"/>
                                    <sheetData></sheetData>
                                    <phoneticPr fontId=\"1\" type=\"noConversion\"/>
                                    <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>
                                </worksheet>";

    static $stringInitialString = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>";


    static $stringPath = "/xl/sharedStrings.xml";
    static $sheetPath = "/xl/worksheets/sheet1.xml";

    function __construct()
    {
        $this->template = EXCEL_ROOT.DS."template.xlsx";
        $this->temp_path = EXCEL_ROOT.DS."temp";

        is_dir($this->temp_path) && $this->removeTemp();
        mkdir($this->temp_path);
        $this->create_emptry_file();

    }

    // 재귀적으로 디렉토리를 삭제한다
    private function removeTemp($path = null){
        $path = $path ? $path : $this->temp_path;
        $items = scandir($path);
        foreach($items as $item){
            if($item === "." || $item === "..") continue;
            if(is_dir($path.DS.$item)) $this->removeTemp($path.DS.$item);
            else @unlink($path.DS.$item);
        }
        rmdir($path);
    }

    // 템플릿 파일을 연다
    private function create_emptry_file(){
        $zip = new ZipArchive();
        $result = $zip->open($this->template);
        if(!$result) throw "템플릿 파일을 읽을 수 없습니다.";
        $zip->extractTo($this->temp_path);
        $zip->close();

        $this->sheet = new SimpleXMLElement(self::$sheetInitialString);
        $this->string = new SimpleXMLElement(self::$stringInitialString);
    }

    private function takeCelName($row, $col){
        $strings = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        $strlen = strlen($strings);
        $colName = "";
        while($col >= 0){
            if($col === 0) {
                $colName .= "A";
                break;
            }
            $colName .= $strings[$col % $strlen];
            $col = (int)($col / $strlen) - 1;
        }
        $colName = strrev($colName);
        $row++;
        return "{$colName}{$row}";
    }

    private function addString($data){
        $current_index = (int)$this->string['count'];
        $this->string['count'] = $current_index + 1;
        $this->string['uniqueCount'] = (int)$this->string['uniqueCount'] + 1;;
        $si = $this->string->addChild("si");
        $si->t = $data;
        $ph = $si->addChild("phoneticPr");
        $ph->addAttribute("fontId", "1");
        $ph->addAttribute("type", "noConversion");
        return $current_index;
    }
    private function addColumn($rowXML, $row, $rowData = []){
        foreach($rowData as $col => $colData){
            $c = $rowXML->addChild("c");
            $c->addAttribute("r", $this->takeCelName($row, $col));
            $c->addAttribute("t", "s");
            $c->v = $this->addString($colData);
        }
    }

    // 2차원 배열을 입력 받아 엑셀을 작성한다
    public function write($data){
        $max_colcnt = 0;
        $rowCount = count($data);
        foreach($data as $row => $rowData){
            $colCount = count($rowData);
            max($max_colcnt, $colCount);

            $rowXML = $this->sheet->sheetData->addChild("row");
            $rowXML->addAttribute("r", $row + 1);
            $rowXML->addAttribute("spans", "1:".$colCount);
            $rowXML->addAttribute("x14ac:dyDescent", "0.3");
            $this->addColumn($rowXML, $row, $rowData);
        }
        $this->sheet->dimension['ref'] = "A1:".$this->takeCelName($rowCount - 1, $colCount -1);

        $this->string->asXML($this->temp_path . self::$stringPath);
        $this->sheet->asXML($this->temp_path . self::$sheetPath);

        return $this;
    }

    public function save($filename){
        $filename = __DIR__.DS.$filename.".xlsx";
        touch($filename);

        $zip = new ZipArchive();
        $result = $zip->open($filename, ZipArchive::OVERWRITE);

        $result or die("아카이브 파일을 열 수 없습니다.");

        $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($this->temp_path));
        foreach($files as $name => $file){
            if($file->isDir()) continue;
            $filePath = $file->getRealPath();
            $relativePath = substr($filePath, strlen($this->temp_path));
            $zip->addFile($filePath, $relativePath);
        }
        $zip->close();

        $zip->open($filename);
        $zip->extractTo(__DIR__.DS."/compare");
        $zip->close();

        // $this->removeTemp();
    }

    public function download($data){
        header("Content-Type: application/vnd.ms-excel");
        header("Content-Disposition: attachement; filename=excel-download.xlsx");
        echo $data;
    }
}