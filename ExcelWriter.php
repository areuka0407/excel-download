<?php

defined("DS") OR define("DS", DIRECTORY_SEPARATOR);
defined("EXCEL_ROOT") OR define("EXCEL_ROOT", __DIR__);

error_reporting(E_ALL);
ini_set("memory_limit", -1);
set_time_limit(0);


class ExcelWriter {
    static $stringPath = "/xl/sharedStrings.xml";
    static $sheetPath = "/xl/worksheets/sheet1.xml";

    function __construct()
    {
        $this->filename = "./test.xlsx";
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
        $this->sheet = simplexml_load_file($this->temp_path . self::$sheetPath);
        $this->string = simplexml_load_file($this->temp_path . self::$stringPath);

        $zip->close();
        // dd($this->string, $this->sheet);
    }

    // 2차원 배열을 입력 받아 엑셀을 작성한다
    public function _write($data){
        $colName = "ABCD";
        $this->string['count'] = 0;
        $this->string['uniqueCount'] = 0;

        $idx = 0;
        foreach($data as $idx => $item){
            $this->sheet->sheetData->row[$idx]['r'] = (string)($idx + 1);
            $this->sheet->sheetData->row[$idx]['spans'] = "1:4";
            
            for($i = 0; $i < count($item); $i ++){
                $cell = $item[$i];
                // 문자열이면 string에 넣어줘야함
                if(gettype($cell) === "string"){
                    $stringIdx = $this->string['count'] * 1;

                    // 스트링 데이터
                    $this->string['uniqueCount'] = (string)($stringIdx + 1);
                    $this->string['count'] = (string)($stringIdx + 1);
                    
                    $this->string->si[$stringIdx]->t = $cell;

                    // 시트 데이터
                    $this->sheet->sheetData->row[$idx]->c[$i]['r'] = $colName[$i].($idx + 1);
                    $this->sheet->sheetData->row[$idx]->c[$i]['t'] = "s"; // 문자열일 때만 t 속성 있어야함
                    $this->sheet->sheetData->row[$idx]->c[$i]->v = $stringIdx;
                }
                else {
                    $this->sheet->sheetData->row[$idx]->c[$i]['r'] = $colName[$i].($idx + 1);
                    $this->sheet->sheetData->row[$idx]->c[$i]->v = $cell;
                }
            }   
        }

        $this->string->asXML($this->temp_path . self::$stringPath);
        $this->sheet->asXML($this->temp_path . self::$sheetPath);

        return $this;
    }

    public function write($data)
    {

        $cellName = ["A", "B", "C", "D"];  //샘플데이터가 4열까지만 이뤄져있으니 이 4개면 된다.
        $this->string['count'] = "0";
        $this->string['uniqueCount'] = "0";
        foreach ($data as $idx => $item){
            $this->sheet->sheetData->row[$idx]['r'] = (string)($idx + 1);
            //샘플데이터가 4열까지 정방으로 채워지니 여기는 무조건 1:4로 놓으면 된다.
            //만약 샘플데이터가 행마다 열의 갯수가 다르다면 카운트를 세어서 직접 적어주어야 한다.
            $this->sheet->sheetData->row[$idx]['spans'] = (string)("1:4");
            for($i = 0; $i < count($item); $i++){
                $cell = $item[$i];
                if(gettype($cell) === "string" ){
                    //입력데이터가 문자일경우 string에 먼저 삽입하고 셀을 추가하는 방식으로 해야한다.
                    //현재 스트링 인덱스를 하나 추가시켜주고
                    $stringIdx = $this->string['count'] * 1;

                    $this->string['count'] = (string)($stringIdx + 1);
                    $this->string['uniqueCount'] = (string)($stringIdx + 1);
                    //스트링에 추가한다.
                    //$stringXML = simplexml_load_string("<t>{$cell}</t>");
                    $this->string->si[$stringIdx]->t = $cell;
                    //이후에 셀데이터를 xml로 만든다.
                    $this->sheet->sheetData->row[$idx]->c[$i]['r'] = $cellName[$i] . ($idx + 1);
                    $this->sheet->sheetData->row[$idx]->c[$i]['t'] = 's';
                    $this->sheet->sheetData->row[$idx]->c[$i]->v = $stringIdx;
                }else {
                    //입력데이터가 숫자면 바로 셀로 삽입한다.
                    $this->sheet->sheetData->row[$idx]->c[$i]['r'] = $cellName[$i] . ($idx + 1);
                    $this->sheet->sheetData->row[$idx]->c[$i]->v =  $cell;
                }
            }
        }
        //xml 파일로 저장해주고
        $this->sheet->asXML($this->temp_path . '/xl/worksheets/sheet1.xml');
        $this->string->asXML($this->temp_path . '/xl/sharedStrings.xml');

        return $this;
    }


    public function save(){
        $zip = new ZipArchive();
        $result = $zip->open($this->filename, ZipArchive::OVERWRITE | ZipArchive::CREATE);

        $result or die("아카이브 파일을 열 수 없습니다.");

        $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($this->temp_path));

        foreach($files as $file){
            if($file->isDir()) continue;
            $filePath = $file->getRealPath();
            $relativePath = substr($filePath, strlen($this->temp_path) + 1);
            $zip->addFile($filePath, $relativePath);
        }
        $zip->close();

        $zip->open($this->filename);
        $zip->extractTo(__DIR__."/compare");
        $zip->close();

        return $this;
    }

    public function download(){
        $target = $this->filename;
        if($target){
            header('Content-Type: application/octet-stream');
            header("Content-Disposition:attachement; filename=down.xlsx");
            
            ob_clean();
            
            readfile($target);
            exit;
        }
    }
}