<?php
/**
 * 
 * Entry
 * 
 * 기본적으로 하단의 코드를 통해 파일을 읽을 수 있다.
 * $buffer = zip_entry_read($zip_entry, zip_entry_filesize($zip_entry))
 * $xmldata = simplexml_load_string($buffer)
 * 
 * 핵심적인 엑셀 파일을 읽기 위해 필요한 엔트리 파일은 다음과 같다.
 * 
 * - "xl/sharedStrings/xml" 
 *      직접적인 엑셀 파일의 문자열이 담겨져 있는 파일이다.
 *      $xmldata->si :: (object) 각 셀의 정보가 담긴 객체가 담긴 배열
 *      $xmldata->si[i] :: (object) 각 셀의 정보가 담긴 객체 (템플릿을 바탕으로 파일을 만든다면 이 객체로 만들어야 할 것임)
 *      $xmldata->si[i]->t :: (string) 셀의 내용물이 문자열로 담겨져 있다
 * 
 * - "xl/worksheets/sheet1.xml"
 *      각 행렬의 칸 개수에 대한 정보가 담겨져 있는 파일이다.
 *      $xmldata->seetData->row :: (array) '행의 개수'만큼의 길이를 가진 배열
 *      $xmldata->seetData->row[i] :: (object) (i+1)번째 행에 대한 정보가 담긴 '객체'
 *      $xmldata->seetData->row[i]->c :: (array) '열의 개수'만큼의 길이를 가진 배열
 * 
 * 
 */

ini_set("memory_limit", -1);
set_time_limit(0);

class Excel {
    static $template = "./template.xlsx";
    function __construct()
    {
        $this->zip = zip_open(self::$template);
        if(!$this->zip) throw "템플릿 파일을 읽을 수 없습니다.";
        
        while($zip_entry = zip_read($this->zip)){
            $name = zip_entry_name($zip_entry);
            if($name === "xl/sharedStrings.xml") $this->cloneStringObject($zip_entry);
        }
        // zip_close($this->zip);
    }

    function cloneStringObject($zip_entry){
        zip_entry_open($this->zip, $zip_entry, "r");
        $buffer = zip_entry_read($zip_entry, zip_entry_filesize($zip_entry));
        $xmldata = simplexml_load_string($buffer);
        $target = $xmldata->si[0];
        $clone = clone $target;
        $clone->t = "apple";
        // dd($target, $clone);
    }

    /**
     * $arr 에는 행열로 구분된 2차월 배열이 삽입되어야 한다.
     */
    function create($arr){

    }

}