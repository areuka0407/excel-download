
<?php
    require "helper.php";

    // $zip = new ZipArchive();
    // $zip->open("./down.xlsx");
    // $zip->extractTo(__DIR__."/download");

    // exit;

    require "ExcelWriter.php";

    $data = [
        [ "번호", "이름", "나이", "성별" ],
        [ 1, "김민재", 19, "남" ],
        [ 2, "연지수", 18, "여" ],
        [ 3, "조예진", 18, "여" ],
        [ 4, "황성철", 18, "남" ],
        [ 5, "최선한", 39, "남" ]
    ];

    $excel = new ExcelWriter();

    $excel->write($data)->save()->download();

    ?>