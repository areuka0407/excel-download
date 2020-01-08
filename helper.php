<?php

function dd(){
    call_user_func_array("dump", func_get_args());
    exit;
}

function dump(){
    foreach(func_get_args() as $arg){
        echo "<pre>";
        var_dump($arg);
        echo "</pre>";
    }
}

function filescan($target){
    if(!is_dir($target)) return $target;
    
    $result = [];
    $founds = scandir($target);
    foreach($founds as $found){
        if($found === "." || $found === "..") continue;
        $foundPath = $target.DIRECTORY_SEPARATOR.$found;
        if(is_dir($foundPath)) $result = array_merge($result, filescan($foundPath));
        else $result[] = $foundPath;
    }
    return $result;
}