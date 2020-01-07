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