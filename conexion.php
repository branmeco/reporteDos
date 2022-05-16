<?php
$mysqli = new mysqli("localhost", "root", "", "planillas_notas", "81");

if($mysql->connect_errno){
    echo 'Fallo la conexion' . $mysql->connect_error;
    die();
}
?>