<?php
session_start();
if (isset($_POST['format'])) {
    if ($_POST['format'] == "Word") {
        header("Location:record-word.php");
    } elseif ($_POST['format'] == "Excel") {
        header("Location:record-excel.php");
    }
}