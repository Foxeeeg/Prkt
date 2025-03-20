<?php
    session_start();
    if (!isset($_SESSION['user'])) {
        header("Location:../index.php");
    }
    if (isset($_POST['fio'])) {
        $_SESSION['student'] = $_POST['fio'];
    }else{
        header("Location:record-book.php");
    }
?>
<!DOCTYPE php>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Выбор формата зачетки</title>
    <link rel="stylesheet" href="../css/choose.css">
    <link rel="icon" type="image/png" href="../images/icons/logo.png">
</head>
<body>
    <div class="container">
        <div class="border"></div>
        <form action="../php/choose.php" class="choose-form" method="POST">
            <div class="link-back"><a href="record-book.php">Назад</a></div>
            <h2>Выберите формат зачетной книжки</h2>
            <div class="submit-place">
                <input type="submit" name="format" value="Word">
                <input type="submit" name="format" value="Excel">
            </div>
        </form>
    </div>
    <script type="text/javascript" src="../js/height-page.js"></script>
</body>
</html>