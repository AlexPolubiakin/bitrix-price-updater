<?php
require 'vendor/autoload.php';
require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_before.php");

$inputFileName = 'goods.xlsx';
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
$worksheet = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

foreach ($worksheet as $item) {
    if ($item["G"] != "Код") {
        $excelData[$item["G"]] = array(
            'P_NAME' => $item["B"],
            'P_CATEGORY' => $item["A"],
            'P_PRICE' => clearPrice($item["D"]),
            'P_CODE' => $item["G"],
            'P_QNT' => $item["C"]
        );
    }
}

$IBLOCK_ID = 2;
$IBLOCK_TYPE = "catalog";
$arSelect = array(
    "ID",
    "NAME",
    "IBLOCK_ID",
    "PROPERTY_NO_AUTO_PRICE",
    "PROPERTY_CML2_BAR_CODE",
    "PROPERTY_Packing_size",
    "PROPERTY_COMPLECT",
    "PROPERTY_NUMORDER",
    "CATALOG_QUANTITY",
    "CATALOG_GROUP_1",
    "CATALOG_PRICE_1",
    "property_PROD_CODE",
);
$arFilter = array("IBLOCK_ID" => IntVal($IBLOCK_ID), "ACTIVE_DATE" => "Y", "ACTIVE" => "Y");

if (CModule::IncludeModule("iblock")) {
    $res = CIBlockElement::GetList(array(), $arFilter, false, array("nPageSize" => 50), $arSelect);

    while ($ob = $res->GetNext()) {
        $bitrixProducts[$ob['PROPERTY_PROD_CODE_VALUE']] = array(
            "ID" => $ob['ID'],
            "NAME" => $ob['NAME'],
            "PROD_CODE" => $ob['PROPERTY_PROD_CODE_VALUE'],
            "PRICE" => $ob['CATALOG_PRICE_1'],
            "QUANTITY" => $ob['CATALOG_QUANTITY'],
        );
    }
}

$updated_prod = 0;
foreach ($excelData as $item => $key) {
    if ($key["P_CODE"] === $bitrixProducts[$key["P_CODE"]]['PROD_CODE']) {
        $elem_id = $bitrixProducts[$key["P_CODE"]]['ID'];
        $arQuantityUpd = array(
            "QUANTITY" => $key["P_QNT"]
        );
        $arPriceUpd = array(
            "PRODUCT_ID" => $elem_id,
            "CATALOG_GROUP_ID" => 1,
            "PRICE" => $key["P_PRICE"],
            "CURRENCY" => "RUB",
        );
        CCatalogProduct::Update($elem_id, $arQuantityUpd);
        CPrice::Update(3, $arPriceUpd);
        $updated_prod++;
    } else {
        $prodToAdd[] = array(
            'NAME' => $key["P_NAME"]
        );
    }
}
echo 'Обновление продуктов завершено. Всего обновлено продуктов: ' . $updated_prod;
echo "<br>";
echo 'Дата обновления: ' . date("Y-m-d H:i:s");
echo "<br>";
echo "<hr>";
echo "Список товаров которые надо добавить на сайт: ";
echo "<br>";
foreach ($prodToAdd as $key) {
    echo $key["NAME"];
    echo "<br>";
}

function clearPrice($item)
{
    $arr_str = str_split(($item));
    foreach ($arr_str as $key => $value) {
        if ($value === ',') {
            unset($arr_str[$key]);
        }
    }
    $result = implode($arr_str);
    return $result;
}


// Примеры данных

// $excelData

// ["НФ-00001092"]=>
//   array(5) {
//     ["P_NAME"]=>
//     string(53) "Фонарь кемпинговый SUPRA SFL-LTR-15L"
//     ["P_CATEGORY"]=>
//     string(12) "ФОНАРИ"
//     ["P_PRICE"]=>
//     string(3) "300"
//     ["P_CODE"]=>
//     string(13) "НФ-00001092"
//     ["P_QNT"]=>
//     string(1) "1"
//   }

//$bitrixProducts

// [0] => Array
//         (
//             [ID] => 213
//             [NAME] => BBK 50LEM-1027/FTS2C
//             [PROD_CODE] => НФ-00001071
//             [PRICE] => 15000.00
//             [QUANTITY] => 0
//         )

// $prodToAdd
// [246] => Array
//         (
//             [NAME] => Электропечь SUPRA MTS-2001B 20л 1300Вт
//         )



