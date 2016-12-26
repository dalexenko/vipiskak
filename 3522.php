<?

error_reporting(E_ALL);
// Переменные

//$workfile ="Y:\\v3811\\nez352231082010.xls";
//$filename_dir = "d:\\Visual Studio 2008\\Projects\\vipiskak\\Debug\\";


// Функции

function arrah ()
{

global $workfile;
$filename = str_replace("\\", "/" , $workfile);

//$filename = "D:/vipiska/3522/3522.xls";

$sheet1 = "Лист1";
$excel_app = new COM("Excel.application") or Die ("Did not connect");
$excel_app->Visible = 0;
$Workbook = $excel_app->Workbooks->Open("$filename") or Die("Did not open $filename $Workbook");

$i=5;

$result_A = '0';

$Worksheet = $Workbook->Worksheets($sheet1);
$Worksheet->activate;

while ($result_A != '')
{
$col_A = "A".$i;
$col_B = "B".$i;
$col_C = "C".$i;
$col_D = "D".$i;
$col_E = "E".$i;
$col_F = "F".$i;
$col_G = "G".$i;
$col_H = "H".$i;
$col_I = "I".$i;
$col_J = "J".$i;
$col_K = "K".$i;
$col_L = "L".$i;
$col_M = "M".$i;
$col_N = "N".$i;
$col_O = "O".$i;
$col_P = "P".$i;
$col_Q = "Q".$i;
$col_R = "R".$i;
$col_S = "S".$i;


$cell_A = $Worksheet->Range($col_A);
$cell_A->activate;
$result_A = trim ($cell_A->value);

$cell_B = $Worksheet->Range($col_B);
$cell_B->activate;
$result_B = trim ($cell_B->value);

$cell_C = $Worksheet->Range($col_C);
$cell_C->activate;
$result_C = trim ($cell_C->value);

$cell_D = $Worksheet->Range($col_D);
$cell_D->activate;
$result_D = trim ($cell_D->value);

$cell_E = $Worksheet->Range($col_E);
$cell_E->activate;
$result_E = trim ($cell_E->value);

$cell_F = $Worksheet->Range($col_F);
$cell_F->activate;
$result_F = trim ($cell_F->value);

$cell_G = $Worksheet->Range($col_G);
$cell_G->activate;
$result_G = trim ($cell_G->value);

$cell_H = $Worksheet->Range($col_H);
$cell_H->activate;
$result_H = trim ($cell_H->value);

$cell_I = $Worksheet->Range($col_I);
$cell_I->activate;
$result_I = trim ($cell_I->value);

$cell_J = $Worksheet->Range($col_J);
$cell_J->activate;
$result_J = trim ($cell_J->value);

$cell_K = $Worksheet->Range($col_K);
$cell_K->activate;
$result_K = trim ($cell_K->value);

$cell_L = $Worksheet->Range($col_L);
$cell_L->activate;
$result_L = trim ($cell_L->value);

$cell_M = $Worksheet->Range($col_M);
$cell_M->activate;
$result_M = trim ($cell_M->value);

$cell_N = $Worksheet->Range($col_N);
$cell_N->activate;
$result_N = trim ($cell_N->value);

$cell_O = $Worksheet->Range($col_O);
$cell_O->activate;
$result_O = trim ($cell_O->value);

$cell_P = $Worksheet->Range($col_P);
$cell_P->activate;
$result_P = trim ($cell_P->value);

$cell_R = $Worksheet->Range($col_R);
$cell_R->activate;
$result_R = trim ($cell_R->value);

$cell_S = $Worksheet->Range($col_S);
$cell_S->activate;
$result_S = trim ($cell_S->value);


if (strlen($result_A)>5){$mass_rah[$i-5]=$result_D;}

$arr_rah=array_unique($mass_rah);



$i = $i + 1;

}

//echo count($arr_rah)."\n";

$y=0;

$rahlistfile = fopen('rahlist.txt', 'w');

for ($i=0; $i<=count($mass_rah); $i++)
{
if (array_key_exists($i, $arr_rah))
{
//copy ("3522.xls", "clients/".$arr_rah[$i].".xls");

copy ($filename, "clients/".$arr_rah[$i].".xls");

fwrite($rahlistfile, $arr_rah[$i]."\n");
$rah_list[$y]=$arr_rah[$i];
$y=$y+1;
}

}

fclose($rahlistfile);

$excel_app->ActiveWorkbook->Save();

$excel_app->Quit(); //Закрываем приложение

//$excel_app->Release(); //Высвобождаем объекты

$excel_app = null;

//$range = Null;

}

function t()
{

global $work_dir;


$file_dir = str_replace("\\", "/" , $work_dir);

$S=file('rahlist.txt');

$excel_app = new COM("Excel.application") or Die ("Did not connect");
$excel_app->Visible = 0;
$sheet1 = "Лист1";
echo "Обработаны файлы выписок:\n";

for ($y=0; $y<count($S); $y++)
{


$filedest= $file_dir."/clients/".trim($S[$y]).".xls";
$Workbook_blank = $excel_app->Workbooks->Open("$filedest") or Die("Did not open $filedest $Workbook_blank");

$Worksheet_blank = $Workbook_blank->Worksheets($sheet1);
$Worksheet_blank->activate;

$i=5;

$result_A = '0';


while ($result_A != '')
{

$col_A = "A".$i;
$col_E = "D".$i;


$cell_A = $Worksheet_blank->Range($col_A);
$cell_A->activate;
$result_A = trim ($cell_A->value);

$cell_E = $Worksheet_blank->Range($col_E);
$cell_E->activate;
$result_E = trim ($cell_E->value);

if ($result_E <> trim($S[$y]))
{
$range=$excel_app->Range("$i:$i");           // Определяем строку
//$range->EntireRow->Hidden = True;
$range->EntireRow->Delete();
$i=5;
}
else
{
$i = $i + 1;
}

}


//$excel_app->Workbooks->Save($filedest);

//$filedest_saved = str_replace("/", "\\" , $filedest);

$excel_app->ActiveWorkbook->Save();


echo ($y+1).": ".$filedest."\n";


}

$excel_app->Quit(); //Закрываем приложение

//$excel_app->Release(); //Высвобождаем объекты

$excel_app = null;

$range = Null;


}

// обработка XLS файла

arrah ();

// print_r ($S);

t();

$cmd_del_dest = "del x:\\v3522\\*.* /Q";

$cmd_copy_files = "copy ".$work_dir."\\clients\\ x:\\v3522";

$cmd_del_src = "del ".$work_dir."\\clients\\*.* /Q";

exec ($cmd_del_dest);
exec ($cmd_copy_files);
exec ($cmd_del_src);

?>