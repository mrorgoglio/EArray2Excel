EArray2Excel
============

EArray2Excel is a Yii framework extension.

Introduction
============
EArray2Excel extension converts an associative array in excel sheets using PHPExcel.

The associative array must be in the form:

    {
      {'column_name_A'=>$column_value_a1,'column_name_b'=>$column_value_b1,...},
      {'column_name_A'=>$column_value_a2,'column_name_b'=>$column_value_b2,...},
      ...
    }

EArray2Excel supports multiple sheets; in this case the associative array must be in the form:

    {
      {'sheet_name_1'=>{
        {'column_name_A'=>$column_value_a1,'column_name_b'=>$column_value_b1,...},
        {'column_name_A'=>$column_value_a2,'column_name_b'=>$column_value_b2,...},
        ...,
      },
      'sheet_name_2'=>{
        {'column_name_A'=>$column_value_a1,'column_name_b'=>$column_value_b1,...},
        {'column_name_A'=>$column_value_a2,'column_name_b'=>$column_value_b2,...},
        ...,
      },
      ...
    }


How to use it
============

Requriments
-----------

* Yii 1.1 or above
* PHPExcel 1.7.8 or above

  In order to use EArray2Excel you need to install [PHPExcel](http://phpexcel.codeplex.com/releases/view/107442) first.

  For more information about PHPExcel please visit the [PHPExcel](http://phpexcel.codeplex.com/) website.

Installation
------

1.Add EArray2Excel.php to your extension directory

2.Update your main.php to import this new extension

    // autoloading model and component classes
    'import'=>array(
      'application.models.*',
      'application.components.*',
      'ext.EArray2Excel',
      ...
    ),

3.(Optional) Set the $libPath variable (EArray2Excel.php, row 21) with the path of PHP excel lib.
  
   NOTE: $libPath variable is already set with a standard path. Please modify it carefully.
   You can also specify a different lib path later, in the export phase. Please refer to 'Advanced Usage' to know how to do it.

Basic Usage
------

1.Create an associative array

    ...
    //In this example we suppose you have a db table named 'users' with 'first_name', 'last_name' and 'city' fields.
    $sql = "SELECT first_name, last_name, city FROM users LIMIT 10;";
    $sheet = Yii::app()->db->createCommand($sql)->queryAll();
    ...

2.Export

    ...
    $objEArray2Excel = new EArray2Excel();
    $objEArray2Excel->export($sheet,'my_users'); 
    ...

Advanced Usage
------

1.Create an associative array with (optional) multiple sheets

    ...
    //In this example we suppose you have a db table named 'users' with 'first_name', 'last_name' and 'city' fields.
    $sql = "SELECT first_name, last_name, city FROM users LIMIT 10;";
    $sheets['My Users'] = Yii::app()->db->createCommand($sql)->queryAll();
    $sql = "SELECT city FROM users ORDER BY city DESC LIMIT 10;";
    $sheets['My Cities'] = Yii::app()->db->createCommand($sql)->queryAll();
    ...

2.Export

    ...
    $objEArray2Excel = new EArray2Excel('A title','Salvo Di Mare','A subject','A description','A category','ext.phpexcel.Classes.PHPExcel');
    //default: output to browser with Excel5 writer (xls extension)
    $objEArray2Excel->export($sheets,'my_users_and_cities');
    //output to browser with Excel5 writer (xls extension)
    $objEArray2Excel->export($sheets,'my_users_and_cities',true,'Excel5');
    //save to file with Excel5 writer (xls extension)
    $objEArray2Excel->export($sheets,'my_users_and_cities',false,'Excel5');
    //output to browser with Excel2007 writer (xlsx extension)
    $objEArray2Excel->export($sheets,'my_users_and_cities',true,'Excel2007');
    ...
