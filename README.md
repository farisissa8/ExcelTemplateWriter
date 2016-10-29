## Dependencies:
* [ZipArchive](http://php.net/manual/en/class.ziparchive.php)
* [DOMDocument](http://php.net/manual/en/class.domdocument.php)
* [Exception](http://php.net/manual/en/language.exceptions.php)
* [RecursiveDirectoryIterator](http://php.net/manual/en/class.recursivedirectoryiterator.php)
* [RecursiveIteratorIterator](http://php.net/manual/en/class.recursiveiteratoriterator.php)

## Public Methods:

#### __construct($source_path, $options):

* source_path: path to xlsx file
* options: unused (reserved for future features).

        $writer = new ExcelTemplateWriter('sheets/template.xlsx');

#### selectSheet($name):

* name: sheet name

        $writer->selectSheet("Stats");

#### writeToCell($column, $row, $data, $type):

* column: letter or position (starts at 1)
* row: position (starts at 1)
* data: string or number
* type: null (auto detect, default), "string", "number"

        $writer->writeToCell("B", 3, 200);
        $writer->writeToCell(2, 4, 201);

#### fillColumn($column, $starting_row, $data, $type):

* column: letter or position (starts at 1)
* starting_row: starting position (starts at 1). The row will be incremented to write all the data.
* data: string or number
* type: null (auto detect, default), "string", "number"

        $dates = array("1/10", "2/10", "3/10", "4/10", "5/10");
        $shares = array(16, 17, 18, 19, 20);
        $writer->fillColumn("E", 2, $dates);
        $writer->fillColumn("F", 2, $reach);

#### save($destination_path):

* destination_path: path for the new xlsx file (make sure the directory is writable)

        $writer->save("sheets/template_filled.xlsx");

## Supported Data Types:

* String (implemented as an inline string)
* Number
