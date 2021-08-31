If you have to convert Access file (.accdb) to Excel (.xlsx) you may use this module.
IMPORTANT - This module may be used on Windows only (as Microsoft does provide Access drivers for Linux).

As input you have to point source Access file and get one Excel file in output. Target file is one to one to Access (contains all tables as sheets).

## USE AS MODULE

Module contains one class: AccessToExcel, instance is callable

To inicialize AccessToExcel instance you should pass args:
source_file=<accdb_file_path>
target_dir=<target_dir_path> 

As source_file you should provide path to Access file that has to be converted.
As target_dir you should provide path where Excel file should be loaded, may be same as source_file directory.

Exaple:
obj = AccessToExcel(
    source_file="C:\\foo\bar\file.accdb",
    target_dir ="C"\\path\to\target\dir"
)
obj()

## USE STANDALONE
You may use it via terminal as well:
>>> path_to_executing_file\run.exe source_file=path_to_access\file.accdb target_dir=dir_to_drop_excel