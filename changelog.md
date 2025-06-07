### changelog | xlwings_utils

[The full documentation can be found here.](https://www.salabim.org/xlwings_utils)

#### version 25.0.2  2025-06-07
- sometimes reading from Dropbox with xwu.read_dropbox does not return the right contents.
  In that case, now an OSError exception is raised. This makes it possible to retry then.

#### version 25.0.1  2025-05-26
- now uses *calver* versioning
- complete overhaul of capture_stdout, which is now to be initiated with `capture = xwu.Capture()`
- block has now methods to lookup values in a block, which is very convenient for project / scenario sheets:
  - `lookup()` / `vlookup()`
  - `hlookup()`
  - `lookup_row()`
  - `lookup_column()`
- block now has some additional class methods:
  - `from_xlrd_sheet()`
  - `from_openpyxl_sheet()`
  - `from_dataframe()`
  -  from_file()

- block now has a method to write to an openpyxl sheet: `to_openpyxl_sheet`

- an application can now check `xwu.xlwings` to see whether it runs in an xlwings environment.

#### version 0.0.7  2025-05-07

- maximum_row is renamed to highest_used_row_number
- maximum_column is renamed to highest_used_column_number
- block can now also be initialised with a block, which can be helpful to redimension a block
- more tests
- documentation improvements

#### version 0.0.6  2025-05-06

- under Pythonista, Dropbox credentials will now be retrieved from .environ.toml file
- list_dropbox and list_local can now optionally show folders

#### version 0.0.5  2025-05-05

- no need to call dropbox_init anymore, if REFRESH_TOKEN, APP_KEY and APP_SECRET
  are specified as environment variables. 

#### version 0.0.4  2025-05-04

- list_pyodide, read_pyodide and write_pyodide renamed to list_local, read_local, write_local

- added block functionality

- added test script

#### version 0.0.0  2025-04-20

- Initial version
