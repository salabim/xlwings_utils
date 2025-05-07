### changelog | xlwings_utils

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
