### changelog | xlwings_utils

[The full documentation can be found here.](https://www.salabim.org/xlwings_utils)

#### version 26.0.4  2026-01-18

  - added `undecorated` function. With this, a function can be undecorated. This can be very useful in xlwings Lite if you would like to call a function that is decorated with @xw.script. Example:
    ```
    @xw.script
    def Task_1(book: xw.Book):
        ...
    
    @xw.script
    def Task_2(book: xw.Book):
        ...
        
    @xw.script
    def Task_1_2(book: xw.Book):
        xwu.undecorated(Task_1)(book)
        xwu.undecorated(Task_2)(book)    
    ```
    
#### version 26.0.3  2026-01-15

  - the method `import_from_folder` is now part of a cloud module itself, so for instance `xwu.dropbox.import_from_folder`

  - xwu.github is a new cloud module, that allows accessing the complete GitHub-space. As of now, writing and deleting files is not supported.

    It is strongly recommended to obtain a (free) GitHub token and set the environment variable GITHUB.TOKEN to that value.
    `xwu.github.dir()`, `xwu.github.read()` and `xwu.github.import_from_folder` require a tuple of the owner,  repo and path to be given.
    Specific to this cloud module, the method `xwu.github.get_repos` is implemented to get the names of all repos of a given owner.
    
    Note that, if you don't specify a token, GitHub allows only 60 requests per hour, whereas with a token 5000 requests per hour are allowed. A warning will be printed in that case.

#### version 26.0.2  2026-01-12

  - the optional parameter `module_name` of `import_from_folder` has been removed, as it doesn't serve any purpose.

#### version 26.0.1  2026-01-11

  - This version introduces `import_from_folder`, which allows importing a module from a folder in a cloud service (just dropbox for now), rather than from PyPI via `requirements.txt`

    This can be very useful in case:
    - the module is not (yet) available under PyPI

    - the module is being developed and the developer wants to debug the module prior to uploading to PyPI

    - the module is just uploaded to PyPI, but it is not yet visible in pyodide.

    Example:
    ```
    import xlwings_utils as xwu
    vardict = xwu.import_from_folder(cloud=xwu.dropbox, folder_name="/Python/vardict/vardict")
    # vardict will now be imported from dropbox
    ```

    Note that if a module with the same name is already imported, calling `xwu.import_from_folder` has no effect.

#### version 26.0.0  2026-01-03

  - `block.highest_used_row_number` and `block.highest_used_column_number` did ignore cells containing the value 0. Fixed,

#### version 25.2.2  2025-12-21
  - Both `dropbox.read()` and `nextcloud.read()` now cache the contents by default. The cache parameter may be used to switch off caching.

#### version 25.2.1  2025-12-20
  - Bug in the calculation of number of seconds of @timer fixed.

#### version 25.2.0  2025-12-15

  - Overhaul of accessing cloud storage. Not only Dropbox, but also Nextcloud is now supported (although not working under xlwings Lite (yet?)).
    The interface is unified and goes via a submodule, i.e. `xlwings_utils.dropbox` or `xlwings_utils.nextcloud`.
    
    -  `init_dropbox` is now `dropbox.init`
    -  `read_dropbox` is now `dropbox.read`
    -  `write_dropbox` is now `dropbox.write`
    -  `list_dropbox` is now `dropbox.dir`
    -  `delete_from_dropbox` is now `dropbox.delete`
    
    The same functionality is provided for Nextcloud. For Nextcloud to work, it is required to provide url, username and password. This can be done via `nextcloud.init` or -preferably- via the environment variables `NEXTCLOUD.URL`, `NEXTCLOUD.USERNAME` and `NEXTCLOUD.PASSWORD`.
    
  - Overhaul of accessing local storage.
    The interface is now more or less the same as the cloud access interface.
    `read_local` is now `local.read`
    `write_local` is now `local.write`
    `list_local` is now `local.dir`


 #### version 25.1.3  2025-09-22

  - Added the timer decorator, which should be placed after the @xw.script decorator.
    Upon completion, it will then show the name, entry time, exit time and the duration, like

    `Done MyScript  11:51:13.24 - 11:51:20.28 (7.04s)`

#### version 25.1.2  2025-09-10

  - Some internal changes and doc string updates

  #### version 25.1.1  2025-09-09

  - Added `requests` to dependencies.

#### version 25.1.0  2025-09-09

  - Dropbox file handling does not use the dropbox module anymore
  - File reading from dropbox is now reliable ... at last
  - Introduced `delete_from_dropbox()` function.
  - ssl is not required in  requirements.txt anymore

#### version 25.0.10  2025-09-03

  - `block.highest_used_row_number` and `block.highest_used_column_number` now uses caching, if possible.
  - tests for this caching have been added.

#### version 25.0.9  2025-08-23

  - Special support for Dropbox credentials under Pythonista is deprecated as environment variables can be set upon start of Pythonista.

#### version 25.0.8  2025-08-22

  - Dropbox credentials (environment variables) are now called `DROPBOX.APP_KEY`, `DROPBOX.APP_SECRET` and `DROPBOX.REFRESH_TOKEN`.
  - under Pythonista, environment info should now be in `environ.toml`, rather than `.environ.toml`
  - block.transpose()` is now called `block.transposed()`
  - optimised setting None values in a block (now deleted from the dict)
  - equality test added (based on value)
  - updated tests

#### version 25.0.7  2025-08-05

  - `block` does not accept a value anymore; use the new `block.from_value()` or use the new -arguably more useful- `block.from_range()`.
  - `block.value` can't be used to set a value anymore (see above)
  - `block.encode_files` is now called `block.encode_file` and can encode only one file at a time
  - the class methods `from_xxx` do not support number_of_rows and number_of_columns parameters anymore; instead use the new `reshape` method 
  - the repo now contains the file xlwings_utils.bas, which can be used as a module to encode and decode encoded files on a sheet.

#### version 25.0.6  2025-08-01

- #### File transfers via VBA are now done with:

  - `block.encode_files(*files)`
  * `block.decode_files()`
  * `trigger_macro()`

- `read_dropbox()` now retries automatically (by default 100 times)


#### version 25.0.5  2025-07-21

- added:
  * `trigger_VBA()`
  * `init_transfer_files()`
  * `transfer_file()`

#### version 25.0.4  2025-07-14

- ssl is no longer in the requirements section of pyproject.toml as that caused problems with pip install on regular installations. As a consequence, ssl has to be added to the requirement.txt xlwings lite panel.
- dropbox_init can now propagate keyword parameters to dropbox.dropbox, to facilitate experimenting.

#### version 25.0.3  2025-06-08
- updated tests
- now lookup, hlookup and vlookup have an optional default parameter. If used, default will be returned if
  the value is not found. Otherwise, a ValueError exception will be raised.
- now lookup_row and lookup_column have an optional default parameter. If used, default will be returned if
  the row or columnn cannot be found. Otherwise, a ValueError will be raised.
- internal change: optional parameters are now handled with missing rather than None.
- a script to set up dropbox is now available under https://www.salabim.org/dropbox%20setup.py. Also, instructions on how to run that file are added to the readme file.

#### version 25.0.2  2025-06-07
- sometimes reading from Dropbox with xwu.read_dropbox does not return the right contents.
  In that case, now an OSError exception is raised. This makes it possible to retry then.
  The message will show both the actual and expected file size.

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
