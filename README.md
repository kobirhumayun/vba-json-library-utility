
# Getting Started

First make sure you have [git](https://git-scm.com/) installed on your system.

1.  import `Project.bas` into your project.
2.  Set a reference to `Microsoft Visual Basic For Applications Extensibility 5.3` and `Microsoft Scripting Runtime`
3.  In the immediate widow type `Project.InitializeProject`. This will create a gitignore file and run `git init`.
4.  This module is created to easily export and import VBA code to ./src directory. This gives the ability to use Git to version control your VBA.

This is needed as Git can't read Excel files directly, but can read the source files that are exported.
That's it! Version control can now be managed easily with VBA.

## How to maintain

1.  This template only contain project structure, not Excel Macro file.
    After clone this template create Excel Macro file in root directory & insert `Project.bas` file from src directory. Then export or import VBA module.
2.  Also on production copy not contain Excel Macro file in git repository, just create Excel Macro file in local directory when need, due to Excel file
    may arise problem on repository.

### Additional ref

