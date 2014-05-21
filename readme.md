pyReconciliation-xla - A project to solve data reconciliation problem in Excel with Python
README.TXT
--------------------------------------------------------------------------------

Version 0.01 alpha - May 21, 2014
Copyright (c) 2014, Marcos Narciso

Introduction
------------

pyReconciliation-xla is a project to solve data reconciliation problem in Microsoft Excel with Python.

Pyinex project was used to embed Python with Excel. 

Requirements
------------

- Excel 2002, 2003, 2007 or 2010. I suppose it works well in Excel 2013, but I've not tested yet.

- Python 2.5, 2.6, or 3.1 (the python.org distribution) installed on your machine. Other distributions may work, but Pyinex hasn't been tested against anything else.

Pyinex is developed with Visual C++ 2008 Express Edition, available free of charge from Microsoft. You can download it from their website.

Examples
--------

In "test" folder has a example file.

XLW
---

Pyinex is written using the open-source XLW library (available at xlw.sourceforge.net), which greatly facilitates the production of XLLs. With the exception of one (presumed) bug fix to XLW 4.0's code, it is used as supplied by the authors. I have submitted the proposed fix to them for inclusion in their future releases.

License
-------

Pyinex is released under the modified BSD license; the license does not require the advertising clause present in the original BSD license. In particular, you can use Pyinex freely for commercial projects and products, and you do not need to expose or otherwise redistribute the resulting source code. You do need to include elements of the copyright notice, though - see the license comment at the top of all Pyinex files for details.