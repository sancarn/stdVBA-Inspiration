# Iterate the ROT

I have written and tested this custom GetWorkbookLike function which is an API based alternative to the native VBA GetObject function.

This alternative function is supposed to work with Full workbook names as well as with Partial workbook names... It also accesses workbooks that are opened in seperate remote excel instances.

The function is passed part of the name of the seeked workbook but doesn't take wildcards


The tilde character (~) preceeding the exported "~TM..." workbook looks suspiciously like a file short path name so I have amended the API code to cater for that.

Also, I have changed the GetWorkbookLike function so that it takes in its argument part of the workbook name instead of taking part of the workbook full path name.


## Original thread

Author: Jaafar Tribak
Link:   https://www.mrexcel.com/board/threads/how-to-target-instances-of-excel.1118789/page-2#post-5395037

