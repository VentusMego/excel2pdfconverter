# excel2pdfconverter
convert excel to pdf 【xls与xlsx批量转化为pdf的东西】
原理：通过pywin32模块调用Windows API，批量导出同一文件夹下的.xls/.xlsx文件的pdf版本。可以选择全工作簿导出或者是仅当前页导出。界面简陋还需要进一步修改。
注意：需要先安装pywin32，按照提示进行操作。目前相同文件夹下除了表格文件不建议放其它文件，以免程序崩溃。

# doc converter
另一个文件是doc批量转pdf文件。代码还需要优化，目前是通过导出完成后杀进程的方式来控制内存。代码仍需后期优化。
