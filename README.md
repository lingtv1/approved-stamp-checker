The main goal of this code is to automatically scan the PDF files in a specified folder and check whether the first page of these files contains the "APPROVED" stamp. It will create an Excel file, which includes the names of each PDF file and whether it contains an approved stamp. In addition, if a PDF file does not have an approved stamp, then in the Excel file, the cell next to this file name will be marked in green for easy reference. This process greatly reduces the manual work of checking PDF files and improves the accuracy of checking. Please note that this code requires the user to have the poppler library installed on the computer and specify the path of the poppler library in the code. Moreover, this code only checks the first page of the file, if the word "APPROVED" appears on other pages, the code will not detect it.

这段代码的主要目标是自动扫描指定文件夹内的PDF文件，并检查这些文件的第一页是否包含"APPROVED"（已批准）的标记。它会创建一个Excel文件，其中包含每个PDF文件的名称以及它是否包含已批准的标记。此外，如果某个PDF文件没有已批准的标记，那么在Excel文件中，这个文件名旁边的单元格将被标记为绿色，以方便查阅。这个过程可以大大减少人工检查PDF文件的工作量，并且提高了检查的准确性
要注意的是，这段代码需要用户在电脑上安装有poppler库，并且在代码中指定poppler库的路径。而且，这段代码仅检查文件的第一页，如果“APPROVED”字样出现在其他页面，代码是检测不到的。
