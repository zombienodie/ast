dim filesys, filetxt
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set filesys = CreateObject("Scripting.FileSystemObject")
Set filetxt = filesys.OpenTextFile("T:\Pr\VBS\FileNFolder\a.txt",8, True)
filetxt.WriteLine("cccc")
filetxt.Close