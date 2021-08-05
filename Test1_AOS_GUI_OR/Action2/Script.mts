
'Iterate through all the item categories as textname property of "CategoryTxt" object is parameterized and value read from data table column "Category"
Browser("Advantage Shopping").Page("Advantage Shopping").Link("CategoryTxt").Click
Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Item_Name").Click
Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Item_Name").Output CheckPoint("Item_Name") @@ script infofile_;_ZIP::ssf3.xml_;_
'Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Item_Price").Highlight @@ script infofile_;_ZIP::ssf5.xml_;_

Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Item_Price").Output CheckPoint("Item_Price") @@ script infofile_;_ZIP::ssf5.xml_;_

'Click on Home page link
Browser("Advantage Shopping").Page("Advantage Shopping").Link("HOME").Click

'dim oFSO
'' creating the file system object
'set oFSO = CreateObject ("Scripting.FileSystemObject")
'
'FilePath = (".\temp.csv")
'CreateFile (FilePath)
'
'
'
''Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Bose Soundlink Bluetooth").Click DataTable("Item_Name", dtLocalSheet)
'
'
' 
''Option Explicit
'' Create a new txt file
' 
'' Parameters:
'' FilePath - location of the file and its name
'' *********************************************************************************************
'Function CreateFile (FilePath)
'    ' varibale that will hold the new file object
'    dim NewFile
'    ' create the new text file
'    set NewFile = oFSO.CreateTextFile(FilePath, True)
'    set CreateFile = NewFile
' End Function
' 
'' *********************************************************************************************
'' Check if a specific file exist
' 
'' Parameters:
'' FilePath - location of the file and its name
'' *********************************************************************************************
'Function CheckFileExists (FilePath)
'    ' check if file exist
'    CheckFileExists = oFSO.FileExists(FilePath)
' End Function
' 
'' *********************************************************************************************
'' Write data to file
' 
'' Parameters:
'' FileRef - reference to the file
'' str - data to be written to the file
'' *********************************************************************************************
'Function WriteToFile (byref FileRef,str)
'   ' write str to the text file
'   FileRef.WriteLine(str)
'End Function
' 
'' *********************************************************************************************
'' Read line from file
' 
'' Parameters:
'' FileRef - reference to the file
'' *********************************************************************************************
'Function ReadLineFromFile (byref FileRef)
'    ' read line from text file
'    ReadLineFromFile = FileRef.ReadLine
'End Function
' 
'' *********************************************************************************************
'' Closes an open file.
'' Parameters:
'' FileRef - reference to the file
'' *********************************************************************************************
'Function CloseFile (byref FileRef)
'    FileRef.close
'End Function
' 

