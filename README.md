# Cover-Letter-Autofill
Are you tired of typing the cover letter every time you apply for another job? This VBA file might be of some help!!!

## Objective

In this project, I built a VBA file which autofill the important information of the cover letter and export the cover letter as a PDF file which is proven to be more error-proofing than a Word file if the employers want to view your cover letter.
The main objective is to remove boring, repetitive tasks while some is preparing for the job application. It is designed to be as user-friendly as possible, there are user guides in the form of excel notes all over the main sheet.
Tools of choice
Excel and Word is used in this project due to its availability and commonness. Those apps will be used for customizing the cover letter to your specific needs before you run the code. Some basic Excel formulas will also be utilized to support the VBA code.
VBA is the main programming language in this project. All the main code is written inside the VBA Excel file which connects directly to the Word file containing the sample framework of the cover letter.

## User Guides

### 1. Modify static information and create bookmarks in the Word file
   
In order for the VBA code know where to place the information,  bookmarks have to be created in the original cover letter Word template. However, there is already a list of bookmarks added by default. If there is any change, remember to re-bookmark that spot otherwise the VBA code will not run as expected. Besides, the static information (the ones with no square brackets) could be modified if the default writing style or some contents of the letters didn't fit.

### 2.  Input the information in the Excel file
   
Fill in all the information in the "Input" section of the "Input" sheet, and also make sure the name of your Word Template in cell "D5" is correct so the code can find the targeted Word file.
Guides in the "Note" section on the right side needs to be followed for the better outcome. There are also some notes hidden in a few cells (red-marked ones), please read them carefully so as to get the expected outcome of the cover letter.
Remember NOT to change the name of any field unless the code will not perform as expected. If changes are made, go to the "Prep" sheet to make the same changes in the "Respective Field Name" as well; Otherwise, the outcome will not be as expected.

In the "Prep" sheet, the existing values in the "Default Value" column is triggered only when users input no information into that specific field.
If "Get Bookmarks" button in the "Prep" sheet is clicked, the order of rows in the "Respective Field Name" column will need to be manually adjusted according to the "Bookmark Name" column.


### 3. Run the Excel file

After filling in all the information required and assuring that the Word file and Excel file are in the same folder, click the "Export PDF Button" to export the modified Cover Letter in PDF format. 
The new PDF file will be in the same folder with the Excel File, and the name of the file is hardcoded as "Coverletter_Candidate Name_Job Title_Company Name_ENG.pdf".
If the name of the file already exists,  the new file will replace the old file with the old name. There is also chance that the name of the new file will be added a number at the end so it could tell differences with the old one's.

## Code Breakdown 
### 1. "Info_Autofill" Sub
   
#### a. Purpose

Loop through all the bookmarks and replace them with respective information in the Excel file, eventually export the complete Word file as PDF to the file path.

#### b. Details
   
Define all the required variables
   
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim ExistPDF As Workbook
    Dim BookRange As Word.Range, BookName As String
    Dim r As Byte
    Dim v As Byte
    Dim FileName As String
    Dim Answer As VbMsgBoxResult

Assign the "WordApp" variable to a Word application object to control Word from Excel

    On Error Resume Next
    Set WordApp = GetObject(, "word.application")
    
    On Error GoTo 0
    If WordApp Is Nothing Then
        Set WordApp = New Word.Application
    End If
    WordApp.Visible = True

Assign the 'WordDoc' variable to the specific Word file so we can interact with it
   
     Set WordDoc = WordApp.Documents.Open(ThisWorkbook.Path & "\" & ShInp.Range("Template_Name").Value)
   
Create a loop to replace bookmarks in the Word file with the respective information from Excel
     
    r = 2
    Do Until ShMap.Cells(r, 1).Value = ""
        BookName = ShMap.Cells(r, 1).Value
        Set BookRange = WordDoc.Bookmarks(BookName).Range
        If ShMap.Cells(r, 3).Text = "" And ShMap.Cells(r, 4).Text <> "" Then
            NewText = ShMap.Cells(r, 4).Text
        ElseIf ShMap.Cells(r, 3).Text = "" And ShMap.Cells(r, 4).Text = "" Then
            MsgBox "You leave the field " & ShMap.Cells(r, 2).Text & " empty, please try again"
            GoTo close_app
            Exit Sub
        Else
            NewText = ShMap.Cells(r, 3).Text
        End If
        
        BookRange.Text = NewText
        WordDoc.Bookmarks.Add BookName, BookRange
        r = r + 1
    Loop
        
Create a file path and export the file as PDF to that file path , add error-handling code to prevent duplications in the exported file name

   
    FileName = ThisWorkbook.Path & "\Cover Letter_PHAM DUC TOAN_" & ShMap.Cells(13, 3).Text & "_" & ShMap.Cells(4, 3).Text & "_ENG.pdf"
    On Error GoTo handling
    WordDoc.ExportAsFixedFormat ExportFormat:=wdExportFormatPDF, OutputFileName:=FileName, OpenAfterExport:=True
    GoTo close_app
    
    
    handling:
    v = v + 1
    FileName = ThisWorkbook.Path & "\Cover Letter_PHAM DUC TOAN_" & ShMap.Cells(13, 3).Text & "_" & ShMap.Cells(4, 3).Text & "_ENG (" & v & ").pdf"
    WordDoc.ExportAsFixedFormat ExportFormat:=wdExportFormatPDF, OutputFileName:=FileName
    


Close the Word file without saving, set the 'WordApp variable' to None, copy the full text of the cover letter to CLIPBOARD, allow users to choose if they want to open the PDF file.
    
      close_app:
      Set BookRange = WordDoc.Bookmarks("Cover_Letter").Range
      CreateObject("htmlfile").ParentWindow.ClipboardData.SetData "text", BookRange.Text
      
      WordDoc.Close False
      Set WordApp = Nothing
      
      Answer = MsgBox("The PDF file was COPIED to clipboard and was EXPORTED successfully in the following directory: " & FileName & vbNewLine & "Do you want to open the file?", vbYesNo, "Exported Successfully")
      
      If Answer = vbYes Then ActiveWorkbook.FollowHyperlink FileName
    

### 2. "get_bookmark_name" Sub

#### a. Purpose

Find all the bookmarks users have been created in the Word file and print them to the Excel sheet for name reference when the "info_autofill" sub is running.

#### b. Details

Define all the required variables
 

    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim BookMark As Word.BookMark
    Dim r As Byte
    Dim Answer as vbMsgBoxResult
   
Assign the "WordApp" variable to a Word application object to control Word from Excel
  
    On Error Resume Next
    Set WordApp = GetObject(, "word.application")
    On Error GoTo 0
    If WordApp Is Nothing Then
        Set WordApp = New Word.Application
    End If
 
Add a Message Box to stop the code if needed
   
    Answer = MsgBox("Do you want to overwrite the existing bookmarks, this might cause indexing errors in VBA", vbYesNo, "Are you sure?")
    If Answer = vbNo Then GoTo close_app
 
Assign the 'WordDoc' variable to the specific Word file so we can interact with its contents
   
    Set WordDoc = WordApp.Documents.Open(ThisWorkbook.Path & "\" & ShInp.Range("Template_Name").Value)
Loop through the Bookmarks collection of that Word file and print the BookMark name in the Mapping sheet
       
    For Each BookMark In WordDoc.Bookmarks
        ShMap.Cells(r, 1).Value = BookMark.Name
        r = r + 1
    Next
       
Close the Word file without saving, set "WordApp" varibale to None
   
    
    WordDoc.Close False
    Set WordApp = Nothing
    
