# excelToSharepoint
Import Systems test cases into CARES connect from excel/csv file


# Header requirements for the Excel file
The Excel/csv document must have these coloums with exact names(no spaces, but not case sentitive), and not contain other columns
1.	TestCaseName
2.	ExpectedResults
3.	Steps
4.	Status
5.	Priority
6.	TestCaseType
7.	TestCaseGroup
8.	SystemTestReleaseDate
9.	RelatedRequirements



# Steps to copy cases
2.	Save the Excel document as a .csv file somewhere on your local drive.
3.  Open the Dialog to import the Systems Test cases into CARES Connect.
3.	Enter your CARES Connect username and password. (same as email credentials) include @dhs.
4.	Click Connect
5.  Copy the path of your local import file into 'Import File Location' textbox
6.  Copy the CARES Connect location of where the systems test cases should be copied to. This must be in the ‘Systems Test Cases’ location     of the project you are working on.
7.  Click OK.
8.  Copying the test cases will take some time.

# Set Desktop short cut
1. Right click and select 'create new short cut'
2. copyt this into the dialog textbox. Make sure your file path matches your local filepath
    C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -ExecutionPolicy Bypass -windowstyle Hidden -File        C:\excelToSharepoint\powershellScript\excelToSharepointList.ps1
3. Enter the name and click Finish.
    
