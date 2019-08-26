<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

###############################################################################
# This script will copy Systems test Cases from an imported .csv file
# into the CARES Connect -Systems Test Cases- location of a specified project
###############################################################################

#region Imports 
Import-Module SharePointPnPPowerShellOnline -DisableNameChecking
Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
#endregion Imports

#region Variables
$CaresConectURL = "https://wigov.sharepoint.com/sites/dhs-cares/"
#$Username       = "test user name"
#$Password       = "test password"

$ErrorMsg       = New-Object system.Windows.Forms.Label
$Form           = New-Object system.Windows.Forms.Form
$FromLocation   = New-Object system.Windows.Forms.TextBox
$toLocation     = New-Object system.Windows.Forms.TextBox
$connectBtn     = New-Object system.Windows.Forms.Button
$OKbtn          = New-Object system.Windows.Forms.Button
$Cancel         = New-Object system.Windows.Forms.Button
$usernameTxtBox = New-Object system.Windows.Forms.TextBox
$passwordTxtBox = New-Object system.Windows.Forms.MaskedTextBox
$msgLbl         = New-Object system.Windows.Forms.Label

$requiredColumnsAry = "TESTCASENAME", "STEPS", "EXPECTEDRESULTS", "STATUS", 
                      "PRIORITY", "TESTCASETYPE", "TESTCASEGROUP", "SYSTEMTESTRELEASEDATE", "RELATEDREQUIREMENTS"
#endregion Variables

############################### BUTTON EVENTS START ################################ 

Function setupCancelClick(){
    $Cancel.Add_Click({
        $Form.DialogResult = 'ok'
    })
}

Function setupConnectClick(){

    ### Connect to CARES Connect on click of Connect button
    $connectBtn.Add_Click({
        
        try {
            
            # Validate that we have username and password
            if(!(validateCredentials)){
                $ErrorMsg.Text = "Username and Password are required"
                return
            }
            $msgLbl.Text = "Connecting..."

            # Set Credentials 
            [SecureString]$SecurePass = ConvertTo-SecureString $passwordTxtBox.Text -AsPlainText -Force 
            [System.Management.Automation.PSCredential]$PSCredentials = New-Object System.Management.Automation.PSCredential($usernameTxtBox.Text, $SecurePass) 
            
            # Connect to https://wigov.sharepoint.com/sites/dhs-cares/
            # on successful connection we will enable the to and from textboxes and the OK button
            Connect-PnPOnline -Url $CaresConectURL -Credentials $PSCredentials
            if (-not (Get-PnPContext)) {
                $ErrorMsg.Text = "Was not able to connect to CARES Connect."
                $msgLbl.Text = ""
                return 
            }else {
                $OKbtn.Enabled        = $true
                $toLocation.Enabled   = $true
                $FromLocation.Enabled = $true
                $msgLbl.Text = ""
                $ErrorMsg.Text = ""
            }
           
        } catch { 
            $ErrorMsg.Text = "Was not able to connect to CARES Connect."
            $msgLbl.Text = ""
            return 
        }
    })
}

### Get the From and To locations and copy CSV file to Sharepoint location
Function setupOKBtnClick(){

     $OKbtn.Add_Click({

        # Reset Error and message labels
        $ErrorMsg.Text = ""
        $msgLbl.Text = ""
        
        try {

            # Validate that we have a valid input file.
            if (!(validateInputFile $FromLocation)) {
                return
            }

            # Validate the CARES Connect URL and connect to the location
            if(validateSharepointURL){

                $msgLbl.Text = "Copying test cases... Please Wait"

                # Parse the URL and setup the connection.
                # We connect to the URL up until the /Lists... part
                $toURL = $toLocation.Text
                $caresConURL = $toURL.Substring(0, $toURL.IndexOf('/Lists'))

                # Set Credentials 
                [SecureString]$SecurePass = ConvertTo-SecureString $passwordTxtBox.Text -AsPlainText -Force 
                [System.Management.Automation.PSCredential]$PSCredentials = New-Object System.Management.Automation.PSCredential($usernameTxtBox.Text, $SecurePass) 
               
                Connect-PnPOnline -Url $caresConURL -Credentials $PSCredentials
                if (-not (Get-PnPContext)) {
                    $ErrorMsg.Text = "Was not able to establish connection to CARES Connect"
                    $msgLbl.Text = ""
                    return 
                }else {

                    # We have successfully connected to the location where the Systems Test Cases should be copied to,
                    # so continue to read in the input file and copy the values.
                    # Import the CSV file.
                    $fromContent = Import-CSV $FromLocation.text
                    [int]$LinesInFile = $fromContent.Count
            
                    # Validate columns of the imported file. 
                    # The input file columns must contain the correct columns
                    if(!(validateImportColumns $fromContent[0])) {
                        $ErrorMsg.Text = "Import file is missing required column(s)"
                        $msgLbl.Text = ""
                        return
                    }
                    
                    # Get the 'Systems Test Cases' list, and setup the columns we need to copy our data to.
                    $sysList = Get-PnPList "System Test Cases"
                    $testCasNamCol = (Get-PnPField -List $sysList -Identity "Test Case Name").InternalName
                    $stepsCol      = (Get-PnPField -List $sysList -Identity "Steps").InternalName
                    $expecRsltCol  = (Get-PnPField -List $sysList -Identity "Expected Results").InternalName
                    $statusCol     = (Get-PnPField -List $sysList -Identity "Status").InternalName
                    $priorityCol   = (Get-PnPField -List $sysList -Identity "Priority").InternalName
                    $testCasTypCol = (Get-PnPField -List $sysList -Identity "Test Case Type").InternalName
                    $testCasGrpCol = (Get-PnPField -List $sysList -Identity "Test Case Group").InternalName
                    $sysRelDtCol   = (Get-PnPField -List $sysList -Identity "System Test Release Date").InternalName
                    $relRequireCol = (Get-PnPField -List $sysList -Identity "Related Requirement(s)").InternalName
                    $currDate      = Get-Date -Format "MM/dd/yyyy"

                    # Loop trough our input file and copy all data to the CARES Connect Systems Test Cases list
                    $i = 0
                    foreach ($row in $fromContent ) {

                        # Default dates to the current date if empty
                        if($row.SystemTestReleaseDate.Length -eq 0){
                            $row.SystemTestReleaseDate = $currDate
                        }

                        

                        # TODO - Set Test Case Group like this. separate by comma
                        #$d = "Reports(IMMR) Layout", "Reports(IMMR) Functionality"
                        #Add-PnPListItem -List "Demo List" -Values @{"MultiUserField"="user1@domain.com","user2@domain.com"}
                        
                        
                        #Adds a new list item to the "Demo List" and sets the user field called MultiUserField to 2 users. 
                        #Separate multiple users with a comma.

                        # Add items to the list
                        Add-PnPListItem -List $sysList -Values @{$testCasNamCol = $row.TestCaseName;
                                                                $stepsCol      = $row.Steps;
                                                                $expecRsltCol  = $row.ExpectedResults;
                                                                $statusCol     = $row.Status
                                                                $priorityCol   = $row.Priority;
                                                                $testCasTypCol = $row.TestCaseType; 
                                                                $testCasGrpCol = $row.TestCaseGroup;
                                                                $sysRelDtCol   = $row.SystemTestReleaseDate;
                                                                $relRequireCol = $row.RelatedRequirements}
                            
                        $i += 1
                        # Display progress message to the user
                        $msgLbl.Text = "Copied " + $i + " of " + $LinesInFile + " ... Please Wait"
                    }
                }
            }else {
                $ErrorMsg.Text = "CARES Connect URL is not valid!"
                return
            }

        } catch { 
            $ErrorMsg.Text = "Lost connection to CARES Connect. Please verify CARES Connect URL."
            $msgLbl.Text = ""
            return 
        }
        
        # close the form
        $Form.DialogResult = 'ok'
        Disconnect-PnPOnline
        
    }) 
}#End OK button click

############################## BUTTON EVENTS END ##########################

############################## MAIN METHOD START ##########################
# Generate the input form to get excel file location and sharepoint CARES connect location
# copy the excel file as a Sharepoint list into the desired location.
function generateForm {

    # build form for user input
    buildDialog

    ### Close the form on Cancel
    setupCancelClick

    ### Connect to CARES Connect on click of Connect button
    setupConnectClick

    ### Connect to the CARES Connect Systems Test Cases location and copy 
    ### test cases form the input .CSV file
    setupOKBtnClick


    #Write your logic code here
    [void]$Form.ShowDialog()

} # generateForm
############################ MAIN METHOD END ################################

############################ VALIDATIONS START ##############################

Function validateCredentials(){

    # username and password is required
    if(($usernameTxtBox.Text.Length -eq 0) -or ($passwordTxtBox.Text.Length -eq 0)){
        return $false
    }
    return $true
}


Function validateSharepointURL() {

    # should be the correct host
    $sharepointURI = "https://" + ([System.Uri]$toLocation.Text).Host
    if($sharepointURI -ne "https://wigov.sharepoint.com") {
        return $false
    }

    # should include /Lists/System%20Test%20Cases
    if($toLocation.Text.IndexOf('/Lists/System%20Test%20Cases') -eq -1 ){
        return $false
    }
    return $true
}


# Validate that all require columns are included in the file
Function validateImportColumns($row1) {

    $columnNames = $row1 | get-member -MemberType NoteProperty | Select-Object Name
    foreach($columnName in $columnNames) {

        if(-not($requiredColumnsAry.Contains($columnName.Name.ToUpper()))) {
            return $false
        }
    }
    return $true
}

# Validate that the import file exist and is in .CSV format
Function validateInputFile($inputFile) {

    $IsValid = $true
    if($inputFile.text.Length -eq 0) {
        $ErrorMsg.Text = "Import file is Required"
        $IsValid = $false
    }elseif (-not ($inputFile.text.EndsWith('.CSV', 1))) {
        $ErrorMsg.Text = "Import File format not valid"
        $IsValid = $false
    }elseif (!(Test-Path $inputFile.text)) {
        $ErrorMsg.Text = "Import File not found"
        $IsValid = $false
    }

    return $IsValid
}

################################ VALIDATIONS END ##################################

# Build the dialog
Function buildDialog(){

    $Form.ClientSize                 = '400,450'
    $Form.text                       = "Import Systems Test Cases"
    $Form.TopMost                    = $false
    $Form.StartPosition              = "CenterScreen"

    $ErrorMsg.AutoSize               = $true
    $ErrorMsg.width                  = 25
    $ErrorMsg.height                 = 10
    $ErrorMsg.location               = New-Object System.Drawing.Point(25,15)
    $ErrorMsg.Font                   = 'Microsoft Sans Serif,10,style=Bold'
    $ErrorMsg.ForeColor              = "#d0021b"
    $ErrorMsg.MaximumSize            = New-Object System.Drawing.Size(340, 60);

    $Label1                          = New-Object system.Windows.Forms.Label
    $Label1.text                     = "CARES Connect user ID. (Include @dhs)"
    $Label1.AutoSize                 = $true
    $Label1.width                    = 100
    $Label1.height                   = 10
    $Label1.location                 = New-Object System.Drawing.Point(32,54)
    $Label1.Font                     = 'Microsoft Sans Serif,10'

    $usernameTxtBox.multiline       = $false
    $usernameTxtBox.width           = 332
    $usernameTxtBox.height          = 20
    $usernameTxtBox.location        = New-Object System.Drawing.Point(26,80)
    $usernameTxtBox.Font            = 'Microsoft Sans Serif,10'

    $Label2                          = New-Object system.Windows.Forms.Label
    $Label2.text                     = "CARES Connect password"
    $Label2.AutoSize                 = $true
    $Label2.width                    = 25
    $Label2.height                   = 10
    $Label2.location                 = New-Object System.Drawing.Point(32,126)
    $Label2.Font                     = 'Microsoft Sans Serif,10'

    $passwordTxtBox.PasswordChar     = '*'
    $passwordTxtBox.multiline       = $false
    $passwordTxtBox.width           = 332
    $passwordTxtBox.height          = 20
    $passwordTxtBox.location        = New-Object System.Drawing.Point(26,150)
    $passwordTxtBox.Font            = 'Microsoft Sans Serif,10'

    $connectBtn.text                 = "Connect"
    $connectBtn.width                = 70
    $connectBtn.height               = 30
    $connectBtn.location             = New-Object System.Drawing.Point(290,196)
    $connectBtn.Font                 = 'Microsoft Sans Serif,10'

    $Label3                          = New-Object system.Windows.Forms.Label
    $Label3.text                     = "Import file location (must be .CSV)"
    $Label3.AutoSize                 = $true
    $Label3.width                    = 25
    $Label3.height                   = 10
    $Label3.location                 = New-Object System.Drawing.Point(32,256)
    $Label3.Font                     = 'Microsoft Sans Serif,10'

    $FromLocation.multiline          = $false
    $FromLocation.width              = 332
    $FromLocation.height             = 20
    $FromLocation.location           = New-Object System.Drawing.Point(26,279)
    $FromLocation.Font               = 'Microsoft Sans Serif,10'
    #$FromLocation.Text               = 'C:\excelToSharepoint\systemsTestCases.csv'
    $FromLocation.Enabled            = $false

    $Label4                          = New-Object system.Windows.Forms.Label
    $Label4.text                     = "CARES Connect Systems test cases location"
    $Label4.AutoSize                 = $true
    $Label4.width                    = 25
    $Label4.height                   = 10
    $Label4.location                 = New-Object System.Drawing.Point(32,320)
    $Label4.Font                     = 'Microsoft Sans Serif,10'

    $toLocation.multiline            = $false
    $toLocation.width                = 332
    $toLocation.height               = 20
    $toLocation.location             = New-Object System.Drawing.Point(26,343)
    $toLocation.Font                 = 'Microsoft Sans Serif,10'
    #$toLocation.Text                 = 'https://wigov.sharepoint.com/sites/dhs-cares/TrainingDEV/Lists/System%20Test%20Cases/AllItems.aspx'
    $toLocation.Enabled              = $false

    $OKbtn.text                      = "OK"
    $OKbtn.width                     = 60
    $OKbtn.height                    = 30
    $OKbtn.location                  = New-Object System.Drawing.Point(299,386)
    $OKbtn.Font                      = 'Microsoft Sans Serif,10'
    $OKbtn.Enabled                   = $false

    $Cancel.text                     = "Cancel"
    $Cancel.width                    = 60
    $Cancel.height                   = 30
    $Cancel.location                 = New-Object System.Drawing.Point(230,386)
    $Cancel.Font                     = 'Microsoft Sans Serif,10'

   
    $msgLbl.text                     = ""
    $msgLbl.AutoSize                 = $true
    $msgLbl.width                    = 100
    $msgLbl.height                   = 10
    $msgLbl.ForeColor              = "#008000"
    $msgLbl.location                 = New-Object System.Drawing.Point(80,200)
    $msgLbl.Font                     = 'Microsoft Sans Serif,10'

    $Form.controls.AddRange(@($FromLocation,$Label1,$Label2,$toLocation,$OKbtn,$Cancel,$Label3,$usernameTxtBox,$Label4,$passwordTxtBox,$ErrorMsg,$connectBtn,$msgLbl))
    
}# End - build form


# Generate the form for user input
generateForm

