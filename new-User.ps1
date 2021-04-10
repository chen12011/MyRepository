Add-Type -AssemblyName presentationframework
add-type -assemblyName system.windows.forms

<#

.SYNOPSIS

    GUI Script that creates a new user based off of the new-user Excel form submitted by HR through Excel. 
    
.Description

    This script pulls data from specified fields in an Excel file and requires data to be in the right fields to work

.Notes

#>


[xml]$xaml = @"
<Window x:Name="MainWindow1" 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Add Active Directory User 1.0" Height="437.09" Width="805.533">
    <Grid>
        <Button x:Name="OpenFile" Content="Open File" HorizontalAlignment="Left" Height="22" Margin="10,10,0,0" VerticalAlignment="Top" Width="69" RenderTransformOrigin="1.45,0.711" ClickMode="Press"/>
        <Label x:Name="LogLabel" Content="Logs" HorizontalAlignment="Left" Height="26" Margin="537,25,0,0" VerticalAlignment="Top" Width="97" FontSize="11"/>
        <ListView x:Name="logView" HorizontalAlignment="Left" Height="313" Margin="421,71,0,0" VerticalAlignment="Top" Width="367" ItemsSource="{Binding Path=ReportEntries}"  
          VerticalContentAlignment="Top"  
          ScrollViewer.VerticalScrollBarVisibility="Visible"
          ScrollViewer.CanContentScroll="False">
            <ListView.View>
                <GridView>
                    <GridViewColumn/>
                </GridView>
            </ListView.View>
        </ListView>
        <Label Content="Name" HorizontalAlignment="Left" Height="26" Margin="10,130,0,0" VerticalAlignment="Top" Width="76"/>
        
        <Label Content="Username" HorizontalAlignment="Left" Height="26" Margin="10,170,0,0" VerticalAlignment="Top" Width="76"/>
        <Label Content="Title" HorizontalAlignment="Left" Height="26" Margin="10,210,0,0" VerticalAlignment="Top" Width="76"/>
        <TextBox x:Name="firstNameTextBox" HorizontalAlignment="Left" Height="26" Margin="100,130,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256" IsReadOnly="True" />
        
        <TextBox x:Name="usernameTextBox" HorizontalAlignment="Left" Height="26" Margin="100,170,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256" IsReadOnly="True" />
        <TextBox x:Name="titleTextBox" HorizontalAlignment="Left" Height="26" Margin="100,210,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256" IsReadOnly="True" />
        <Label Content="Department" HorizontalAlignment="Left" Height="26" Margin="10,250,0,0" VerticalAlignment="Top" Width="76"/>
        <Label Content="Supervisor" HorizontalAlignment="Left" Height="26" Margin="10,290,0,0" VerticalAlignment="Top" Width="76"/>
        <TextBox x:Name="departmentTextBox" HorizontalAlignment="Left" Height="26" Margin="100,250,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256" IsReadOnly="True" />
        <TextBox x:Name="supervisorTextBox" HorizontalAlignment="Left" Height="26" Margin="100,290,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256" IsReadOnly="True" />
        <Label Content="Email" HorizontalAlignment="Left" Height="26" Margin="10,330,0,0" VerticalAlignment="Top" Width="76"/>
        <TextBox x:Name="emailTextBox" HorizontalAlignment="Left" Height="26" Margin="100,330,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256" IsReadOnly="True" />
        <Button x:Name="createUser" Content="Create User" HorizontalAlignment="Left" Height="22" Margin="10,42,0,0" VerticalAlignment="Top" Width="69" RenderTransformOrigin="1.45,0.711"/>
        <TextBox x:Name="fileTextBox" HorizontalAlignment="Left" Height="54" Margin="100,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256" IsReadOnly="True" ScrollViewer.CanContentScroll="True" VerticalScrollBarVisibility="Visible" />
        <Button x:Name="createSetupForm" Content="Create Setup Form" HorizontalAlignment="Left" Height="22" Margin="10,74,0,0" VerticalAlignment="Top" Width="105" RenderTransformOrigin="1.45,0.711"/>

    </Grid>
</Window>
"@


#initializing gui
$reader = (new-object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

#initializing gui components
$firstNameTextBox = $window.FindName("firstNameTextbox")
$openFileButton = $window.FindName("OpenFile")
$createUserButton = $window.FindName("createUser")
$fileTextBox = $window.FindName("fileTextBox")
$firstnameTextBox = $window.FindName("firstNameTextBox")
$lastNameTextBox = $window.FindName("lastNameTextBox")
$usernameTextBox = $window.FindName("usernameTextBox")
$titleTextBox = $window.FindName("titleTextBox")
$departmentTextBox = $window.FindName("departmentTextBox")
$supervisorTextBox = $window.FindName("supervisorTextBox")
$emailTextBox = $window.FindName("emailTextBox")
$listview = $window.FindName("logView")
$setupFormButton = $window.FindName("createSetupForm")

function get-ExcelData(){

    <#
    .SYNOPSIS
        Grabs relevent new hire data from the User Setup Form provided by HR
    .Description
        This function will open an instance of the provided Excel file selected within File Explorer.
        It will then go to the predefined cell numbers and grab the relevent new user data such as 
        name, supervisor, department, etc. Once completed it will display the information within the 
        for reference.
    #>


    [cmdletbinding()]
    
    param (
        #The name of the user setup form
        [Parameter(Mandatory=$true)]
        [string]$excelFilename
    )

    #Clear textbox contents
    $firstNameTextBox.clear()
    $usernameTextBox.clear()
    $titleTextBox.clear()
    $departmentTextBox.clear()
    $supervisorTextBox.clear()
    $emailTextBox.clear()

    #Creates a new excel object and opens the selected excel file from the openfiledialog function
    $excel = new-object -ComObject excel.Application
    $workbook = $excel.Workbooks.Open($excelFilename)
  

    #opening worksheet #1
    try{
    
        $Global:sheet = $workbook.Sheets.Item('Sheet1')

    }
    catch{

        $error = $_.exception.message 
        $listview.addtext($error)
        $listview.addtext("File could not be opened, please try opening an Excel file")
        $workbook.Close()
        return

    }

    
    #grabbing relevant information from the user setup form
    #checking to make sure the string is not empty otherwise the form might not be aligned properly with the cells
    $global:employeeName = ($sheet.Range('D4')).text

    if([string]::IsNullorEmpty($employeeName)){

        $listview.AddText("employee Field is empty please check the form and make sure it is setup properly.")
        return

    }

    $global:department = ($sheet.Range('D6')).text
    if([string]::IsNullorEmpty($department)){

        $listview.AddText("employee Field is empty please check the form and make sure it is setup properly.")
        return

    }

    $global:title = ($sheet.Range('I4')).text
    if([string]::IsNullorEmpty($title)){

        $listview.AddText("employee Field is empty please check the form and make sure it is setup properly.")
        return

    }

    $global:supervisor = ($sheet.Range('I8')).text
    if([string]::IsNullorEmpty($supervisor)){

        $listview.AddText("employee Field is empty please check the form and make sure it is setup properly.")
        return

    }


    #initializes the domain variable
    $domain = "@WorkDomain"
    $listview.AddText("Getting Info!")

    #testing to see if supervisor is spelled correctly or found in active
    $arraySupervisor = $supervisor.Split(" ")

    #if the supervisor variable contains more than one word, converts the words into 1 word.
    if ($arraySupervisor.Count -gt 1 ) {

        $Global:supervisor = $arraySupervisor[0] + (($arraySupervisor[1])[0])

    }
    try {

        $variable = Get-ADUser -Identity $supervisor

    }
    catch {

        $error = $_.exception.message
        $listview.AddText($error)
        $workbook.Close()
        return

    }


    #setting up Password
    $password = "Password1231"
    $adPass = ConvertTo-SecureString -String $password -AsPlainText -Force

    #path to the user OU
    $adPath = "Organizational Unit"

    #Creating the AD username by parsing the fullname and grabbing first name and last initial
    $nameArray = $employeeName -split "\s+"
    $global:adUsername = $nameArray[0] + (($nameArray[1])[0])

    #Creating the email address including the domain name
    $userPrincipalName = $adUsername + $domain

    #applying fields to GUI
    $firstNameTextBox.AddText($employeeName)
    $usernameTextBox.AddText($userPrincipalName)
    $titleTextBox.AddText($title)
    $departmentTextBox.AddText($department)
    $supervisorTextBox.addtext($supervisor)

    #adding email information
  
    $global:email = $adUsername + "@domain.com"
    $emailTextBox.AddText($email)


}


function add-user()
{

    <#
    .SYNOPSIS
        Adds a new Active Directory user based on the data retrieved from the
        Get-ExcelData function.
    .DESCRIPTION
        Utilizes the data retreieved from the Get-ExcelData function and converts
        specific variables into data that can be inputted into the New-User Cmdlet. 
        Once user has been created, it will copy over the Active Directory groups
        from the Copy User field in Excel. If the field is empty or not found, it will
        load a default list of groups that all users are default a member of.
    #>


  [cmdletbinding()]


    #setting up Password
    $password = "Password1234"
    $adPass = ConvertTo-SecureString -String $password -AsPlainText -Force

    #path to the User OU
    $adPath = "Organizational Unit"

    #Creating the AD username by parsing the fullname and grabbing first name and last initial
    $nameArray = $employeeName -split "\s+"
    $adUsername = $nameArray[0] + (($nameArray[1])[0])

    #Creating the email address including the domain name
    $userPrincipalName = $adUsername + $domain

    #grabbing admin credentials
    $credential = Get-Credential 

    #command to create user, using the data pulled from above
    $newUserParams = @{

        Name = $employeeName 
        GivenName = $nameArray[0]
        Surname = $nameArray[1] 
        DisplayName = $adUsername 
        email = $adUsername + "@domain.com"
        SamAccountName = $adUsername 
        UserPrincipalName = $userPrincipalName 
        Path = $adPath 
        AccountPassword = $adPass 
        enabled = $true 
        description = $title 
        Office = "Office Location" 
        Manager = $supervisor 
        Department = $supervisor.Department 

    } 
    #creating new ad user 
    try {

        New-ADUser @newUserParams -Credential $credential 
    
    }
    catch {

        $Errormessage = $_.exception.message
        $listview.addText($Errormessage)
        return

    }

    $listview.addText("$adUsername has been created")



    #storing new user in variable $adUser
    $adUser = get-aduser $adUsername 
    Set-ADUser -Identity $adUsername -Title $title -Company "Company Name" -Credential $credential 

    #grabbing profile to copy from
    $copyUser = ($sheet.Range('F36')).text 

    #testing to see if the copy user is found within Active Directory 
    $listview.addText("checking if copy profile user exists")
    try {

        $adCopy = get-aduser $copyUser

    }
    catch {
    
    #Checking to see if AD username can be retrieved using the fullname
    $adCopy = get-aduser -Filter{Name -eq $copyUser} 
    if ($null -eq $adCopy) {

        $listview.AddText("Copy profile not found within Active Directory, Will add standard groups instead")

    }
  
} 


    #grabbing a list of users within company to compare with the profile to copy from
    $userList = get-aduser -Filter * -SearchBase "Organization Unit"

    #checking to see if username is contained within Organizational Unit
    if($userList.name -contains $adCopy.Name )
    {
        $groupList = Get-ADPrincipalGroupMembership -Identity $adcopy.DistinguishedName 
        foreach($group in $groupList){

            Add-ADPrincipalGroupMembership -identity $adUser.distinguishedName -MemberOf $group.name -Credential $credential 
            $listview.AddText("$(group.name) has been added from $copyUser")
        
        }  
    }
#add a standard list of security groups if the "copy user" cannot be found in Active Directory
    else {
        Write-Output "Now adding standard array of groups"
        $standardList = @(
           "List of Active Directory Groups"
            )
    foreach($groupName in $standardList){

        $standardGroup = Get-ADGroup -Filter{name -eq $groupname}
        Add-ADPrincipalGroupMembership -Identity $adUser -MemberOf $standardGroup.SamAccountName -Credential $credential  

    }

}
    $listview.AddText("Groups have been added")
    
    $workbook.Close()
}


function set-SetupForm()
{
    <#
    .SYNOPSIS
        Creates a new user welcome form in Microsoft Word
    .DESCRIPTION
        Creates a new user setup form and fills in the relevent fields with the new hire 
        username. Drops the file within server share folder.
    #>

    [cmdletbinding()]
   
    #word guide setup
    $wordfilename = "\\server\share\location"
    $word = new-object -comobject word.Application
    $wordfile = $word.documents.Open($wordfilename)
    $wordSelection = $word.Selection

    #script will replace every instance of the following strings with the $replacewith variable.
    $findText = "DOMAIN_NAME"

    $MatchCase = $false
    $MatchWholeWord = $true
    $MatchWildCards = $false
    $MatchSoundsLike = $false
    $MatchallWordForms = $false
    $Forward = $true
    $ReplaceWith = $adUsername


    $Replace = 2
    $format = $false
    $wrap = 1
    Write-Output "editing file now"
    try{

        $wordSelection.find.Execute($findText,$MatchCase,$MatchWholeWord,$MatchWildCards,$MatchSoundsLike,$MatchallWordForms,$Forward,$wrap,$format,$ReplaceWith,$Replace)
    }
    catch{
        $Errormessage = $_.exception.message
        $listview.addText($Errormessage)
        return
    }
    $listview.AddText("User setup form has been created!")
    $listview.AddText("File can be found here: \\Server\share\location\$employeename ")
    $wordfile.SaveAs("\\server\share\location\$employeename")

    $wordfile.Save()
    $wordfile.close()

}



function openFileDialog(){
  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $excelFilename = $OpenFileDialog.FileName
    $filetextbox.addText($OpenFileDialog.filename)

    get-ExcelData $excelFilename


}

$openFileButton.add_Click({openFileDialog})
$createUserButton.add_Click({add-user})
$setupFormButton.add_Click({set-SetupForm})
 
$window.ShowDialog() | out-null










