Add-Type -AssemblyName System.Windows.Forms
$Icon = New-Object system.drawing.icon (".\Files\Logo.ico")
$Font = New-Object System.Drawing.Font("Times New Roman",9)

$TANYRHealthcare = New-Object system.Windows.Forms.Form
$TANYRHealthcare.Text = "TANYR"
$TANYRHealthcare.TopMost = $false
$TANYRHealthcare.Width = 180
$TANYRHealthcare.Height = 245
$TANYRHealthcare.Icon = $Icon
$TANYRHealthcare.Font = $Font
$TANYRHealthcare.FormBorderStyle = 'FixedDialog'

##----------------------------MAIN PAGE GUI----------------------------##
$Button_Task1 = New-Object system.windows.Forms.Button
$Button_Task1.Text = "Select Sheet"
$Button_Task1.Width = 126
$Button_Task1.Height = 35
$Button_Task1.location = new-object system.drawing.point(17,13)
$Button_Task1.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task1)
$TANYRHealthcare.AcceptButton = $Button_Task1

$Button_Task3 = New-Object system.windows.Forms.Button
$Button_Task3.Text = "Manage Rules"
$Button_Task3.Width = 126
$Button_Task3.Height = 35
$Button_Task3.location = new-object system.drawing.point(17,60)
$Button_Task3.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task3)

$Button_Task5 = New-Object system.windows.Forms.Button
$Button_Task5.Text = "Apply Rules"
$Button_Task5.Width = 126
$Button_Task5.Height = 35
$Button_Task5.location = new-object system.drawing.point(17,107)
$Button_Task5.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task5)

$Button_Task7 = New-Object system.windows.Forms.Button
$Button_Task7.Text = "Sort Tabs"
$Button_Task7.Width = 126
$Button_Task7.Height = 35
$Button_Task7.location = new-object system.drawing.point(17,154)
$Button_Task7.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task7)
     
	 
	 
##----------------------------LOAD VARIABLES--------------------------------##

#"Select Sheet" Variables
Try
{
   $global:DataSheet = import-clixml ./Files/DataSheet.clixml
}
Catch
{
}

#"Manage Rules" Variables
Try
{
	[System.Collections.ArrayList]$Global:RuleListArray=@()
	$Global:RuleListArray = Import-Csv "./Files/Rules.csv"
	$Global:Match=$false
	$Global:ABC=$false
	[System.Collections.ArrayList]$Global:RuleArray=@()
	[System.Collections.ArrayList]$Global:MatchList=@()
	if ($Global:RuleListArray.count -eq 0)
	{
	[System.Collections.ArrayList]$Global:RuleListArray=@()
	}
	
}
Catch
{
	[System.Collections.ArrayList]$Global:RuleListArray=@()
	[System.Collections.ArrayList]$Global:RuleArray=@()
	[System.Collections.ArrayList]$Global:MatchList=@()
	$Global:Match=$false
	$Global:ABC=$false
}

#"Apply Rules" Variables
[System.Collections.ArrayList]$Global:MatchColumn=@()
[System.Collections.ArrayList]$Global:MatchColumnABC=@()
[System.Collections.ArrayList]$Global:Length=@()
[System.Collections.ArrayList]$Global:LengthName=@()
[System.Collections.ArrayList]$Global:THEVARIABLE_Number=@()
[System.Collections.ArrayList]$Global:NameArray=@()


#"Sort Tabs" Variables
Try
{
	[System.Collections.ArrayList]$Global:TabRuleArray=@()
	[System.Collections.ArrayList]$Global:TabCopyColumn=@()
	$Global:TabRuleArray = import-clixml ./Files/TabRules.clixml
	$Global:NameLocation = import-clixml ./Files/NameLocation.clixml
	$Global:String2 = import-clixml ./Files/Location.clixml
	
   if ($Global:TabRuleArray.count -eq 0)
	{
	[System.Collections.ArrayList]$Global:TabRuleArray=@()
	}
}
Catch
{
[System.Collections.ArrayList]$Global:TabRuleArray=@()
[System.Collections.ArrayList]$Global:TabCopyColumn=@()
}


##----------------------------MAIN PAGE COMMANDS----------------------------##

    #Choose Sheet
    $Button_Task1.Add_Click({
	$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
	$New_Sheets=@()

if ($global:DataSheet -ne $null)
{
	$New_Sheets += $global:DataSheet
}

	$Sheets = $excel.sheets | Select-Object -Property Name
	$Sheets= $Sheets -replace '@{.+?=' -replace '}'
	$New_Sheets += $Sheets

		$ChooseSheet = New-Object System.Windows.Forms.Form 
		$ChooseSheet.Text = "TANYR"
		$ChooseSheet.Size = New-Object System.Drawing.Size(185,190) 
		$ChooseSheet.StartPosition = "CenterScreen"
		$ChooseSheet.Topmost = $True
		$ChooseSheet.Font = $Font
		$ChooseSheet.FormBorderStyle = 'FixedDialog'

		$objTextBox = New-Object System.Windows.Forms.ComboBox
		$objTextBox.Location = New-Object System.Drawing.Size(10,30) 
		$objTextBox.Size = New-Object System.Drawing.Size(150,20) 
		$objTextBox.DropDownStyle= 'DropDownList'
		
		$New_Sheets = $New_Sheets | select -uniq
		$objTextBox.Items.AddRange(($New_Sheets));
		$objTextBox.SelectedIndex = 0
		
		$objLabel = New-Object System.Windows.Forms.Label
		$objLabel.Location = New-Object System.Drawing.Size(10,15) 
		$objLabel.Size = New-Object System.Drawing.Size(305,25) 
		$objLabel.Text = "Please select the data sheet:"
		
		$OKButton = New-Object System.Windows.Forms.Button
		$OKButton.Location = New-Object System.Drawing.Size(10,120)
		$OKButton.Size = New-Object System.Drawing.Size(75,23)
		$OKButton.Text = "OK"
		$ChooseSheet.AcceptButton = $OKButton
		$OKButton.Add_Click({
		$global:DataSheet=$objTextBox.Text;
		
		$ChooseSheet.Close();
		$global:DataSheet | export-clixml ./Files/DataSheet.clixml})
		
		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Size(85,120)
		$CancelButton.Size = New-Object System.Drawing.Size(75,23)
		$CancelButton.Text = "Cancel"
		$CancelButton.Add_Click({$ChooseSheet.Close()})
		
		$ChooseSheet.Icon = $Icon
		$ChooseSheet.Controls.Add($CancelButton)
		$ChooseSheet.Controls.Add($objTextBox) 
		$ChooseSheet.Controls.Add($objLabel) 
		$ChooseSheet.Controls.Add($OKButton)
		$ChooseSheet.Add_Shown({$ChooseSheet.Activate();$objTextBox.focus()})
		[void] $ChooseSheet.ShowDialog()
	})

    #Manage Rules
    $Button_Task3.Add_Click({
	
	
	})	

    #Apply Rules
    $Button_Task5.Add_Click({
	$ie = New-Object -ComObject 'InternetExplorer.Application'
	$ie.AddressBar = $false
	$ie.MenuBar = $false
	$ie.StatusBar = $false
	$ie.ToolBar = $false
	$ie.Visible = $true
	$ie.navigate("http://moridb.com/items/furniture/")
	
	While ($ie.Busy)
{
    Sleep -Milliseconds 100
}

	$Document=$ie.document
	Sleep -Milliseconds 1000
	
	$UsernameBox=$Document.getElementsByTagName("search") | where-object {$_.name -eq "q"}
	$UsernameBox.value="Help"
<# 
	$PasswordBox=$Document.getElementsByTagName("input") | where-object {$_.name -eq "password"}
	$PasswordBox.value="$Password"

    Sleep -Milliseconds 1000
	
	$LoginButton=$Document.getElementById("login-button")
	$LoginButton.click()

While ($ie.Busy)
{
    Sleep -Milliseconds 100
}
    Sleep -Milliseconds 500
$Catchpa=$Document.getElementsByTagName("strong") | where-object {$_.innerText-eq " What was your first car? "}
$Catchpa2=$Document.getElementsByTagName("strong") | where-object {$_.innerText -eq " What was your high school mascot? "}
    Sleep -Milliseconds 500

if ($Catchpa -ne $null)
{
$verifyAnswer=$Document.getElementById("verifyAnswer")
$verifyAnswer.value="Camero"
$verifyAnswer=$Document.getElementById("VerifyButton")
$verifyAnswer.click()
}

if ($Catchpa2 -ne $null)
{
$verifyAnswer=$Document.getElementById("verifyAnswer")
$verifyAnswer.value="hawks"
$verifyAnswer=$Document.getElementById("VerifyButton")
$verifyAnswer.click()
}

    Sleep -Milliseconds 1000

	$ie.navigate("https://remits.zirmed.com/Payments_NoDX.aspx?appid=13")
		
	While ($ie.Busy)
{
    Sleep -Milliseconds 100
}

$global:AmountBox=$null
$global:SearchButton=$null
$global:PopOut=$null
$global:DateOne=$null
$global:DateTwo=$null
$global:NotFound=@()

for($i=0; $i -lt $global:Credit.Length; $i++)
{
	if($AmountBox -eq $null){
	$global:AmountBox=$Document.getElementById("txtAmount")
	}
	$global:AmountBox.value=$global:Credit[$i]
	
	if($global:DateOne -eq $null){
	$global:DateOne=$Document.getElementsByTagName("input") | where-object {$_.name -eq "ctl11"}
	$global:DateTwo=$Document.getElementsByTagName("input") | where-object {$_.name -eq "ctl13"}
	$global:DateOne.value=$global:dtmDate
	$global:DateTwo.value=$global:dtmDate
	}
	
	
	if($SearchButton -eq $null){
	$global:SearchButton=$Document.getElementById("btnSearch")
	}
	$global:SearchButton.click()
	
	
	sleep -milliseconds 1000
	
	$viewEOB=$Document.getElementById("viewEOB")
	if($viewEOB -eq $null)
	{
	$global:NotFound += $global:Name[$i]
	Continue
	}
	$viewEOB.click()
    
    sleep -milliseconds 2000

  	$global:PopOut=$Document.getElementsByTagName("span") | where-object {$_.innerText -eq "popout"}
	$global:PopOut.click()
}
$ie.quit()
$objTextBox3.text="Done!"
$empty=@()
if ($global:NotFound -ne $null)
{
		$NotFoundPDFs = New-Object System.Windows.Forms.Form 
		$NotFoundPDFs.Text = "TANYR"
		$NotFoundPDFs.Size = New-Object System.Drawing.Size(320,220) 
		$NotFoundPDFs.StartPosition = "CenterScreen"
		$NotFoundPDFs.Icon = $Icon
		$NotFoundPDFs.Font = New-Object System.Drawing.Font("Times New Roman",10)
		$NotFoundPDFs.FormBorderStyle = 'FixedDialog'

		$UpdateLabel = New-Object System.Windows.Forms.Label
		$UpdateLabel.Location = New-Object System.Drawing.Size(5,7) 
		$UpdateLabel.Size = New-Object System.Drawing.Size(305,15) 
		$UpdateLabel.Text = "PDFs Not Found:"
		$NotFoundPDFs.Controls.Add($UpdateLabel)
		
		$outputBox = New-Object System.Windows.Forms.TextBox 
		$outputBox.Location = New-Object System.Drawing.Size(7,23)
		$outputBox.Size = New-Object System.Drawing.Size(291,155) 
		$outputBox.MultiLine = $True 
		$outputBox.ScrollBars = "Vertical"
		
		for ($i=0;$i -lt $global:NotFound.Length-1;$i++)
		{		
		$empty += "`r`n"+"`r`n"
		}	
		
		for ($i=0;$i -lt $global:NotFound.Length;$i++)
		{		
		$outputBox.text += $global:NotFound[$i]+$empty[$i]
		}	
		$NotFoundPDFs.Controls.Add($outputBox) 

		$NotFoundPDFs.Add_Shown({$NotFoundPDFs.Activate()})
		[void] $NotFoundPDFs.ShowDialog() 
} #>


	})
	
	#Sort Tabs
	$Button_Task7.Add_Click({})
	

[void]$TANYRHealthcare.ShowDialog()
$TANYRHealthcare.Dispose()