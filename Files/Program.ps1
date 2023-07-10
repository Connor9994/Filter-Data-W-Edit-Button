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
	
		$ManageRules = New-Object System.Windows.Forms.Form 
		$ManageRules.Text = "TANYR"
		$ManageRules.Size = New-Object System.Drawing.Size(205,300) 
		$ManageRules.StartPosition = "CenterScreen"
		$ManageRules.Topmost = $True
		$ManageRules.Font = $Font
		$ManageRules.FormBorderStyle = 'FixedDialog'
		
		$RuleList = New-Object System.Windows.Forms.ListBox
		$RuleList.Location = New-Object System.Drawing.Size(10,40) 
		$RuleList.Size = New-Object System.Drawing.Size(155,215)
		
		if ($Global:RuleListArray.Count -ne 0)
		{
		$RuleList.Items.AddRange($Global:RuleListArray.Name)
		}
		
		$MoveUp = New-Object System.Windows.Forms.Button
		$MoveUp.Location = New-Object System.Drawing.Size(166,40) 
		$MoveUp.Size = New-Object System.Drawing.Size(20,20)
		$MoveUp.Text = "/\"
		$MoveUp.Add_Click({
		 # only if the first item isn't the current one
    if($RuleList.SelectedIndex -gt 0)
    {
        $RuleList.BeginUpdate()
        #Get starting position
        $pos = $RuleList.selectedIndex
        # add a duplicate of original item up in the listbox
        $RuleList.items.insert($pos -1,$RuleList.Items.Item($pos))
		$Global:RuleListArray[$pos],$Global:RuleListArray[$pos-1] = $Global:RuleListArray[$pos-1],$Global:RuleListArray[$pos]
        # make it the current item
        $RuleList.SelectedIndex = ($pos -1)
        # delete the old occurrence of this item
        $RuleList.Items.RemoveAt($pos +1)
        $RuleList.EndUpdate()
		$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
    }
	ELSE
	{}
	})
	
		$MoveDown = New-Object System.Windows.Forms.Button
		$MoveDown.Location = New-Object System.Drawing.Size(166,60) 
		$MoveDown.Size = New-Object System.Drawing.Size(20,20)
		$MoveDown.Text = "\/"
		$MoveDown.Add_Click({
		
		if(($RuleList.SelectedIndex -ne -1)   -and   ($RuleList.SelectedIndex -lt $RuleList.Items.Count - 1)    )   {
        $RuleList.BeginUpdate()
        #Get starting position 
        $pos = $RuleList.selectedIndex
        # add a duplicate of item below in the listbox
        $RuleList.items.insert($pos,$RuleList.Items.Item($pos +1))
		$Global:RuleListArray[$pos],$Global:RuleListArray[$pos+1] = $Global:RuleListArray[$pos+1],$Global:RuleListArray[$pos]
        # delete the old occurrence of this item
        $RuleList.Items.RemoveAt($pos +2 )
        # move to current item
        $RuleList.SelectedIndex = ($pos +1)
        $RuleList.EndUpdate()
		$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
   }
   ELSE
   {}
	})
	
		$EditRule = New-Object System.Windows.Forms.Button
		$EditRule.Location = New-Object System.Drawing.Size(166,80) 
		$EditRule.Size = New-Object System.Drawing.Size(20,20)
		$EditRule.Text = "*"
		$EditRule.Add_Click({
		$j=$RuleList.selectedIndex
		$Here = $Global:RuleListArray | Select-Object -index $j
		if($j -ge 0)
		{
		
	$CreateRule = New-Object System.Windows.Forms.Form 
	$CreateRule.Text = "TANYR"
	$CreateRule.Size = New-Object System.Drawing.Size(365,475) 
	$CreateRule.StartPosition = "CenterScreen"
	$CreateRule.Topmost = $True
	$CreateRule.Font = $Font
	$CreateRule.FormBorderStyle = 'FixedDialog'
	
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size(20,40) 
	$NameLabel.Size = New-Object System.Drawing.Size(305,25) 
	$NameLabel.Text = "Rule Name:"
	
	$NameBox = New-Object System.Windows.Forms.TextBox
	$NameBox.Location = New-Object System.Drawing.Size(20,55) 
	$NameBox.Size = New-Object System.Drawing.Size(150,40) 
	$NameBox.ReadOnly = $true
	$Namebox.Text=$Global:RuleListArray.name[$j]
	
	$AssignedName = New-Object System.Windows.Forms.Label
	$AssignedName.Location = New-Object System.Drawing.Size(20,80) 
	$AssignedName.Size = New-Object System.Drawing.Size(305,15) 
	$AssignedName.Text = "Assigned Name:"
	
	$AssignedNameBox = New-Object System.Windows.Forms.TextBox
	$AssignedNameBox.Location = New-Object System.Drawing.Size(20,95) 
	$AssignedNameBox.Size = New-Object System.Drawing.Size(150,40)
	$AssignedNameBox.Text=$Global:RuleListArray.AssignedName[$j]
	$AssignedNameBox.Add_TextChanged({
	$Here.AssignedName=$AssignedNameBox.Text
	$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
	})
	
##--------------------------------------------MATCH BOX------------------------------------------#
			$checkBoxColumn = New-Object System.Windows.Forms.CheckBox 
			$checkBoxColumn.Text="Match Values?"
			$checkBoxColumn.Location = New-Object System.Drawing.Size(5,137) 
			$checkBoxColumn.Size = New-Object System.Drawing.Size(100,23)
			if($Global:RuleListArray."Match?"[$j] -eq $true)
			{
			$checkBoxColumn.Checked=$True
			}
			
			$ColumnToMatchText = New-Object System.Windows.Forms.Label
			$ColumnToMatchText.Location = New-Object System.Drawing.Size(20,120) 
			$ColumnToMatchText.Size = New-Object System.Drawing.Size(100,15) 
			$ColumnToMatchText.Text = "Column to match:"
			$ColumnToMatchText.Visible = $false
			
			$ColumnToMatch = New-Object System.Windows.Forms.ComboBox
			$ColumnToMatch.Location = New-Object System.Drawing.Size(20,135) 
			$ColumnToMatch.Size = New-Object System.Drawing.Size(150,20) 
			$ColumnToMatch.DropDownStyle= 'DropDownList'
			$ColumnToMatch.Visible = $false
			
			$MatchList = New-Object System.Windows.Forms.ListBox
			$MatchList.Location = New-Object System.Drawing.Size(20,210) 
			$MatchList.Size = New-Object System.Drawing.Size(148,215)
			$MatchList.Visible = $false
			if($Global:RuleListArray."Match?"[$j] -eq $True)
			{
			if($Global:RuleListArray.Matching[$j] -ne $null)
			{
			$ValueOfMatching = $Global:RuleListArray.Matching[$j].split(",")
			$MatchList.Items.AddRange($ValueOfMatching)
			$Global:RuleArray.AddRange($ValueOfMatching)
			}
			}
			
			$AddToList = New-Object System.Windows.Forms.TextBox
			$AddToList.Location = New-Object System.Drawing.Size(20,188) 
			$AddToList.Size = New-Object System.Drawing.Size(148,35)
			$AddToList.Visible = $false
			
			$AddButton2 = New-Object System.Windows.Forms.Button
			$AddButton2.Location = New-Object System.Drawing.Size(20,163) 
			$AddButton2.Size = New-Object System.Drawing.Size(75,23)
			$AddButton2.Text = "Add"
			$AddButton2.Visible = $false
			$AddButton2.Add_Click({
			If ($MatchList.Items -contains $AddToList.Text)
			{													
			}
			else
			{
			$MatchList.Items.Add($AddToList.Text)
			$Global:RuleArray.Add($AddToList.Text)
			$AddToList.Text=""
			for ($i=0; $i -lt $Global:RuleArray.Count; $i++)
			{
			$RuleString = $RuleString + $Global:RuleArray[$i] + ","
			}
			$RuleString = $RuleString.TrimEnd(',')
			$Here.Matching=$RuleString
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			$RuleString=""
			}
			})
			
			$RemoveButton2 = New-Object System.Windows.Forms.Button
			$RemoveButton2.Location = New-Object System.Drawing.Size(85,163) 
			$RemoveButton2.Size = New-Object System.Drawing.Size(80,23)
			$RemoveButton2.Text = "Remove"
			$RemoveButton2.Visible = $false
			$RemoveButton2.Add_Click({
			$Global:RuleArray.RemoveAt($MatchList.SelectedIndex)
			$MatchList.Items.RemoveAt($MatchList.SelectedIndex)
			for ($i=0; $i -lt $Global:RuleArray.Count; $i++)
			{
			$RuleString = $RuleString + $Global:RuleArray[$i] + ","
			}
			$RuleString = $RuleString.TrimEnd(',')
			$Here.Matching=$RuleString
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			$RuleString=""
			})

			if ($checkBoxColumn.Checked -eq $True)
			{
			$Global:Match=$True
			$ColumnToMatchText.Visible = $true
			$ColumnToMatch.Visible = $true
			$MatchList.Visible = $true
			$AddToList.Visible = $true
			$AddButton2.Visible = $true
			$RemoveButton2.Visible = $true
			if ($ColumnToMatch.Items.Count -le 1)
			{
			$ColumnToMatch.Items.Add("Loading...")
			$ColumnToMatch.SelectedIndex = 0
			$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
			$worksheet = $excel.sheets.item($global:DataSheet)
			$count=$worksheet.UsedRange.columns.count
			$Name=@()
			$Name+= $Global:RuleListArray.MatchColumnName[$j]
			for($i=1; $i -le $count; $i++)
			{
			$Name += $worksheet.Cells.Item(1,$i).Text
			}
			$Name = $Name | select -uniq
			$ColumnToMatch.Items.AddRange($Name)
			$ColumnToMatch.Items.Remove("Loading...")
			}
			}
			
			$ColumnToMatch.Add_SelectionChangeCommitted({
			$Here.MatchColumnName=$ColumnToMatch.Text
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			})
	
			$checkBoxColumn.Add_Click({
			if ($checkBoxColumn.Checked -eq $True)
			{
			$Global:Match=$True
			$Here."Match?"=$Match
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			$ColumnToMatchText.Visible = $true
			$ColumnToMatch.Visible = $true
			$MatchList.Visible = $true
			$AddToList.Visible = $true
			$AddButton2.Visible = $true
			$RemoveButton2.Visible = $true
			if ($ColumnToMatch.Items.Count -le 1)
			{
			$ColumnToMatch.Items.Add("Loading...")
			$ColumnToMatch.SelectedIndex = 0
			$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
			$worksheet = $excel.sheets.item($global:DataSheet)
			$count=$worksheet.UsedRange.columns.count
			$Name=@()
			for($i=1; $i -le $count; $i++)
			{
			$Name += $worksheet.Cells.Item(1,$i).Text
			}
			$ColumnToMatch.Items.AddRange($Name)
			$ColumnToMatch.Items.Remove("Loading...")
			}
			}
			else
			{
			$Global:Match=$False
			$Here."Match?"=$Match
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			$ColumnToMatchText.Visible = $false
			$ColumnToMatch.Visible = $false
			$MatchList.Visible = $false
			$AddToList.Visible = $false
			$AddButton2.Visible = $false
			$RemoveButton2.Visible = $false
			$ColumnToMatch.SelectedIndex = -1
			}
			})
			
##--------------------------------------------Alphabatize------------------------------------------#	

			$checkBoxColumn_ABC = New-Object System.Windows.Forms.CheckBox 
			$checkBoxColumn_ABC.Text="Alphabetize?"
			$checkBoxColumn_ABC.Location = New-Object System.Drawing.Size(175,137) 
			$checkBoxColumn_ABC.Size = New-Object System.Drawing.Size(100,23) 
			if($Global:RuleListArray."Alphabetize?"[$j] -eq $true)
			{
			$checkBoxColumn_ABC.Checked=$True
			}

			$ColumnToMatchText_123 = New-Object System.Windows.Forms.Label
			$ColumnToMatchText_123.Location = New-Object System.Drawing.Size(190,120) 
			$ColumnToMatchText_123.Size = New-Object System.Drawing.Size(100,15) 
			$ColumnToMatchText_123.Text = "Column to match:"
			$ColumnToMatchText_123.Visible = $false
			
			$ColumnToMatch_ABC = New-Object System.Windows.Forms.ComboBox
			$ColumnToMatch_ABC.Location = New-Object System.Drawing.Size(190,135) 
			$ColumnToMatch_ABC.Size = New-Object System.Drawing.Size(150,20) 
			$ColumnToMatch_ABC.DropDownStyle= 'DropDownList'
			$ColumnToMatch_ABC.Visible = $false
			
			$ColumnToMatchText_ABC = New-Object System.Windows.Forms.Label
			$ColumnToMatchText_ABC.Location = New-Object System.Drawing.Size(190,160) 
			$ColumnToMatchText_ABC.Size = New-Object System.Drawing.Size(100,15) 
			$ColumnToMatchText_ABC.Text = "Range to match:"
			$ColumnToMatchText_ABC.Visible = $false
		
			$HelperText_ABC = New-Object System.Windows.Forms.Label
			$HelperText_ABC.Location = New-Object System.Drawing.Size(233,178) 
			$HelperText_ABC.Size = New-Object System.Drawing.Size(15,15) 
			$HelperText_ABC.Text = "to"
			$HelperText_ABC.Visible = $false
			
			$FirstLetter= New-Object System.Windows.Forms.ComboBox
			$FirstLetter.Location = New-Object System.Drawing.Size(192,175) 
			$FirstLetter.Size = New-Object System.Drawing.Size(40,20) 
			$FirstLetter.DropDownStyle= 'DropDownList'
			$FirstLetter.Visible = $false
			$FirstLetter.Add_SelectionChangeCommitted({
			$Here.FirstLetter= $FirstLetter.Text
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			})	
			
			$SecondLetter= New-Object System.Windows.Forms.ComboBox
			$SecondLetter.Location = New-Object System.Drawing.Size(248,175) 
			$SecondLetter.Size = New-Object System.Drawing.Size(40,20) 
			$SecondLetter.DropDownStyle= 'DropDownList'
			$SecondLetter.Visible = $false
			$SecondLetter.Add_SelectionChangeCommitted({
			$Here.SecondLetter= $SecondLetter.Text
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			})	
			
			$FirstLetter.Items.AddRange("A");$FirstLetter.Items.AddRange("B");$FirstLetter.Items.AddRange("C")
			$FirstLetter.Items.AddRange("D");$FirstLetter.Items.AddRange("E");$FirstLetter.Items.AddRange("F")
			$FirstLetter.Items.AddRange("G");$FirstLetter.Items.AddRange("H");$FirstLetter.Items.AddRange("I")
			$FirstLetter.Items.AddRange("J");$FirstLetter.Items.AddRange("K");$FirstLetter.Items.AddRange("L")
			$FirstLetter.Items.AddRange("M");$FirstLetter.Items.AddRange("N");$FirstLetter.Items.AddRange("O")
			$FirstLetter.Items.AddRange("P");$FirstLetter.Items.AddRange("Q");$FirstLetter.Items.AddRange("R")
			$FirstLetter.Items.AddRange("S");$FirstLetter.Items.AddRange("T");$FirstLetter.Items.AddRange("U")
			$FirstLetter.Items.AddRange("V");$FirstLetter.Items.AddRange("W");$FirstLetter.Items.AddRange("X")
			$FirstLetter.Items.AddRange("Y");$FirstLetter.Items.AddRange("Z")
			
			$SecondLetter.Items.AddRange("A");$SecondLetter.Items.AddRange("B");$SecondLetter.Items.AddRange("C")
			$SecondLetter.Items.AddRange("D");$SecondLetter.Items.AddRange("E");$SecondLetter.Items.AddRange("F")
			$SecondLetter.Items.AddRange("G");$SecondLetter.Items.AddRange("H");$SecondLetter.Items.AddRange("I")
			$SecondLetter.Items.AddRange("J");$SecondLetter.Items.AddRange("K");$SecondLetter.Items.AddRange("L")
			$SecondLetter.Items.AddRange("M");$SecondLetter.Items.AddRange("N");$SecondLetter.Items.AddRange("O")
			$SecondLetter.Items.AddRange("P");$SecondLetter.Items.AddRange("Q");$SecondLetter.Items.AddRange("R")
			$SecondLetter.Items.AddRange("S");$SecondLetter.Items.AddRange("T");$SecondLetter.Items.AddRange("U")
			$SecondLetter.Items.AddRange("V");$SecondLetter.Items.AddRange("W");$SecondLetter.Items.AddRange("X")
			$SecondLetter.Items.AddRange("Y");$SecondLetter.Items.AddRange("Z")
			
			if ($checkBoxColumn_ABC.Checked -eq $True)
			{
			$Global:ABC=$True
			$ColumnToMatchText_ABC.Visible = $true
			$ColumnToMatchText_123.Visible = $true
			$HelperText_ABC.Visible = $true
			$SecondLetter.Visible = $true
			$ColumnToMatch_ABC.Visible = $true
			$checkBoxColumn_ABC.Size = New-Object System.Drawing.Size(15,23) 
			$FirstLetter.Visible = $true
			$SecondLetter.Visible = $true
			$checkBoxColumn_ABC.Text = ""
			if ($ColumnToMatch_ABC.Items.Count -le 1)
			{
			$ColumnToMatch_ABC.Items.Add("Loading...")
			$ColumnToMatch_ABC.SelectedIndex = 0
			$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
			$worksheet = $excel.sheets.item($global:DataSheet)
			$count=$worksheet.UsedRange.columns.count
			$Name=@()
			$Name += $Global:RuleListArray.ABCColumnName[$j]
			for($i=1; $i -le $count; $i++)
			{
			$Name += $worksheet.Cells.Item(1,$i).Text
			}
			$Name = $Name | select -uniq
			$ColumnToMatch_ABC.Items.AddRange($Name)
			$ColumnToMatch_ABC.Items.Remove("Loading...")
			}
			}
			
			$ColumnToMatch_ABC.Add_SelectionChangeCommitted({
			$Here.ABCColumnName= $ColumnToMatch_ABC.Text
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			})	
			
			if($Global:RuleListArray."Alphabetize?" -eq $True)
			{
			$index1 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".IndexOf($Global:RuleListArray.FirstLetter[$j])
			$index2 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".IndexOf($Global:RuleListArray.SecondLetter[$j])
			$FirstLetter.SelectedIndex=$index1
			$SecondLetter.SelectedIndex=$index2
			}
			
			$checkBoxColumn_ABC.Add_Click({
			if ($checkBoxColumn_ABC.Checked -eq $True)
			{
			$Global:ABC=$True
			$Here."Alphabetize?"=$ABC
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			$ColumnToMatchText_ABC.Visible = $true
			$ColumnToMatchText_123.Visible = $true
			$HelperText_ABC.Visible = $true
			$SecondLetter.Visible = $true
			$ColumnToMatch_ABC.Visible = $true
			$checkBoxColumn_ABC.Size = New-Object System.Drawing.Size(15,23) 
			$FirstLetter.Visible = $true
			$SecondLetter.Visible = $true
			$checkBoxColumn_ABC.Text = ""
			if ($ColumnToMatch_ABC.Items.Count -le 1)
			{
			$ColumnToMatch_ABC.Items.Add("Loading...")
			$ColumnToMatch_ABC.SelectedIndex = 0
			$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
			$worksheet = $excel.sheets.item($global:DataSheet)
			$count=$worksheet.UsedRange.columns.count
			$Name=@()
			for($i=1; $i -le $count; $i++)
			{
			$Name += $worksheet.Cells.Item(1,$i).Text
			}
			$ColumnToMatch_ABC.Items.AddRange($Name)
			$ColumnToMatch_ABC.Items.Remove("Loading...")
			}
			}
			else
			{
			$Global:ABC=$false
			$Here."Alphabetize?"=$ABC
			$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
			$ColumnToMatchText_ABC.Visible = $false
			$ColumnToMatchText_123.Visible = $false
			$HelperText_ABC.Visible = $false
			$ColumnToMatch_ABC.Visible = $false
			$checkBoxColumn_ABC.Size = New-Object System.Drawing.Size(100,23) 
			$FirstLetter.Visible = $false
			$SecondLetter.Visible = $false
			$checkBoxColumn_ABC.Text = "Alphabetize?"
			}
			})
			
	$AddButton = New-Object System.Windows.Forms.Button
	$AddButton.Location = New-Object System.Drawing.Size(20,13) 
	$AddButton.Size = New-Object System.Drawing.Size(77,23)
	$AddButton.Text = "Change Rule"
	$CreateRule.AcceptButton = $AddButton
	$AddButton.Add_Click({
	
	$Here = $Global:RuleListArray | Select-Object -index $j
	Write-Host $Global:RuleArray -Separator `n -foregroundcolor "Green"
	Write-Host $Global:MatchList -Separator `n -foregroundcolor "Blue"
	$Here.AssignedName=$AssignedNameBox.Text
	$Here."Match?"=$Match
	$Here.MatchColumnName=$ColumnToMatch.Text
	$Here."Alphabetize?"=$ABC
	$Here.ABCColumnName= $ColumnToMatch_ABC.Text
	$Here.FirstLetter= $FirstLetter.Text
	$Here.SecondLetter= $SecondLetter.Text
	[System.Collections.ArrayList]$Global:RuleArray=@()
	[System.Collections.ArrayList]$Global:MatchList=@()
	
	if ($Global:RuleListArray.count -ne 0)
		{
		$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
		}
	$CreateRule.close()
	})
	$CreateRule.Icon = $Icon
	$CreateRule.Controls.Add($NameBox)  
	$CreateRule.Controls.Add($AddButton)
	$CreateRule.Controls.Add($NameLabel)
	$CreateRule.Controls.Add($AssignedName)
	$CreateRule.Controls.Add($AssignedNameBox)
	$CreateRule.Controls.Add($ColumnToMatch)
	$CreateRule.Controls.Add($ColumnToMatchText)
	$CreateRule.Controls.Add($checkBoxColumn)
	$CreateRule.Controls.Add($MatchList)
	$CreateRule.Controls.Add($AddToList)
	$CreateRule.Controls.Add($AddButton2)
	$CreateRule.Controls.Add($RemoveButton2)
	$CreateRule.Controls.Add($checkBoxColumn_ABC)
	$CreateRule.Controls.Add($ColumnToMatchText_ABC)
	$CreateRule.Controls.Add($ColumnToMatchText_123)
	$CreateRule.Controls.Add($FirstLetter)
	$CreateRule.Controls.Add($HelperText_ABC)
	$CreateRule.Controls.Add($SecondLetter)
	$CreateRule.Controls.Add($ColumnToMatch_ABC)
	
	$CreateRule.Add_Shown({$CreateRule.Activate();$AssignedNameBox.focus()})
	[void] $CreateRule.ShowDialog()
	}
		})
		
		$CreateButton = New-Object System.Windows.Forms.Button
		$CreateButton.Location = New-Object System.Drawing.Size(10,13) 
		$CreateButton.Size = New-Object System.Drawing.Size(72,23)
		$CreateButton.Text = "Create Rule"
		$ManageRules.AcceptButton = $CreateButton
		$CreateButton.Add_Click({
		
	[System.Collections.ArrayList]$Global:RuleArray=@()
	[System.Collections.ArrayList]$Global:MatchList=@()
	$CreateRule = New-Object System.Windows.Forms.Form 
	$CreateRule.Text = "TANYR"
	$CreateRule.Size = New-Object System.Drawing.Size(365,475)
	$CreateRule.StartPosition = "CenterScreen"
	$CreateRule.Topmost = $True
	$CreateRule.Font = $Font
	$CreateRule.FormBorderStyle = 'FixedDialog'
	
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size(20,40) 
	$NameLabel.Size = New-Object System.Drawing.Size(305,25) 
	$NameLabel.Text = "Rule Name:"
	
	$NameBox = New-Object System.Windows.Forms.TextBox
	$NameBox.Location = New-Object System.Drawing.Size(20,55) 
	$NameBox.Size = New-Object System.Drawing.Size(150,40) 
	
	$AssignedName = New-Object System.Windows.Forms.Label
	$AssignedName.Location = New-Object System.Drawing.Size(20,80) 
	$AssignedName.Size = New-Object System.Drawing.Size(305,15) 
	$AssignedName.Text = "Assigned Name:"
	
	$AssignedNameBox = New-Object System.Windows.Forms.TextBox
	$AssignedNameBox.Location = New-Object System.Drawing.Size(20,95) 
	$AssignedNameBox.Size = New-Object System.Drawing.Size(150,40)
	
##--------------------------------------------MATCH BOX------------------------------------------#
			$checkBoxColumn = New-Object System.Windows.Forms.CheckBox 
			$checkBoxColumn.Text="Match Values?"
			$checkBoxColumn.Location = New-Object System.Drawing.Size(5,137) 
			$checkBoxColumn.Size = New-Object System.Drawing.Size(100,23) 
			$checkBoxColumn.Checked = $False
			$Global:Match=$False
			
			$ColumnToMatchText = New-Object System.Windows.Forms.Label
			$ColumnToMatchText.Location = New-Object System.Drawing.Size(20,120) 
			$ColumnToMatchText.Size = New-Object System.Drawing.Size(100,15) 
			$ColumnToMatchText.Text = "Column to match:"
			$ColumnToMatchText.Visible = $false
			
			$ColumnToMatch = New-Object System.Windows.Forms.ComboBox
			$ColumnToMatch.Location = New-Object System.Drawing.Size(20,135) 
			$ColumnToMatch.Size = New-Object System.Drawing.Size(150,20) 
			$ColumnToMatch.DropDownStyle= 'DropDownList'
			$ColumnToMatch.Visible = $false
			
			$MatchList = New-Object System.Windows.Forms.ListBox
			$MatchList.Location = New-Object System.Drawing.Size(20,210) 
			$MatchList.Size = New-Object System.Drawing.Size(148,215)
			$MatchList.Visible = $false
			
			$AddToList = New-Object System.Windows.Forms.TextBox
			$AddToList.Location = New-Object System.Drawing.Size(20,188) 
			$AddToList.Size = New-Object System.Drawing.Size(148,35)
			$AddToList.Visible = $false
			
			$AddButton2 = New-Object System.Windows.Forms.Button
			$AddButton2.Location = New-Object System.Drawing.Size(20,163) 
			$AddButton2.Size = New-Object System.Drawing.Size(75,23)
			$AddButton2.Text = "Add"
			$AddButton2.Visible = $false
			$AddButton2.Add_Click({
			If ($MatchList.Items -contains $AddToList.Text)
			{
			}
			else
			{
			$MatchList.Items.Add($AddToList.Text)
			$Global:RuleArray.Add($AddToList.Text)
			$AddToList.Text=""
			}
			})
			$RemoveButton2 = New-Object System.Windows.Forms.Button
			$RemoveButton2.Location = New-Object System.Drawing.Size(85,163) 
			$RemoveButton2.Size = New-Object System.Drawing.Size(80,23)
			$RemoveButton2.Text = "Remove"
			$RemoveButton2.Visible = $false
			$RemoveButton2.Add_Click({
			$Global:RuleArray.RemoveAt($MatchList.SelectedIndex)
			$MatchList.Items.RemoveAt($MatchList.SelectedIndex)
			})
	
			$checkBoxColumn.Add_Click({
			if ($checkBoxColumn.Checked -eq $True)
			{
			$Global:Match=$True
			$ColumnToMatchText.Visible = $true
			$ColumnToMatch.Visible = $true
			$MatchList.Visible = $true
			$AddToList.Visible = $true
			$AddButton2.Visible = $true
			$RemoveButton2.Visible = $true
			if ($ColumnToMatch.Items.Count -le 1)
			{
			$ColumnToMatch.Items.Add("Loading...")
			$ColumnToMatch.SelectedIndex = 0
			$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
			$worksheet = $excel.sheets.item($global:DataSheet)
			$count=$worksheet.UsedRange.columns.count
			$Name=@()
			for($i=1; $i -le $count; $i++)
			{
			$Name += $worksheet.Cells.Item(1,$i).Text
			}
			$ColumnToMatch.Items.AddRange($Name)
			$ColumnToMatch.Items.Remove("Loading...")
			}
			}
			else
			{
			$Global:Match=$False
			$ColumnToMatchText.Visible = $false
			$ColumnToMatch.Visible = $false
			$MatchList.Visible = $false
			$AddToList.Visible = $false
			$AddButton2.Visible = $false
			$RemoveButton2.Visible = $false
			$ColumnToMatch.SelectedIndex = -1
			}
			})
			
##--------------------------------------------Alphabatize------------------------------------------#	

			$checkBoxColumn_ABC = New-Object System.Windows.Forms.CheckBox 
			$checkBoxColumn_ABC.Text="Alphabetize?"
			$checkBoxColumn_ABC.Location = New-Object System.Drawing.Size(175,137) 
			$checkBoxColumn_ABC.Size = New-Object System.Drawing.Size(100,23)
			$checkBoxColumn_ABC.checked = $false
			$Global:ABC=$False			
			
			$ColumnToMatchText_123 = New-Object System.Windows.Forms.Label
			$ColumnToMatchText_123.Location = New-Object System.Drawing.Size(190,120) 
			$ColumnToMatchText_123.Size = New-Object System.Drawing.Size(100,15) 
			$ColumnToMatchText_123.Text = "Column to match:"
			$ColumnToMatchText_123.Visible = $false
			
			$ColumnToMatch_ABC = New-Object System.Windows.Forms.ComboBox
			$ColumnToMatch_ABC.Location = New-Object System.Drawing.Size(190,135) 
			$ColumnToMatch_ABC.Size = New-Object System.Drawing.Size(150,20) 
			$ColumnToMatch_ABC.DropDownStyle= 'DropDownList'
			$ColumnToMatch_ABC.Visible = $false
			
			$ColumnToMatchText_ABC = New-Object System.Windows.Forms.Label
			$ColumnToMatchText_ABC.Location = New-Object System.Drawing.Size(190,160) 
			$ColumnToMatchText_ABC.Size = New-Object System.Drawing.Size(100,15) 
			$ColumnToMatchText_ABC.Text = "Range to match:"
			$ColumnToMatchText_ABC.Visible = $false
		
			$HelperText_ABC = New-Object System.Windows.Forms.Label
			$HelperText_ABC.Location = New-Object System.Drawing.Size(233,178) 
			$HelperText_ABC.Size = New-Object System.Drawing.Size(15,15) 
			$HelperText_ABC.Text = "to"
			$HelperText_ABC.Visible = $false
			
			$FirstLetter= New-Object System.Windows.Forms.ComboBox
			$FirstLetter.Location = New-Object System.Drawing.Size(192,175) 
			$FirstLetter.Size = New-Object System.Drawing.Size(40,20) 
			$FirstLetter.DropDownStyle= 'DropDownList'
			$FirstLetter.Visible = $false	
			
			$SecondLetter= New-Object System.Windows.Forms.ComboBox
			$SecondLetter.Location = New-Object System.Drawing.Size(248,175) 
			$SecondLetter.Size = New-Object System.Drawing.Size(40,20) 
			$SecondLetter.DropDownStyle= 'DropDownList'
			$SecondLetter.Visible = $false
			
			$FirstLetter.Items.AddRange("A");$FirstLetter.Items.AddRange("B");$FirstLetter.Items.AddRange("C")
			$FirstLetter.Items.AddRange("D");$FirstLetter.Items.AddRange("E");$FirstLetter.Items.AddRange("F")
			$FirstLetter.Items.AddRange("G");$FirstLetter.Items.AddRange("H");$FirstLetter.Items.AddRange("I")
			$FirstLetter.Items.AddRange("J");$FirstLetter.Items.AddRange("K");$FirstLetter.Items.AddRange("L")
			$FirstLetter.Items.AddRange("M");$FirstLetter.Items.AddRange("N");$FirstLetter.Items.AddRange("O")
			$FirstLetter.Items.AddRange("P");$FirstLetter.Items.AddRange("Q");$FirstLetter.Items.AddRange("R")
			$FirstLetter.Items.AddRange("S");$FirstLetter.Items.AddRange("T");$FirstLetter.Items.AddRange("U")
			$FirstLetter.Items.AddRange("V");$FirstLetter.Items.AddRange("W");$FirstLetter.Items.AddRange("X")
			$FirstLetter.Items.AddRange("Y");$FirstLetter.Items.AddRange("Z")
			
			$SecondLetter.Items.AddRange("A");$SecondLetter.Items.AddRange("B");$SecondLetter.Items.AddRange("C")
			$SecondLetter.Items.AddRange("D");$SecondLetter.Items.AddRange("E");$SecondLetter.Items.AddRange("F")
			$SecondLetter.Items.AddRange("G");$SecondLetter.Items.AddRange("H");$SecondLetter.Items.AddRange("I")
			$SecondLetter.Items.AddRange("J");$SecondLetter.Items.AddRange("K");$SecondLetter.Items.AddRange("L")
			$SecondLetter.Items.AddRange("M");$SecondLetter.Items.AddRange("N");$SecondLetter.Items.AddRange("O")
			$SecondLetter.Items.AddRange("P");$SecondLetter.Items.AddRange("Q");$SecondLetter.Items.AddRange("R")
			$SecondLetter.Items.AddRange("S");$SecondLetter.Items.AddRange("T");$SecondLetter.Items.AddRange("U")
			$SecondLetter.Items.AddRange("V");$SecondLetter.Items.AddRange("W");$SecondLetter.Items.AddRange("X")
			$SecondLetter.Items.AddRange("Y");$SecondLetter.Items.AddRange("Z")
	
			$checkBoxColumn_ABC.Add_Click({
			if ($checkBoxColumn_ABC.Checked -eq $True)
			{
			$Global:ABC=$True
			$ColumnToMatchText_ABC.Visible = $true
			$ColumnToMatchText_123.Visible = $true
			$HelperText_ABC.Visible = $true
			$SecondLetter.Visible = $true
			$ColumnToMatch_ABC.Visible = $true
			$checkBoxColumn_ABC.Size = New-Object System.Drawing.Size(15,23) 
			$FirstLetter.Visible = $true
			$SecondLetter.Visible = $true
			$checkBoxColumn_ABC.Text = ""
			if ($ColumnToMatch_ABC.Items.Count -le 1)
			{
			$ColumnToMatch_ABC.Items.Add("Loading...")
			$ColumnToMatch_ABC.SelectedIndex = 0
			$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
			$worksheet = $excel.sheets.item($global:DataSheet)
			$count=$worksheet.UsedRange.columns.count
			$Name=@()
			for($i=1; $i -le $count; $i++)
			{
			$Name += $worksheet.Cells.Item(1,$i).Text
			}
			$ColumnToMatch_ABC.Items.AddRange($Name)
			$ColumnToMatch_ABC.Items.Remove("Loading...")
			}
			}
			else
			{
			$Global:ABC=$false
			$ColumnToMatchText_ABC.Visible = $false
			$ColumnToMatchText_123.Visible = $false
			$HelperText_ABC.Visible = $false
			$ColumnToMatch_ABC.Visible = $false
			$checkBoxColumn_ABC.Size = New-Object System.Drawing.Size(100,23) 
			$FirstLetter.Visible = $false
			$SecondLetter.Visible = $false
			$checkBoxColumn_ABC.Text = "Alphabetize?"
			}
			})
			
	$AddButton = New-Object System.Windows.Forms.Button
	$AddButton.Location = New-Object System.Drawing.Size(20,13) 
	$AddButton.Size = New-Object System.Drawing.Size(75,23)
	$AddButton.Text = "Add"
	$CreateRule.AcceptButton = $AddButton
	$AddButton.Add_Click({
	
	If ($Global:RuleListArray.Name -contains $NameBox.Text)
	{
	}
	else
	{
	for ($i=0; $i -lt $Global:RuleArray.Count; $i++)
	{
	$RuleString = $RuleString + $Global:RuleArray[$i] + ","
	}
	$RuleString = $RuleString.TrimEnd(',')
	
	Write-Host $Global:RuleArray -Separator `n -foregroundcolor "Green"
	Write-Host $Global:MatchList -Separator `n -foregroundcolor "Blue"
	$NewObject = New-Object System.Object
	$NewObject | Add-Member -type NoteProperty -name Name -Value $NameBox.Text
	$NewObject | Add-Member -type NoteProperty -name AssignedName -Value $AssignedNameBox.Text
	$NewObject | Add-Member -type NoteProperty -name "Match?" -Value $Match
	$NewObject | Add-Member -type NoteProperty -name MatchColumnName -Value $ColumnToMatch.Text
	$NewObject | Add-Member -type NoteProperty -name Matching -Value $RuleString
	$NewObject | Add-Member -type NoteProperty -name "Alphabetize?" -Value $ABC
	$NewObject | Add-Member -type NoteProperty -name ABCColumnName -Value $ColumnToMatch_ABC.Text
	$NewObject | Add-Member -type NoteProperty -name FirstLetter -Value $FirstLetter.Text
	$NewObject | Add-Member -type NoteProperty -name SecondLetter -Value $SecondLetter.Text
	$RuleString=""
	[System.Collections.ArrayList]$Global:RuleArray=@()
	[System.Collections.ArrayList]$Global:MatchList=@()
	$Global:RuleListArray.Add($NewObject)
	$RuleList.Items.AddRange($NameBox.Text)
	if ($Global:RuleListArray.count -ne 0)
		{
		$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
		}
	}
	$CreateRule.Close()
	})
	$CreateRule.Icon = $Icon
	$CreateRule.Controls.Add($NameBox)  
	$CreateRule.Controls.Add($AddButton)
	$CreateRule.Controls.Add($NameLabel)
	$CreateRule.Controls.Add($AssignedName)
	$CreateRule.Controls.Add($AssignedNameBox)
	$CreateRule.Controls.Add($ColumnToMatch)
	$CreateRule.Controls.Add($ColumnToMatchText)
	$CreateRule.Controls.Add($checkBoxColumn)
	$CreateRule.Controls.Add($MatchList)
	$CreateRule.Controls.Add($AddToList)
	$CreateRule.Controls.Add($AddButton2)
	$CreateRule.Controls.Add($RemoveButton2)
	$CreateRule.Controls.Add($checkBoxColumn_ABC)
	$CreateRule.Controls.Add($ColumnToMatchText_ABC)
	$CreateRule.Controls.Add($ColumnToMatchText_123)
	$CreateRule.Controls.Add($FirstLetter)
	$CreateRule.Controls.Add($HelperText_ABC)
	$CreateRule.Controls.Add($SecondLetter)
	$CreateRule.Controls.Add($ColumnToMatch_ABC)
	
	$CreateRule.Add_Shown({$CreateRule.Activate();$NameBox.focus()})
	[void] $CreateRule.ShowDialog()
		})
		
		$RemoveButton = New-Object System.Windows.Forms.Button
		$RemoveButton.Location = New-Object System.Drawing.Size(85,13) 
		$RemoveButton.Size = New-Object System.Drawing.Size(80,23)
		$RemoveButton.Text = "Remove Rule"
		$RemoveButton.Add_Click({
		
		$Global:RuleListArray.RemoveAt($RuleList.SelectedIndex)
		$RuleList.Items.RemoveAt($RuleList.SelectedIndex)
		Write-Host $Global:RuleListArray.Name -Separator `n -foregroundcolor "Green"
		Write-Host $RuleList -Separator `n -foregroundcolor "Blue"
		if ($Global:RuleListArray.count -ne 0)
		{
		$Global:RuleListArray | export-csv -Path ./Files/Rules.csv -NoTypeInformation
		}
		})
		
		$ManageRules.Icon = $Icon
		$ManageRules.Controls.Add($RuleList) 
		$ManageRules.Controls.Add($CreateButton)
		$ManageRules.Controls.Add($RemoveButton)
		$ManageRules.Controls.Add($MoveUp)
		$ManageRules.Controls.Add($MoveDown)
		$ManageRules.Controls.Add($EditRule)
		[System.Collections.ArrayList]$Global:RuleArray=@()
		[System.Collections.ArrayList]$Global:MatchList=@()
	
		$ManageRules.Add_Shown({$ManageRules.Activate();$RuleList.focus()})
		[void] $ManageRules.ShowDialog()
	})	

    #Apply Rules
    $Button_Task5.Add_Click({
	
		$ApplyRules = New-Object System.Windows.Forms.Form 
		$ApplyRules.Text = "TANYR"
		$ApplyRules.Size = New-Object System.Drawing.Size(190,190) 
		$ApplyRules.StartPosition = "CenterScreen"
		$ApplyRules.Topmost = $True
		$ApplyRules.Font = $Font
		$ApplyRules.FormBorderStyle = 'FixedDialog'

		$objTextBox = New-Object System.Windows.Forms.ComboBox
		$objTextBox.Location = New-Object System.Drawing.Size(10,30) 
		$objTextBox.Size = New-Object System.Drawing.Size(150,20) 
		$objTextBox.DropDownStyle= 'DropDownList'
		
		if ($objTextBox.Items.Count -le 1)
			{
			$objTextBox.Items.Add("Loading...")
			$objTextBox.SelectedIndex = 0
			$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
			$worksheet = $excel.sheets.item($global:DataSheet)
			$count=$worksheet.UsedRange.columns.count
			$Name=@()
			for($i=1; $i -le $count; $i++)
			{
			$Name += $worksheet.Cells.Item(1,$i).Text
			}
			$objTextBox.Items.AddRange($Name)
			$objTextBox.Items.Remove("Loading...")
			}
		
		$objLabel = New-Object System.Windows.Forms.Label
		$objLabel.Location = New-Object System.Drawing.Size(10,15) 
		$objLabel.Size = New-Object System.Drawing.Size(305,25) 
		$objLabel.Text = "Please select the name location:"
		
		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Size(85,120)
		$CancelButton.Size = New-Object System.Drawing.Size(75,23)
		$CancelButton.Text = "Cancel"
		$CancelButton.Add_Click({$ApplyRules.Close()})
		
		$OKButton = New-Object System.Windows.Forms.Button
		$OKButton.Location = New-Object System.Drawing.Size(10,120)
		$OKButton.Size = New-Object System.Drawing.Size(75,23)
		$OKButton.Text = "OK"
		$ApplyRules.AcceptButton = $OKButton
		$OKButton.Add_Click({

		$global:FillRow=$objTextBox.Text;
		
		
##----------------------------------------Run The Rules------------------------------------------##
function Get-ScriptDirectory {
    Split-Path -parent $PSCommandPath
}
$PathV=Get-ScriptDirectory
write-host $PathV
Function ExportWSToCSV ($excelFileName, $csvLoc)
{
	$String=$PathV.substring(0,$PathV.length-5)
	$Global:StringHere=$String+"Files\OriginalFileBackup"
	$xl = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
	$info=$xl.WorkBooks | Select-Object -Property name, path, author
	$Global:NameOfFile=$info.name
	$combo=$info.path+"\"+$info.name
	$wb = $xl.Workbooks.open($combo)
	$wb.SaveAs($StringHere)
    foreach ($ws in $wb.Worksheets)
    {
        $n = $ws.Name
		if ($n -eq $Global:DataSheet)
		{
        $ws.SaveAs($csvLoc+"\Temp\"+$n, 6)
		}
    }
	$wb.close()
	$xl.quit()
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
	[GC]::Collect()
}
ExportWSToCSV -csvLoc $PathV

#LOAD CSV#
$RealPath=$PathV+"\Temp\"+$global:DataSheet+".csv"
$A = Import-Csv $RealPath
$Global:NameLocation=$objTextBox.Text
$Global:NameLocation | export-clixml -Path ./Files/NameLocation.clixml

$StartMs = (Get-Date)
foreach ($line in $A)
{
		for ($j=$Global:RuleListArray.count-1;$j -ge 0; $j--)
		{
		if($Global:RuleListArray[$j]."Match?" -eq $True)
		{
		if($Global:RuleListArray.Matching[$j] -ne $null)
		{
		$MatchName=$Global:RuleListArray.MatchColumnName[$j]
		$ValueOfMatching = $Global:RuleListArray.Matching[$j].split(",")
		$Count=$ValueOfMatching.Count
		}
		if($Global:RuleListArray[$j]."Alphabetize?" -eq $True)
		{
		#MATCH AND ABC
		$MatchABCName=$Global:RuleListArray.ABCColumnName[$j]
		$FirstLetter=$Global:RuleListArray.FirstLetter[$j]
		$SecondLetter=$Global:RuleListArray.SecondLetter[$j]
		$Range = "["+"$FirstLetter"+"-"+"$SecondLetter"+"]"
		if($line.$MatchABCName[0][0] -match "$Range")
		{
		for ($m=0; $m -lt $Count; $m++)
		{
		if($line.$MatchName -like $ValueOfMatching[$m])
		{
		$line.$NameLocation = $Global:RuleListArray.AssignedName[$j]
		}
		}
		}
		#MATCH AND ABC
		}
		else
		{
		#MATCH ONLY
		for ($m=0; $m -lt $Count; $m++)
		{
		if($line.$MatchName -like $ValueOfMatching[$m])
		{
		$line.$NameLocation = $Global:RuleListArray.AssignedName[$j]
		}
		}
		#MATCH ONLY
		}
		
		}
		else
		{
		#ABC ONLY
		if($Global:RuleListArray[$j]."Alphabetize?" -eq $True)
		{
		$MatchABCName=$Global:RuleListArray.ABCColumnName[$j]
		$FirstLetter=$Global:RuleListArray.FirstLetter[$j]
		$SecondLetter=$Global:RuleListArray.SecondLetter[$j]
		$Range = "["+"$FirstLetter"+"-"+"$SecondLetter"+"]"
		if($line.$MatchABCName[0][0] -match "$Range")
		{
		$line.$NameLocation = $Global:RuleListArray.AssignedName[$j]
		}
		}
		#ABC ONLY
		else
		{
		$line.$NameLocation = ""
		}
		}
		}
	}
	
		$EndMs = (Get-Date)
		Write-Host "This script took $($EndMs - $StartMs) seconds to run" 
		$A | Export-Csv $RealPath -NoTypeInformation
		
$csvFiles = Get-ChildItem ($PathV+"\Temp\")
$Excel = New-Object -ComObject "excel.application"
$workbook = $excel.Workbooks.open($Global:StringHere)
$Excel.visible = $false
$Excel.DisplayAlerts = $false

$TempExcel=$PathV+"\Temp\"+$global:DataSheet
$workbook.sheets.item($global:DataSheet).delete()
$SOURCE=$Excel.workbooks.open($TempExcel)
$worksheet = $workbook.sheets.item(1)
$sheetToCopy = $SOURCE.sheets.item(1) # source sheet to copy
$sheetToCopy.copy($worksheet) 		  # copy source sheet to destination workbook
$SOURCE.close($false)

#$ExcelCount=$Excel.Worksheets.Count
$String=$PathV.substring(0,$PathV.length-5)
$Global:String2=$String+$NameOfFile
$Global:String2 | export-clixml -Path ./Files/Location.clixml


	$workbook.SaveAs($String2)
	$workbook.close($false)
	$Excel.quit()
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($SOURCE)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheetToCopy)
	[GC]::Collect()

		$ApplyRules.Close();
		})
		$ApplyRules.Icon = $Icon
		$ApplyRules.Controls.Add($CancelButton)
		$ApplyRules.Controls.Add($objTextBox) 
		$ApplyRules.Controls.Add($objLabel) 
		$ApplyRules.Controls.Add($OKButton)
		$ApplyRules.Add_Shown({$ApplyRules.Activate();$objTextBox.focus()})
		[void] $ApplyRules.ShowDialog()
	})
	
	#Sort Tabs
	$Button_Task7.Add_Click({
	
		$Excel = New-Object -ComObject Excel.Application
		$Workbook = $Excel.Workbooks.Open($String2)
		$worksheet = $workbook.sheets.item($global:DataSheet)
		$Excel.visible=$false
		
		$countRows=$worksheet.UsedRange.rows.count
		$Names = $Global:RuleListArray.AssignedName | select -uniq
		$Global:LengthNames=$Names.count
	
		$SortTabs = New-Object System.Windows.Forms.Form 
		$SortTabs.Text = "TANYR"
		$SortTabs.Size = New-Object System.Drawing.Size(196,400) 
		$SortTabs.StartPosition = "CenterScreen"
		$SortTabs.Topmost = $True
		$SortTabs.Font = $Font
		$SortTabs.FormBorderStyle = 'FixedDialog'

		$objTextBox1 = New-Object System.Windows.Forms.ComboBox
		$objTextBox1.Location = New-Object System.Drawing.Size(10,30) 
		$objTextBox1.Size = New-Object System.Drawing.Size(150,20) 
		$objTextBox1.DropDownStyle= 'DropDownList'
		
		$objTextBox2 = New-Object System.Windows.Forms.ComboBox
		$objTextBox2.Location = New-Object System.Drawing.Size(10,85) 
		$objTextBox2.Size = New-Object System.Drawing.Size(150,20) 
		$objTextBox2.DropDownStyle= 'DropDownList'
		
		$TabMatchList = New-Object System.Windows.Forms.ListBox
		$TabMatchList.Location = New-Object System.Drawing.Size(10,110) 
		$TabMatchList.Size = New-Object System.Drawing.Size(148,215)
		
		if($Global:TabRuleArray -ne $null)
		{
		$TabMatchList.Items.AddRange($Global:TabRuleArray)
		}
		
		if ($objTextBox1.Items.Count -le 1)
			{
			$objTextBox1.Items.Add("Loading...")
			$objTextBox2.Items.Add("Loading...")
			$objTextBox1.SelectedIndex = 0
			$objTextBox2.SelectedIndex = 0
			$count=$worksheet.UsedRange.columns.count
			$Name=@()
			for($i=1; $i -le $count; $i++)
			{
			$Name += $worksheet.Cells.Item(1,$i).Text
			}
			$objTextBox1.Items.AddRange($Name)
			$objTextBox2.Items.AddRange($Name)
			$objTextBox1.Items.Remove("Loading...")
			$objTextBox2.Items.Remove("Loading...")
			}
			
			$AddButton2 = New-Object System.Windows.Forms.Button
			$AddButton2.Location = New-Object System.Drawing.Size(158,113) 
			$AddButton2.Size = New-Object System.Drawing.Size(20,20)
			$AddButton2.Text = "+"
			$AddButton2.Add_Click({
			
			If ($TabMatchList.Items -contains $objTextBox2.SelectedItem)
			{
			}
			else
			{
			$TabMatchList.Items.Add($objTextBox2.SelectedItem)
			$Global:TabRuleArray.Add($objTextBox2.SelectedItem)
			$Global:TabRuleArray | export-clixml -Path ./Files/TabRules.clixml
			}
			})
			$RemoveButton2 = New-Object System.Windows.Forms.Button
			$RemoveButton2.Location = New-Object System.Drawing.Size(158,133) 
			$RemoveButton2.Size = New-Object System.Drawing.Size(20,20)
			$RemoveButton2.Text = "-"
			$RemoveButton2.Add_Click({
			$Global:TabRuleArray.RemoveAt($TabMatchList.SelectedIndex)
			$TabMatchList.Items.RemoveAt($TabMatchList.SelectedIndex)
			$Global:TabRuleArray | export-clixml -Path ./Files/TabRules.clixml
			})
			
			
		
		$objLabel = New-Object System.Windows.Forms.Label
		$objLabel.Location = New-Object System.Drawing.Size(10,15) 
		$objLabel.Size = New-Object System.Drawing.Size(305,25) 
		$objLabel.Text = "Please select the name location:"
		
		$objLabel2 = New-Object System.Windows.Forms.Label
		$objLabel2.Location = New-Object System.Drawing.Size(10,70) 
		$objLabel2.Size = New-Object System.Drawing.Size(305,25) 
		$objLabel2.Text = "Add Columns To Copy Over:"
		
		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Size(85,330)
		$CancelButton.Size = New-Object System.Drawing.Size(75,23)
		$CancelButton.Text = "Cancel"
		$CancelButton.Add_Click({$SortTabs.Close()})
		
		$MoveUp = New-Object System.Windows.Forms.Button
		$MoveUp.Location = New-Object System.Drawing.Size(158,160) 
		$MoveUp.Size = New-Object System.Drawing.Size(20,20)
		$MoveUp.Text = "/\"
		$MoveUp.Add_Click({
		 # only if the first item isn't the current one
    if($TabMatchList.SelectedIndex -gt 0)
    {
        $TabMatchList.BeginUpdate()
        #Get starting position
        $pos = $TabMatchList.selectedIndex
        # add a duplicate of original item up in the listbox
        $TabMatchList.items.insert($pos -1,$TabMatchList.Items.Item($pos))
		$Global:TabRuleArray[$pos],$Global:TabRuleArray[$pos-1] = $Global:TabRuleArray[$pos-1],$Global:TabRuleArray[$pos]
        # make it the current item
        $TabMatchList.SelectedIndex = ($pos -1)
        # delete the old occurrence of this item
        $TabMatchList.Items.RemoveAt($pos +1)
        $TabMatchList.EndUpdate()
		$Global:TabRuleArray | export-clixml -Path ./Files/TabRules.clixml
    }
	ELSE
	{}
	})
	
		$MoveDown = New-Object System.Windows.Forms.Button
		$MoveDown.Location = New-Object System.Drawing.Size(158,180) 
		$MoveDown.Size = New-Object System.Drawing.Size(20,20)
		$MoveDown.Text = "\/"
		$MoveDown.Add_Click({
		
		if(($TabMatchList.SelectedIndex -ne -1)   -and   ($TabMatchList.SelectedIndex -lt $TabMatchList.Items.Count - 1)    )   {
        $TabMatchList.BeginUpdate()
        #Get starting position 
        $pos = $TabMatchList.selectedIndex
        # add a duplicate of item below in the listbox
        $TabMatchList.items.insert($pos,$TabMatchList.Items.Item($pos +1))
		$Global:TabRuleArray[$pos],$Global:TabRuleArray[$pos+1] = $Global:TabRuleArray[$pos+1],$Global:TabRuleArray[$pos]
        # delete the old occurrence of this item
        $TabMatchList.Items.RemoveAt($pos +2)
        # move to current item
        $TabMatchList.SelectedIndex = ($pos +1)
        $TabMatchList.EndUpdate()
		$Global:TabRuleArray | export-clixml -Path ./Files/TabRules.clixml 
   }
   ELSE
   {}
   })

		$OKButton = New-Object System.Windows.Forms.Button
		$OKButton.Location = New-Object System.Drawing.Size(10,330)
		$OKButton.Size = New-Object System.Drawing.Size(75,23)
		$OKButton.Text = "OK"
		$SortTabs.AcceptButton = $OKButton
		$OKButton.Add_Click({
		
		$Global:LocationText=$objTextBox1.text
		$WorkSheet2 = $Excel.WorkSheets
		$Global:Names = $Global:RuleListArray.AssignedName | select -uniq
		$Global:LengthNames=$Global:Names.count
	
		write-host $Global:LocationText -foregroundcolor "Blue"
		$i=1;$j=1
		
		function Get-ScriptDirectory 
		{
		Split-Path -parent $PSCommandPath
		}
		$PathV=Get-ScriptDirectory

		$RealPath=$PathV+"\Temp\"+$global:DataSheet+".csv"
		$String=$PathV.substring(0,$PathV.length-5)
		$String=$String+"Output\"
		foreach($PersonName in $Global:Names)
		{
		$NamePath=$String+$PersonName+".csv"
		write-host $NamePath
		$A = Import-Csv $RealPath 
		$B = $A | Where-Object {$_.$Global:NameLocation -eq $PersonName}
		$C = $B | Select -Property $Global:TabRuleArray

		if ($C -ne $null)
		{
		$C | Export-Csv $NamePath -NoTypeInformation
		}
		}
		$workbook.close($false)
		$Excel.quit()
		
		$String=$PathV.substring(0,$PathV.length-5)
		$Global:StringHere=$String+$NameOfFile
		$csvFiles = Get-ChildItem ($String+"\Output\")
		$Excel = New-Object -ComObject "excel.application"
		$workbook = $excel.Workbooks.open($Global:StringHere)
		$Excel.visible = $false
		$Excel.DisplayAlerts = $false

foreach($file in $csvFiles)
{
$SOURCE=$Excel.workbooks.open($file.fullname)
$worksheet = $workbook.sheets.item(1)
$sheetToCopy = $SOURCE.sheets.item(1) # source sheet to copy
$sheetToCopy.copy($worksheet) 		  # copy source sheet to destination workbook
$SOURCE.close($false)
}
$ExcelCount=$Excel.Worksheets.Count

	$workbook.SaveAs($String2)
	$workbook.close($false)
	$Excel.quit()
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($SOURCE)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheetToCopy)
	[GC]::Collect()
		
		$RemoveItemHere=$String+"\Output\*"
		Remove-Item $RemoveItemHere
		$SortTabs.Close()
		})
	
		$SortTabs.Icon = $Icon
		$SortTabs.Controls.Add($CancelButton)
		$SortTabs.Controls.Add($objTextBox1) 
		$SortTabs.Controls.Add($objTextBox2) 
		$SortTabs.Controls.Add($objLabel) 
		$SortTabs.Controls.Add($objLabel2) 
		$SortTabs.Controls.Add($OKButton)
		$SortTabs.Controls.Add($TabMatchList)
		$SortTabs.Controls.Add($AddButton2)
		$SortTabs.Controls.Add($RemoveButton2)
		$SortTabs.Controls.Add($MoveUp)
		$SortTabs.Controls.Add($MoveDown)
		$SortTabs.Add_Shown({$SortTabs.Activate();$objTextBox2.focus()})
		[void] $SortTabs.ShowDialog()
	
	})
	

[void]$TANYRHealthcare.ShowDialog()
$TANYRHealthcare.Dispose()