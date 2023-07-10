Add-Type -AssemblyName System.Windows.Forms
$Icon = New-Object system.drawing.icon (".\Files\Logo.ico")
$Font = New-Object System.Drawing.Font("Times New Roman",9)

$TANYRHealthcare = New-Object system.Windows.Forms.Form
$TANYRHealthcare.Text = "TANYR"
$TANYRHealthcare.TopMost = $false
$TANYRHealthcare.Width = 230
$TANYRHealthcare.Height = 205
$TANYRHealthcare.Icon = $Icon
$TANYRHealthcare.Font = $Font
$TANYRHealthcare.FormBorderStyle = 'FixedDialog'

##----------------------------MAIN PAGE GUI----------------------------##
#Button - Task 1
$Button_Task1 = New-Object system.windows.Forms.Button
$Button_Task1.Text = "Select Date"
$Button_Task1.Width = 126
$Button_Task1.Height = 35
$Button_Task1.location = new-object system.drawing.point(10,13)
$Button_Task1.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task1)
$TANYRHealthcare.AcceptButton = $Button_Task1


#Button - Task 3
$Button_Task3 = New-Object system.windows.Forms.Button
$Button_Task3.Text = "Load Names"
$Button_Task3.Width = 126
$Button_Task3.Height = 35
$Button_Task3.location = new-object system.drawing.point(10,60)
$Button_Task3.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task3)

$Button_Task5 = New-Object system.windows.Forms.Button
$Button_Task5.Text = "Download PDFs"
$Button_Task5.Width = 126
$Button_Task5.Height = 35
$Button_Task5.location = new-object system.drawing.point(10,107)
$Button_Task5.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task5)

$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Location = New-Object System.Drawing.Size(145,19) 
$objTextBox1.Size = New-Object System.Drawing.Size(60,65) 
$TANYRHealthcare.Controls.Add($objTextBox1) 

$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.Location = New-Object System.Drawing.Size(145,66) 
$objTextBox2.Size = New-Object System.Drawing.Size(60,65) 
$TANYRHealthcare.Controls.Add($objTextBox2) 

$objTextBox3 = New-Object System.Windows.Forms.TextBox 
$objTextBox3.Location = New-Object System.Drawing.Size(145,113) 
$objTextBox3.Size = New-Object System.Drawing.Size(60,65) 
$TANYRHealthcare.Controls.Add($objTextBox3)

$checkBox = New-Object System.Windows.Forms.CheckBox 
$checkBox.Text="Show Browser?" 
$checkBox.Location = New-Object System.Drawing.Size(10,142) 
$checkBox.Size = New-Object System.Drawing.Size(105,23) 
$TANYRHealthcare.Controls.Add($checkBox) 
     

##----------------------------MAIN PAGE COMMANDS----------------------------##

    #Select Date
    $Button_Task1.Add_Click({$objForm = New-Object Windows.Forms.Form 

$objForm.Text = "Select a Date" 
$objForm.Size = New-Object Drawing.Size @(190,190) 
$objForm.StartPosition = "CenterScreen"
$objTextBox1.text="Loading..."

$objForm.KeyPreview = $True

$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
        {
            $dtmDate=$objCalendar.SelectionStart
			$dtmDate = $dtmDate.ToShortDateString()
			$global:dtmDate = "0"+$dtmDate
			
			if ($global:dtmDate[1] -eq "1")
			{
			if ($global:dtmDate[2] -eq "0" -or "1" -or "2")
			{
			$global:dtmDate=$global:dtmDate.TrimStart("0")
			}
			}
            $objForm.Close()
        }
    })

$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
        {
            $objForm.Close()
        }
    })

$objCalendar = New-Object System.Windows.Forms.MonthCalendar 
$objCalendar.ShowTodayCircle = $False
$objCalendar.MaxSelectionCount = 1
$objForm.Controls.Add($objCalendar) 

$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})  
[void] $objForm.ShowDialog() 

if ($dtmDate)
    {
        Write-Host "Date selected: $dtmDate"
    }
	$objTextBox1.text="Done!"
  }
)

    #Load Names
    $Button_Task3.Add_Click({
	$objTextBox2.text="Loading..."
	$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
	#$excel.Visible=$false
	$rowDate,$colDate = 1,1
	$rowName,$colName = 1,3
	$rowCredit,$colCredit = 1,5
	
	$global:Name=@()
	$global:Credit=@()
	
	$j=0
	$CheckIfFalse=$false
	for ($i=1; $i -le 200; $i++)
	{
	if ($excel.Cells.Item($rowDate+$i,$colDate).text -eq $global:dtmDate)
	
	{
	$CheckIfFalse=$true
	if([string]::IsNullOrEmpty($excel.Cells.Item($rowCredit+$i,$colCredit).text))
	{
	}
	else
	{
	$global:Name += $excel.Cells.Item($rowName+$i,$colName).text
	$global:Credit += $excel.Cells.Item($rowCredit+$i,$colCredit).text
	#Write-Host $global:dtmDate -Separator `n -foregroundcolor "red"
	#Write-Host $Name -Separator `n -foregroundcolor "green"
	#Write-Host $Credit -Separator `n -foregroundcolor "blue"
	}
	}
	else 
	{
	#Write-Host "nope" -Separator `n -foregroundcolor "red"
	if ($CheckIfFalse -eq $true)
	{
	break
	}
	}
	#End For
	}
	#End Button Click
	$objTextBox2.text="Done!"
	}
)	

    #Download PDFs
    $Button_Task5.Add_Click({
	$objTextBox3.text="Loading..."
	$Username="awhiteheadOO"
	$Password="#Sooner17"
	$ie = New-Object -ComObject 'InternetExplorer.Application'
	$ie.AddressBar = $false
	$ie.MenuBar = $false
	$ie.StatusBar = $false
	$ie.ToolBar = $false
	$ie.Visible = $false
	if ($checkBox.Checked)
	{
	$ie.Visible = $true
	}
	$ie.navigate("https://login.zirmed.com/ui")
	
	While ($ie.Busy)
{
    Sleep -Milliseconds 100
}

	$Document=$ie.document
	Sleep -Milliseconds 1000
	
	$UsernameBox=$Document.IHTMLDocument3_getElementsByTagName("input") | where-object {$_.name -eq "loginName"}
	$UsernameBox.value="$Username"

	$PasswordBox=$Document.IHTMLDocument3_getElementsByTagName("input") | where-object {$_.name -eq "password"}
	$PasswordBox.value="$Password"

    Sleep -Milliseconds 1000
	
	$LoginButton=$Document.IHTMLDocument3_getElementById("login-button")
	$LoginButton.click()

While ($ie.Busy)
{
    Sleep -Milliseconds 100
}
    Sleep -Milliseconds 500
$Catchpa=$Document.IHTMLDocument3_getElementsByTagName("strong") | where-object {$_.innerText-eq " What was your first car? "}
$Catchpa2=$Document.IHTMLDocument3_getElementsByTagName("strong") | where-object {$_.innerText -eq " What was your high school mascot? "}
    Sleep -Milliseconds 500

if ($Catchpa -ne $null)
{
$verifyAnswer=$Document.IHTMLDocument3_getElementById("verifyAnswer")
$verifyAnswer.value="Camero"
$verifyAnswer=$Document.IHTMLDocument3_getElementById("VerifyButton")
$verifyAnswer.click()
}

if ($Catchpa2 -ne $null)
{
$verifyAnswer=$Document.IHTMLDocument3_getElementById("verifyAnswer")
$verifyAnswer.value="hawks"
$verifyAnswer=$Document.IHTMLDocument3_getElementById("VerifyButton")
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
	$global:AmountBox=$Document.IHTMLDocument3_getElementById("txtAmount")
	}
	$global:AmountBox.value=$global:Credit[$i]
	
	if($global:DateOne -eq $null){
	$global:DateOne=$Document.IHTMLDocument3_getElementsByTagName("input") | where-object {$_.id -like "dp*"}
	$global:DateOne[0].value=$global:dtmDate
	$global:DateOne[1].value=$global:dtmDate
	}
	
	
	if($SearchButton -eq $null){
	$global:SearchButton=$Document.IHTMLDocument3_getElementById("btnSearch")
	}
	$global:SearchButton.click()
	
	
	sleep -milliseconds 1000
	
	$viewEOB=$Document.IHTMLDocument3_getElementById("viewEOB")
	if($viewEOB -eq $null)
	{
	$global:NotFound += $global:Name[$i]
	Continue
	}
	$viewEOB.click()
    
    sleep -milliseconds 2000

  	$global:PopOut=$Document.IHTMLDocument3_getElementsByTagName("span") | where-object {$_.innerText -eq "popout"}
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
}

}
)

[void]$TANYRHealthcare.ShowDialog()
$TANYRHealthcare.Dispose()