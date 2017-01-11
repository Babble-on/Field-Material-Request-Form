######################    S Bonner                            #########################################
#######################################################################################################
#######################################################################################################
######################                                        #########################################
######################     Field Material Request Form        #########################################
######################                                        #########################################
#######################################################################################################
#######################################################################################################
#######################################################################################################
######################       Variables that will change       #########################################
#######################################################################################################
#######################################################################################################
$data = "path to Field.xml"
$u = "path to itextsharp-all-5.5.10(1)\itextsharp-dll-core"
$Global:Template = "path to Template.pdf"
###################################################################################################
#################################      EMail Lists     ############################################
[String[]]$Global:RDG = "array of email addresses"
[String[]]$Global:Bloom ="array of email addresses"
   
#=====================================================================================================#
#                                XML Data Providers                                                   #
#=====================================================================================================#
$field = (Get-Content "$data")
$field = [xml](Get-Content "$data")
$Groups = $xml.Root.InventoryGroups.ID
$SSM = $field.Root.InventoryItem | Where {$_.ItemList -eq "1"}
$SSMM = $field.Root.InventoryItem | where {$_.ItemList -eq "2"}
$PM = $Field.Root.InventoryItem | where {$_.ItemList -eq "3"}
$PS = $Field.Root.InventoryItem | where {$_.ItemList -eq "4"}
$smelly = $field.Root.InventoryGroups.ID | Select-Object -Property name

#=====================================================================================================#
#                                GEN AD Creds                                                         #
#=====================================================================================================#



function GenPW()
	{
		for ($x = 0;$x -le $core.length;$x++)
			{
				if (($x % 2) -ne 0)
					{
						$global:pw+=$core[$x]
					}
			}
	}


Function GenUSer()
	{
		for ($x = 0;$x -le $pore.length;$x++)
			{
				if (($x % 2) -ne 0)
					{
						$global:US+=$pore[$x]
					}
			}
	}
	
$global:pw = ""
$global:US= ""
$core = "secure password.  inquire about how this works"
$pore = "secure user.  inquire about how this works"
GenPW
GENUSer
$password = convertto-securestring $global:pw -asplaintext -force
$cred = New-Object System.Management.Automation.PsCredential("domain01\$Global:US",$password)

#=====================================================================================================#
#                                Add PDF Creation Library                                             #
#=====================================================================================================#
[System.Reflection.Assembly]::LoadFrom("$U\itextsharp.dll")
#====================================================================================================#
#                                      Build PDF                                                     #
#====================================================================================================#
$Global:Location = "path to app location"
$Global:date= (get-date)
$Global:docname= "Material-Request"
$Global:pdf = New-Object iTextSharp.text.Document
[System.Reflection.Assembly]::LoadFrom("$U\itextsharp.dll")
function global:Get-PdfFieldNames
{
	[OutputType([string])]
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[ValidatePattern('\.pdf$')]
		[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
		[string]$FilePath,
		
		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[ValidatePattern('\.dll$')]
		[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
		[string]$ITextLibraryPath = (Find-ITextSharpLibrary).FullName
	)
	begin
	{
		$ErrorActionPreference = 'Stop'
		## Load the iTextSharp DLL to do all the heavy-lifting 
		[System.Reflection.Assembly]::LoadFrom($ITextLibraryPath) | Out-Null
	}
	process
	{
		try
		{
			$reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $FilePath
			$reader.AcroFields.Fields.Key
		}
		catch
		{
			$PSCmdlet.ThrowTerminatingError($_)
		}
	}
}

function global:Save-PdfField
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[Hashtable]$Fields,
		
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[ValidatePattern('\.pdf$')]
		[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
		[string]$InputPdfFilePath,
		
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[ValidatePattern('\.pdf$')]
		[ValidateScript({ -not (Test-Path -Path $_ -PathType Leaf) })]
		[string]$OutputPdfFilePath,
		
		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[ValidatePattern('\.dll$')]
		[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
		[string]$ITextSharpLibrary = (Find-ITextSharpLibrary).FullName
		
	)
	begin
	{
		$ErrorActionPreference = 'Stop'
	}
	process
	{
		try
		{
			$reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $InputPdfFilePath
			$stamper = New-Object iTextSharp.text.pdf.PdfStamper($reader, [System.IO.File]::Create($OutputPdfFilePath))
			
			## Apply all hash table elements into the PDF form
			foreach ($j in $Fields.GetEnumerator())
			{
				$null = $stamper.AcroFields.SetField($j.Key, $j.Value)
			}
		}
		catch
		{
			$PSCmdlet.ThrowTerminatingError($_)
		}
		finally
		{
			## Close up shop 
			$stamper.Close()
			Get-Item -Path $OutputPdfFilePath
		}
	}
}

function global:Generate-PDF{
Get-PdfFieldNames -FilePath "$template" -ITextLibraryPath "$U\itextsharp.dll"
$Global:PDF_Territory= $Global:WPF_Territory.SelectedItem.Content
$Global:PDF_Date= $Global:date
$Global:PDF_Project= $Global:WPF_Enter_PRJ.Text
$Global:PDF_WR= $Global:WPF_Enter_WR.Text
$global:PDF_Additional= $Wpf_Comments.Text
$Global:PDF_STOCK= $Global:ID[0]
$Global:PDF_STOCK0= $Global:ID[1]
$Global:PDF_STOCK1= $ID[2]
$Global:PDF_STOCK2= $ID[3]
$Global:PDF_STOCK3= $ID[4]
$Global:PDF_STOCK4= $ID[5]
$Global:PDF_STOCK5= $ID[6]
$Global:PDF_STOCK6= $ID[7]
$Global:PDF_STOCK7= $ID[8]
$Global:PDF_STOCK8= $ID[9]
$Global:PDF_DES= $Global:Description[0]
$Global:PDF_DES0= $Description[1]
$Global:PDF_DES1= $Description[2]
$Global:PDF_DES2= $Description[3]
$Global:PDF_DES3= $Description[4]
$Global:PDF_DES4= $Description[5]
$Global:PDF_DES5= $Description[6]
$Global:PDF_DES6= $Description[7]
$Global:PDF_DES7= $Description[8]
$Global:PDF_DES8= $Description[9]
$Global:PDF_UOM= $Global:UOM[0]
$Global:PDF_UOM0= $UOM[1]
$Global:PDF_UOM1= $UOM[2]
$Global:PDF_UOM2= $UOM[3]
$Global:PDF_UOM3= $UOM[4]
$Global:PDF_UOM4= $UOM[5]
$Global:PDF_UOM5= $UOM[6]
$Global:PDF_UOM6= $UOM[7]
$Global:PDF_UOM7= $UOM[8]
$Global:PDF_UOM8= $UOM[9]
$Global:PDF_QTY= [String]$Global:Req[0]
$Global:PDF_QTY0= [String]$Req[1]
$Global:PDF_QTY1= [String]$Req[2]
$Global:PDF_QTY2= [String]$Req[3]
$Global:PDF_QTY3= [String]$Req[4]
$Global:PDF_QTY4= [String]$Req[5]
$Global:PDF_QTY5= [String]$Req[6]
$Global:PDF_QTY6= [String]$Req[7]
$Global:PDF_QTY7= [String]$Req[8]
$Global:PDF_QTY8= [String]$Req[9]
Save-PdfField -Fields @{'W R' = "$Global:PDF_WR";
'Territory' = "$Global:PDF_Territory";
'Project Name' = "$Global:PDF_Project";
'Stock ID' = "$Global:PDF_STOCK";
'Stock ID-0' = "$Global:PDF_STOCK0";
'Stock ID-1' = "$Global:PDF_STOCK1";
'Stock ID-2' = "$Global:PDF_STOCK2";
'Stock ID-3' = "$Global:PDF_STOCK3";
'Stock ID-4' = "$Global:PDF_STOCK4";
'Stock ID-5' = "$Global:PDF_STOCK5";
'Stock ID-6' = "$Global:PDF_STOCK6";
'Stock ID-7' = "$Global:PDF_STOCK7";
'Stock ID-8' = "$Global:PDF_STOCK8";
'Description' = "$Global:PDF_DES";
'Description-0' = "$Global:PDF_DES0";
'Description-1' = "$Global:PDF_DES1";
'Description-2' = "$Global:PDF_DES2";
'Description-3' = "$Global:PDF_DES3";
'Description-4' = "$Global:PDF_DES4";
'Description-5' = "$Global:PDF_DES5";
'Description-6' = "$Global:PDF_DES6";
'Description-7' = "$Global:PDF_DES7";
'Description-8' = "$Global:PDF_DES8";
'UOM' = "$Global:PDF_UOM";
'UOM-0' = "$PDF_UOM0";
'UOM-1' = "$PDF_UOM1";
'UOM-2' = "$PDF_UOM2";
'UOM-3' = "$PDF_UOM3";
'UOM-4' = "$PDF_UOM4";
'UOM-5' = "$PDF_UOM5";
'UOM-6' = "$PDF_UOM6";
'UOM-7' = "$PDF_UOM7";
'UOM-8' = "$PDF_UOM8";
'Request' = "$Global:PDF_QTY";
'Request-0' = "$PDF_QTY0";
'Request-1' = "$PDF_QTY1";
'Request-2' = "$PDF_QTY2";
'Request-3' = "$PDF_QTY3";
'Request-4' = "$PDF_QTY4";
'Request-5' = "$PDF_QTY5";
'Request-6' = "$PDF_QTY6";
'Request-7' = "$PDF_QTY7";
'Request-8' = "$PDF_QTY8";
} -InputPdfFilePath "$template" -ITextSharpLibrary "$U\itextsharp.dll" -OutputPdfFilePath 'path to output.pdf'
}

Function service-area (){
if ($WPF_Territory.SelectedItem.content -eq "Reading"){$Global:EmailTo = $Global:RDG} else
{$Global:EmailTo = $Global:Bloom}

}


function Email-form (){
$Global:username= $wpf_Enter_name.text
$Global:password= $WPF_Enter_pw.Password
$EmailFrom = ((Get-ADUser $Global:username -properties mail -Credential $cred).mail)
$EmailPDF = "path to output.pdf"
$Subject = "Field Material Request"
$body = "This is a Field Material Request from $Global:Username. The request was sent $global:Date  The request was sent to $emailTo. Additional Comments: $PDF_Additional"
$Secpasswd = ConvertTo-SecureString $Global:password -AsPlainText -Force
$MyCreds = New-Object System.Management.Automation.PSCredential ("domain01\$Global:username",$Secpasswd)
Send-MailMessage –From "$emailFrom" –To $Global:emailto –Subject $subject -Body $Body -Attachments $EmailPDF –SMTPServer "path to smtp.server" –Credential $MyCreds -DeliveryNotificationOption OnFailure
Clean-up
}

Function Clean-up (){
##rename pdf to old, move it 
$date = (Get-Date -Format dddMMMyyyyhhmm)  
$path =  "path to pdf directory"    
if (test-path $path\Field_Request.pdf){Move-Item -Path $path\Field_Request.pdf -Destination $path\old}
if (test-path $path\old\Field_Request.pdf){rename-Item -Path $path\OLD\Field_Request.pdf -NewName "$date.pdf"}
### test if pdf exist if not say message sent
if (!(test-path -Path "path to output.pdf")){[System.Windows.Forms.MessageBox]::Show("Your email has been sent. Now closing. ")}
#### clean memory
start-sleep -Seconds 2
#### close app
$Form.close()
}

#=====================================================================================================#
#                                           Warning Form                                              #
#=====================================================================================================#
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$MaterialRequestForm = New-Object System.Windows.Forms.Form 
$MaterialRequestForm.Text = "Warning"
$MaterialRequestForm.Size = New-Object System.Drawing.Size(450,300) 
$MaterialRequestForm.StartPosition = "CenterScreen"
$MaterialRequestForm.AutoScroll = $True
$MaterialRequestForm.ShowInTaskbar = $True
$MaterialRequestForm.ControlBox= $False
$MaterialSelectLabel = New-Object System.Windows.Forms.Label
$MaterialSelectLabel.Location = New-Object System.Drawing.Size(2,20) 
$MaterialSelectLabel.Size = New-Object System.Drawing.Size(430,170) 
$MaterialSelectLabel.Text = "You can not order more then 10 items at a time. Please either clear your order and start over or submit additional orders as needed."
$MaterialSelectLabelFont = New-Object System.Drawing.Font("Georgia", 18,[System.Drawing.FontStyle]::Bold)
$MaterialSelectLabel.Font = $MaterialSelectLabelFont
$MaterialRequestForm.Controls.Add($MaterialSelectLabel)
$MaterialSelectOKButton = New-Object System.Windows.Forms.Button
$MaterialSelectOKButton.Location = New-Object System.Drawing.Size(180,200)
$MaterialSelectOKButton.Size = New-Object System.Drawing.Size(75,30)
$MaterialSelectOKButton.Text = "Okay"
$MaterialSelectOKButton.Add_Click({$MaterialRequestForm.Close()})
$MaterialRequestForm.Controls.Add($MaterialSelectOKButton)
$MaterialRequestForm.Topmost = $True
$MaterialRequestForm.Add_Shown({$MaterialRequestForm.Activate()})
#=====================================================================================================#
#                                           Functions                                                 #
#=====================================================================================================#
##########################################################################################################

function checkNull (){
   if ($Global:Amount -and ($Global:integer = [int]$Global:Amount)){
   AddToArray
   BuildTable
   }else {
   [System.Windows.Forms.MessageBox]::Show("Invalid Amount Requested. Please enter a Valid amount.")
   } 
}

Function Global:Test-ADAuthentication {
    param($Global:username,$Global:password)
    (new-object directoryservices.directoryentry "",$Global:username,$Global:password).psbase.name -ne $null
}

$Global:UserName= ""
$Global:Password= ""
###$Global:Domain = $env:USERDOMAIN
$Global:Domain = "company domain"
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$Global:ct = [System.DirectoryServices.AccountManagement.ContextType]::Domain
$Global:pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $Global:ct,$Global:Domain

#################################################################
######################################################################

function FillOutForm (){
    $Global:Name= $Wpf_Enter_Name.text
    $Global:Project= $WPF_Enter_PRJ.text
    $Global:Enter_WR= $WPF_Enter_WR.text
    $GLobal:Enter_Service= $WPF_Territory.SelectedItem
    $Global:Comments= $WPF_Comments.text
    $Global:username = $WPF_Enter_Name.Text
    $Global:password= $WPF_Enter_PW.Password
    $Global:Check = ($Global:pc.ValidateCredentials($Global:UserName,$Global:Password))
   if ($Global:Name -and $Global:Project -and ($Global:WR = [int]$Global:Enter_WR )-and $GLobal:Enter_Service -and $Global:ID -and ($Global:check -eq $True)){
   global:Generate-PDF; service-area; Email-form
   }else {
   $Global:UserName= $null
   $Global:Password= $null
   $WPF_Enter_Name.Clear()
   $WPF_Enter_PW.Clear()
   [System.Windows.Forms.MessageBox]::Show("All Fields except Additional Comments are required prior to sending your order. Either a Require Field is empty or you have an incorrect Username/password")   
   }
}

function global:AddToArray(){
if ($ID.Length -eq 10){
[void] $MaterialRequestForm.ShowDialog()
}Else {
    $Global:ID += ("$Global:Select_ID")
    $Global:Description += ("$Global:Selected")
    $Global:UOM += ("$Global:Select_UOM")
    $Global:Req += ("$Global:Amount")
}
}

function global:BuildTable (){
$WPF_Store1.Text = $ID[0]
$WPF_Store2.Text = $ID[1]
$WPF_Store3.Text = $ID[2]
$WPF_Store4.Text = $ID[3]
$WPF_Store5.Text = $ID[4]
$WPF_Store6.Text = $ID[5]
$WPF_Store7.Text = $ID[6]
$WPF_Store8.Text = $ID[7]
$WPF_Store9.Text = $ID[8]
$WPF_Store10.Text = $ID[9]
$WPF_Des1.Text= $Description[0]
$WPF_Des2.Text= $Description[1]
$WPF_Des3.Text= $Description[2]
$WPF_Des4.Text= $Description[3]
$WPF_Des5.Text= $Description[4]
$WPF_Des6.Text= $Description[5]
$WPF_Des7.Text= $Description[6]
$WPF_Des8.Text= $Description[7]
$WPF_Des9.Text= $Description[8]
$WPF_Des10.Text= $Description[9]
$WPF_UOM1.Text= $UOM[0]
$WPF_UOM2.Text= $UOM[1]
$WPF_UOM3.Text= $UOM[2]
$WPF_UOM4.Text= $UOM[3]
$WPF_UOM5.Text= $UOM[4]
$WPF_UOM6.Text= $UOM[5]
$WPF_UOM7.Text= $UOM[6]
$WPF_UOM8.Text= $UOM[7]
$WPF_UOM9.Text= $UOM[8]
$WPF_UOM10.Text= $UOM[9]
$WPF_Req1.Text= $Req[0]
$WPF_Req2.Text= $Req[1]
$WPF_Req3.Text= $Req[2]
$WPF_Req4.Text= $Req[3]
$WPF_Req5.Text= $Req[4]
$WPF_Req6.Text= $Req[5]
$WPF_Req7.Text= $Req[6]
$WPF_Req8.Text= $Req[7]
$WPF_Req9.Text= $Req[8]
$WPF_Req10.Text= $Req[9]
$Global:Select_ID= $null
$Global:Selected= $null
$Global:Select_UOM= $null
$Global:Amount= $null
$WPF_QTY_Select.Clear()
$WPF_UOM_Select.Clear()
$WPF_QTY.Clear()
}

$ID=@() 
$Description=@()
$UOM=@()
$Req=@()
$Select_ID= $null
$Selected= $null
$Select_UOM= $null
$Amount= $null

#######################################################################################################
#######################################################################################################
#######################################  WPF  GUI  ####################################################
#######################################################################################################
#====================================================================================================#
#                                      WPF GUI Start                                                 #
#====================================================================================================#
$inputXML = @"
<Window x:Class="WpfApplication1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:New_Start_Powershell"
        mc:Ignorable="d"
        Background="#b5c0d1"
        Title="Field Material Request" Height="700" Width="1280" ResizeMode="CanMinimize" WindowState="Maximized">
    <Grid>
        <Grid>
            <TabControl>
                <TabItem>
                    <TabItem.Header>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Find Parts to be Ordered" Width="414" Padding="70,0,0,0" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="20" FontWeight="Black" FontFamily="georgia"/>
                        </StackPanel>
                    </TabItem.Header>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="771*"/>
                            <ColumnDefinition Width="497*"/>
                        </Grid.ColumnDefinitions>
                        <Label Content="Select Material Type:" Padding="60,20,0,0" FontSize="26" FontFamily="georgia" FontWeight="Black" Grid.ColumnSpan="2"/>
                        <ComboBox x:Name="comboBox" HorizontalAlignment="Left" FontSize="18" Margin="30,60,0,0" VerticalAlignment="Top" Width="340"/>
                        <DataGrid x:Name="dataGrid" Margin="0,97,0,270" ItemsSource="{Binding subItems}" AutoGenerateColumns="False" AlternatingRowBackground="#b2ebf2" SelectionMode="Single" Grid.ColumnSpan="2">
                            <DataGrid.Columns>
                                <DataGridTextColumn Binding="{Binding Inventory-Item}" FontSize="14" Width="60" CanUserSort="False" CanUserResize="False" Header="Item" IsReadOnly="True"/>
                                <DataGridTextColumn Binding="{Binding Item-Description}" FontSize="14" Width="1350" CanUserSort="False"  CanUserResize="False" Header="Item Description" IsReadOnly="True"/>
                            </DataGrid.Columns>
                        </DataGrid>
                        <Button x:Name="Add_Selected" Content="Select" HorizontalAlignment="Left" Margin="10,384,0,180" FontFamily="georgia" FontSize="18" Width="150" Height="50"/>
                        <Button x:Name="CLR_Selected" Content="Clear Selected" HorizontalAlignment="Left" Margin="170,384,0,180" FontFamily="georgia" FontSize="18" Width="150" Height="50"/>
                        <Button x:Name="Add_to_Order" Content="Add to Order" HorizontalAlignment="Left" Margin="299,384,0,180" FontFamily="georgia" FontSize="18" Width="150" Height="50" Grid.Column="1"/>
                        <TextBox x:Name="QTY_Select" Background="#e8ebef" Margin="20,460,0,89" FontSize="27" FontWeight="Black" Width="980" HorizontalAlignment="Left" Height="70"  IsReadOnly="True" HorizontalScrollBarVisibility="Visible" Grid.ColumnSpan="2"/>
                        <TextBox x:Name="UOM_Select" Background="#e8ebef" Margin="235,460,0,89" FontSize="27" FontWeight="Black" HorizontalAlignment="Left" Height="70" Width="90" IsReadOnly="True" Grid.Column="1" />
                        <TextBox x:Name="QTY" MaxLength="4"  HorizontalAlignment="Left"  Margin="349,460,0,89" Width="100" Height="70" FontSize="30" FontWeight="Black" Grid.Column="1"/>
                        <Label Content="Selected Item Description" Margin="80,510,0,0" HorizontalAlignment="Left" FontWeight="Black" Height="50" FontSize="26" FontFamily="gerogia" />
                        <Label Content="UOM" Margin="245,540,0,29" HorizontalAlignment="Left" FontWeight="Black" Height="50" FontSize="26" FontFamily="gerogia" RenderTransformOrigin="4.908,0.57" Width="71" Grid.Column="1" />
                        <Label Content="QTY" Margin="365,540,0,29" HorizontalAlignment="Left" FontWeight="Black" Height="50" FontSize="26" FontFamily="gerogia" RenderTransformOrigin="7.556,0.45" Grid.Column="1" />
                    </Grid>
                </TabItem>
                <TabItem>
                    <TabItem.Header>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Preview Order" Width="410" Padding="130,2,0,0" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="20" FontWeight="Black" FontFamily="georgia"/>
                        </StackPanel>
                    </TabItem.Header>
                    <Grid>
                        <TextBox Text="Stock ID #" BorderBrush="Black" BorderThickness="2px" Background="#8a8c91" FontSize="20" Margin="10,0,0,500" HorizontalAlignment="Left" Height="30" Width="120" IsReadOnly="True"/>
                        <TextBox Text="Description" BorderBrush="Black" BorderThickness="2px" Background="#8a8c91" FontSize="20" Margin="130,0,0,500" HorizontalAlignment="Left" Height="30" Width="900" IsReadOnly="True"/>
                        <TextBox Text="UOM" BorderBrush="Black" BorderThickness="2px" Background="#8a8c91" FontSize="20" Margin="1030,0,0,500" HorizontalAlignment="Left" Height="30" Width="110" IsReadOnly="True"/>
                        <TextBox Text="Requested" BorderBrush="Black" BorderThickness="2px" Background="#8a8c91" FontSize="20" Margin="1140,0,0,500" HorizontalAlignment="Left" Height="30" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store1" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px" Margin="10,70,0,500" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des1" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px" Margin="130,70,0,500" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM1" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px" Margin="1030,70,0,500" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req1" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px" Margin="1140,70,0,500" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store2" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="10,110,0,450" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des2" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="130,110,0,450" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM2" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1030,110,0,450" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req2" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1140,110,0,450" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store3" Background="#b3d1ef" BorderBrush="Black" BorderThickness="2px" FontSize="20" Margin="10,150,0,400" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des3" Background="#b3d1ef" BorderBrush="Black" BorderThickness="2px" FontSize="20" Margin="130,150,0,400" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM3" Background="#b3d1ef" BorderBrush="Black" BorderThickness="2px" FontSize="20" Margin="1030,150,0,400" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req3" Background="#b3d1ef" BorderBrush="Black" BorderThickness="2px" FontSize="20" Margin="1140,150,0,400" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store4" Background="#b3d1ef" BorderBrush="Black" BorderThickness="2px" FontSize="20" Margin="10,200,230,360" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des4" Background="#b3d1ef" BorderBrush="Black" BorderThickness="2px" FontSize="20" Margin="130,200,0,360" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM4" Background="#b3d1ef" BorderBrush="Black" BorderThickness="2px" FontSize="20" Margin="1030,200,0,360" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req4" Background="#b3d1ef" BorderBrush="Black" BorderThickness="2px" FontSize="20" Margin="1140,200,0,360" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store5" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="10,230,0,300" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des5" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="130,230,0,300" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM5" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1030,230,0,300" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req5" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1140,230,0,300" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store6" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="10,280,0,260" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des6" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="130,280,0,260" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM6" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1030,280,0,260" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req6" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1140,280,0,260" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store7" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="10,310,0,200" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des7" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="130,310,0,200" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM7" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1030,310,0,200" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req7" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1140,310,0,200" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store8" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="10,350,0,150" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des8" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="130,350,0,150" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM8" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1030,350,0,150" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req8" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1140,350,0,150" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store9" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="10,390,0,100" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des9" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="130,390,0,100" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM9" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1030,390,0,100" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req9" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1140,390,0,100" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Store10" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="10,430,0,50" HorizontalAlignment="Left" Height="45" Width="120" IsReadOnly="True"/>
                        <TextBox x:Name="Des10" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="130,430,0,50" HorizontalAlignment="Left" Height="45" Width="900" IsReadOnly="True"/>
                        <TextBox x:Name="UOM10" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1030,430,0,50" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <TextBox x:Name="Req10" Background="#b3d1ef" FontSize="20" BorderBrush="Black" BorderThickness="2px"  Margin="1140,430,0,50" HorizontalAlignment="Left" Height="45" Width="110" IsReadOnly="True"/>
                        <Button x:Name="CLR_Order" FontSize="30" Content="Clear Order" HorizontalAlignment="Center" FontFamily="georgia" FontWeight="Black"  Height="70" Width="300" Margin="0,520,0,0" />
                    </Grid>
                </TabItem>
                <TabItem>
                    <TabItem.Header>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Send Order" Width="410" Height="40" Padding="150,12,0,0" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="18" FontWeight="Black" FontFamily="georgia"/>
                        </StackPanel>
                    </TabItem.Header>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="203*"/>
                            <ColumnDefinition Width="1066*"/>
                        </Grid.ColumnDefinitions>
                        <Label Content="Enter your Project Name:" FontSize="34" FontFamily="georgia" FontWeight="Black" Margin="20,20,0,0" Grid.ColumnSpan="2"/>
                        <Label Content="Enter Your User Name:" FontSize="34" FontFamily="georgia" FontWeight="Black" Margin="70,150,0,0" Grid.ColumnSpan="2"/>
                        <Label Content="Enter Your Password:" FontSize="34" FontFamily="georgia" FontWeight="Black" Margin="557.5,150,0,0" Grid.Column="1"/>
                        <Label Content="WR#:" FontSize="34" FontFamily="georgia" FontWeight="Black" Margin="382.5,20,0,0" Grid.Column="1"/>
                        <Label Content="Select Service Territory:" FontSize="34" FontFamily="georgia" FontWeight="Black" Margin="592.5,20,0,0" Grid.Column="1"/>
                        <Label Content="Additional Comments:" FontSize="34" FontFamily="georgia" FontWeight="Black" Margin="237.5,280,0,0" Grid.Column="1"/>
                        <ComboBox x:Name="Territory" HorizontalAlignment="Left" FontSize="26" Margin="607.5,70,0,0" Height="40" Width="400" VerticalAlignment="Top" Grid.Column="1">
                            <ComboBoxItem Content="Reading" FontSize="26" FontFamily="georgia"/>
                            <ComboBoxItem Content="Bloomsburg" FontSize="26" FontFamily="georgia"/>
                        </ComboBox>
                        <TextBox x:Name="Enter_Name" Width="450" Height="40" Margin="55,200,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" FontSize="24" FontFamily="georgia" Grid.ColumnSpan="2" />
                        <PasswordBox x:Name="Enter_PW" Width="450" Height="40" Margin="532.5,200,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" FontSize="24" FontFamily="georgia" Grid.Column="1" />
                        <TextBox x:Name="Enter_PRJ" Width="450" Height="40" Margin="15,70,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" FontSize="24" FontFamily="georgia" Grid.ColumnSpan="2" />
                        <TextBox x:Name="Enter_WR" MaxLength="6" Width="130" Height="40" Margin="367.5,70,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" FontSize="24" FontFamily="georgia" Grid.Column="1" />
                        <TextBox x:Name="Comments" Width="900" Height="140" Margin="184,330,184,0" HorizontalAlignment="Center" VerticalAlignment="Top" FontSize="30" FontFamily="georgia" TextWrapping="Wrap" Grid.ColumnSpan="2"/>
                        <Button x:Name="Email" Width="370" Height="80" Content="Complete Order" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="246.5,500,449,0" FontSize="38" FontFamily="georgia" FontWeight="Black" Grid.Column="1"/>
                    </Grid>
                </TabItem>
            </TabControl>
        </Grid>
    </Grid>
</Window>




"@
#######################################################################################################
#######################################################################################################
##############################  Make Powershell Read XAML   ###########################################
####################################################################################################### 
#====================================================================================================#
#                                   End Wpf GUI/GUI Reader                                           #
#====================================================================================================#
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
#===================================================================================================#
#                            Store Form Objects In PowerShell                                       #
#===================================================================================================#
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF_$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables
#######################################################################################################
#######################################################################################################
##############################   Make Buttons Work    #################################################
#######################################################################################################
#=====================================================================================================#
#                                    Populate Combo                                                   #
#=====================================================================================================#
foreach($x in $smelly) {
	$WPF_comboBox.AddText($x.name)
}
#======================================================================================================#
#                                       dataGrid                                                       #
#======================================================================================================#
$WPF_comboBox.Add_SelectionChanged({
	# DataGrid needs ArrayList
	$item_Array = New-Object System.Collections.ArrayList
	#Write-Host $WPF_comboBox.SelectedIndex
	switch ($WPF_comboBox.SelectedIndex)
	{
		0 {$item_Array.AddRange($ssm)}
		1 {$item_Array.AddRange($ssmm)}
		2 {$item_Array.AddRange($pm)}
		3 {$item_Array.AddRange($ps)}
	}
	$WPF_dataGrid.Clear()
	$WPF_dataGrid.ItemsSource = $item_Array
	#$WPF_dataGrid.IsReadOnly = $true
})

#=====================================================================================================#
#                                   Button  Add_Selected                                              #
#=====================================================================================================#
$WPF_Add_Selected.Add_Click({
    $Global:Selected= $WPF_dataGrid.SelectedItem.'Item-Description'
    $Global:Select_UOM= $WPF_dataGrid.SelectedItem.'Primary-Uom'
    $Global:Select_ID= $WPF_dataGrid.SelectedItem.'Inventory-Item'
    $WPF_QTY_Select.Clear()
    $WPF_QTY_Select.Text= $Global:selected
    $WPF_UOM_Select.Text= $Global:Select_UOM
})

$WPF_CLR_Selected.Add_Click({
    $WPF_QTY_Select.Clear()
    $WPF_QTY_Select.Text=""
    $WPF_UOM_Select.Clear()
    $WPF_UOM_Select.Text=""
    $WPF_QTY.Clear()
    $Global:Selected= $null
    $Global:Select_UOM= $null
    $Global:Select_ID= $null
})

$WPF_Add_to_Order.Add_Click({
    $Global:Amount= $WPF_QTY.Text
    CheckNull
    Write-Host $Description.Count
    
})

$WPF_CLR_Order.Add_Click({
    $Global:ID=@() 
    $Global:Description=@()
    $Global:UOM=@()
    $Global:Req=@()
    BuildTable
})

$WPF_Email.Add_Click({
    FillOutForm
})

#=====================================================================================================#
#                                    Variables Pre-PDF                                                #
#=====================================================================================================#

#=====================================================================================================#
#                                          Build PDF                                                  #
#=====================================================================================================#

#######################################################################################################
#######################################################################################################
#######################################################################################################
#######################################################################################################
#######################################################################################################
#######################################################################################################
#=====================================================================================================#
#                                          Shows the form                                             #
#=====================================================================================================#
#write-host "To show the form, run the following" -ForegroundColor Cyan
$Form.ShowDialog() | out-null
#=====================================================================================================#       
