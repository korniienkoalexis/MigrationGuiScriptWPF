#region XAML code
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Migration O365" Height="420" Width="550" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid>
        <DockPanel>
            <Menu DockPanel.Dock="Top">
                <MenuItem Header="_File">
                    <MenuItem x:Name="Btn_Exit" Header="_Exit" />
                </MenuItem>

                <MenuItem Header="_Edit">
                    <MenuItem Command="Cut" />
                    <MenuItem Command="Copy" />
                    <MenuItem Command="Paste" />
                </MenuItem>
            </Menu>
        </DockPanel>
        <TabControl Margin="0,20,0,0">
            <TabItem x:Name="Tab_MailBoxes" Header="Migration" TabIndex="12">
                <Grid Background="White" Height="272" Margin="0,0,0.4,0" VerticalAlignment="Top">
                    <StackPanel>
                        <StackPanel Height="32" HorizontalAlignment="Center" VerticalAlignment="Top" Width="538" Margin="0,0,0,0">
                            <Label Content="Migration mailboxes in Office 365" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0" Height="32" FontWeight="Bold"/>
                        </StackPanel>
                        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Top" Width="538" Margin="0,10,0,0">
                            <GroupBox Header="On-prem MailBoxes:" Width="508" Margin="10,0,0,0" FontSize="11" HorizontalAlignment="Left" VerticalAlignment="Top" Height="164">
                                <Grid Margin="0,10,-2,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="491*"/>
                                        <ColumnDefinition Width="7*"/>
                                    </Grid.ColumnDefinitions>


                                    <ListView Name="listView" Margin="-1,0,0.4,-7.4" HorizontalAlignment="Right" Width="491">
                                        <ListView.View>
                                            <GridView>
                                                <GridViewColumn Header="Display Name" DisplayMemberBinding ="{Binding DisplayName}"/>
                                                <GridViewColumn Header="Status Detail" DisplayMemberBinding ="{Binding StatusDetail}"/>
                                                <GridViewColumn Header="Total Mailbox Size" DisplayMemberBinding ="{Binding TotalMailboxSize}"/>
                                                <GridViewColumn Header="Total Archive Size" DisplayMemberBinding ="{Binding TotalArchiveSize}"/>
                                                <GridViewColumn Header="Percent Complete" DisplayMemberBinding ="{Binding PercentComplete}"/>
                                            </GridView>
                                        </ListView.View>
                                        <ListView.ContextMenu>
                                            <ContextMenu>
                                                <MenuItem Header="Menu item 1" />
                                                <MenuItem Header="Move Request Statistic/Refresh" Name="Btn_MoveRequestStatistic" />
                                                <Separator />
                                                <MenuItem Header="Remove Move Request" Name="Btn_RemoveMoveRequest" />
                                            </ContextMenu>
                                        </ListView.ContextMenu>
                                    </ListView>


                                </Grid>
                            </GroupBox>
                            <GroupBox Header="Options:" Width="508" Margin="10,10,0,0" FontSize="11" HorizontalAlignment="Left" VerticalAlignment="Top" Height="50">
                                <Grid Height="50" Margin="0,10,0,0">
                                    <CheckBox x:Name="Box_SuspendWhenReadyToComplete" TabIndex="10" HorizontalAlignment="Left" VerticalAlignment="Top" Content="SuspendWhenReadyToComplete"/>
                                    <CheckBox x:Name="Box_WhatIf" TabIndex="11" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="20,0,0,0" Content="WhatIf"/>
                                    <StackPanel HorizontalAlignment="Left" VerticalAlignment="Bottom" Orientation="Horizontal"/>
                                </Grid>
                            </GroupBox>
                        </StackPanel>

                    </StackPanel>
                    <Button Name="Btn_Create" Content="New Move Request" TabIndex="13" Margin="182,277,184,-30" RenderTransformOrigin="0.789,0.444" />
                    <Button Name="Btn_Ok" Content="Ok" TabIndex="13" Margin="182,312,282,-65" RenderTransformOrigin="0.872,0.508" />
                    <Button Name="Btn_Cancel" Content="Cancel" Margin="279,312,184,-65" ClickMode="Press" RenderTransformOrigin="0.8,6.793" />

                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_Option" Header="Options" TabIndex="11">
                <Grid Background="White">
                    <StackPanel>
                        <StackPanel>
                            <Grid Margin="0,10,0,0">
                                <Label Content="Parameter specifies the FQDN of on-prem Exchange and Exchange Online" HorizontalAlignment="Left" FontSize="11" FontWeight="Bold"/>
                                <Button Name="Btn_Connect" Content="Connect" HorizontalAlignment="Left" Margin="428,2,0,0" VerticalAlignment="Top" Width="100"/>
                            </Grid>
                            <StackPanel>
                                <Label BorderBrush="Black" BorderThickness="0,0,0,1" VerticalAlignment="Top"/>
                            </StackPanel>
                        </StackPanel>
                        <StackPanel>
                            <Grid Margin="0,10,0,0">
                                <TextBlock Name="Txt_RemoteHostName" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                            </Grid>
                        </StackPanel>
                        <StackPanel>

                        </StackPanel>
                        <StackPanel/>
                        <StackPanel/>
                        <StackPanel/>
                    </StackPanel>
                    <GroupBox x:Name="Exchange_on_prem" Header="Exhange On-prem" Height="130" Margin="10,0,0,166" VerticalAlignment="Bottom" HorizontalAlignment="Left" Width="518"/>
                    <Grid Margin="0,10,0,0">
                        <Label Width="70" VerticalContentAlignment="Center" VerticalAlignment="Center" Margin="0,64,448,237" Height="32" HorizontalAlignment="Right" FontSize="11" Content="On-premiss Server:" RenderTransformOrigin="0.504,-1.812"/>
                        <TextBox Name="RemoteHostName" HorizontalAlignment="Right" Margin="0,69,23,242" TextWrapping="Wrap" Width="408" VerticalContentAlignment="Center" FontSize="11" BorderThickness="1"/>
                        <Label Content="Username:" HorizontalAlignment="Left" Height="32" Margin="20,95,0,206" VerticalAlignment="Center" Width="70" FontSize="11" VerticalContentAlignment="Center" RenderTransformOrigin="0.726,0.167"/>
                        <TextBox Name="Field_User_OnPrem" HorizontalAlignment="Left" Height="22" Margin="107,101,0,210" TextWrapping="Wrap" VerticalAlignment="Center" Width="408" VerticalContentAlignment="Center" FontSize="11" BorderThickness="1" />
                        <Label Content="Password:" HorizontalAlignment="Right" Height="32" Margin="0,127,448,174" VerticalAlignment="Center" Width="70" FontSize="11" VerticalContentAlignment="Center"/>
                        <PasswordBox Name="Field_Pwd_OnPrem" HorizontalAlignment="Left" Height="22" Margin="107,132,0,179" VerticalAlignment="Center" Width="408" VerticalContentAlignment="Center" FontSize="11" BorderThickness="1" />
                    </Grid>
                    <GroupBox Header="Exchange Online" HorizontalAlignment="Left" Height="141" Margin="10,182,0,0" VerticalAlignment="Top" Width="508"/>
                    <Grid Margin="0,10,0,0">
                        <Label Width="70" VerticalContentAlignment="Center" VerticalAlignment="Center" Margin="20,203,0,98" Height="32" HorizontalAlignment="Left" FontSize="11" Content="Cloud Tenant:"/>
                        <TextBox Name="TargetDeliveryDomain" HorizontalAlignment="Right" Height="22" Margin="0,208,28,103" TextWrapping="Wrap" VerticalAlignment="Center" Width="402" VerticalContentAlignment="Center" FontSize="11" BorderThickness="1"/>
                        <Label Content="Username:" HorizontalAlignment="Left" Height="32" Margin="20,233,0,68" VerticalAlignment="Center" Width="70" FontSize="11" VerticalContentAlignment="Center" RenderTransformOrigin="0.726,0.167"/>
                        <TextBox Name="Field_User_OnLine" HorizontalAlignment="Left" Height="22" Margin="108,240,0,71" TextWrapping="Wrap" VerticalAlignment="Center" Width="402" VerticalContentAlignment="Center" FontSize="11" BorderThickness="1" />
                        <Label Content="Password:" HorizontalAlignment="Left" Height="32" Margin="20,265,0,36" VerticalAlignment="Center" Width="70" FontSize="11" VerticalContentAlignment="Center"/>
                        <PasswordBox Name="Field_Pwd_OnLine" HorizontalAlignment="Left" Height="22" Margin="108,270,0,41" VerticalAlignment="Center" Width="402" VerticalContentAlignment="Center" FontSize="11" BorderThickness="1" />
                    </Grid>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
'@

#endregion
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAMLGui = $XAML

$Reader=(New-Object System.Xml.XmlNodeReader $XAMLGui)
$MainWindow=[Windows.Markup.XamlReader]::Load( $Reader )
$XAMLGui.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name "GUI$($_.Name)" -Value $MainWindow.FindName($_.Name)}


function Get-CancelBtn{
    $MainWindow.Close()
    $Script:EndScript = 1
	Close-Window 'Script cancelled'
}

Function Close-Window ($CloseReason) {
    Write-Host "$CloseReason" -ForegroundColor Red
    Exit
}

Function Get-Mailboxes{
    $mailboxes = Get-Recipient -Filter {RecipientType -eq "MailUser"} | select Name, alias, PrimarySmtpAddress 
    $mailboxes | foreach {
    $GUIlistView.Items.Add([pscustomobject]@{VMName=($_.Name);Status=($_.Alias);Email=($_.PrimarySmtpAddress )})
    }
}


Function Connect-EXO{   
      $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CloudCredential -Authentication Basic -AllowRedirection
      Import-PSSession $EXOSession  
}


Function Get-Cred {

    $userNameOnPrem = $GUIField_User_OnPrem.Text
    $PwdOnPrem = $GUIField_Pwd_OnPrem.Password

    $userNameCloud = $GUIField_User_OnLine.Text
    $PwdCloud = $GUIField_Pwd_OnLine.Password

    If (!$userNameOnPrem -or !$PwdOnPrem -or !$userNameCloud -or !$PwdCloud) {
    [System.Windows.MessageBox]::Show("Please enter valid credentials..","Empty credentials",([System.Windows.MessageBoxButton]::OK),([System.Windows.MessageBoxImage]::Warning))
    } else {
    $EncryptPwdOnPrem = $PwdOnPrem | ConvertTo-SecureString -AsPlainText -Force
    $OnPremCredential = New-Object System.Management.Automation.PSCredential($userNameOnPrem,$EncryptPwdOnPrem)

    $EncryptPwdCloud = $PwdCloud | ConvertTo-SecureString -AsPlainText -Force
    $CloudCredential = New-Object System.Management.Automation.PSCredential($userNameCloud,$EncryptPwdCloud)
    } 
    return $CloudCredential
    return $OnPremCredential
}

Function Get-hosts {
    $RemoteHostName =$GUIRemoteHostName.Text
    $TargetDeliveryDomain = $GUITargetDeliveryDomain.Text
    $RemoteHostName
    $TargetDeliveryDomain
}


# Event Handlers

#Field_User_OnPrem




#$mailboxes = get-mailbox

#$mailboxes | foreach {$_.name}

$GUIBtn_Connect.add_Click({

    
    
    $userNameOnPrem = $GUIField_User_OnPrem.Text
    $PwdOnPrem = $GUIField_Pwd_OnPrem.Password

    $userNameCloud = $GUIField_User_OnLine.Text
    $PwdCloud = $GUIField_Pwd_OnLine.Password

    If (!$userNameOnPrem -or !$PwdOnPrem -or !$userNameCloud -or !$PwdCloud) {
    [System.Windows.MessageBox]::Show("Please enter valid credentials..","Empty credentials",([System.Windows.MessageBoxButton]::OK),([System.Windows.MessageBoxImage]::Warning))
    } else {
    $EncryptPwdOnPrem = $PwdOnPrem | ConvertTo-SecureString -AsPlainText -Force
    $OnPremCredential = New-Object System.Management.Automation.PSCredential($userNameOnPrem,$EncryptPwdOnPrem)

    $EncryptPwdCloud = $PwdCloud | ConvertTo-SecureString -AsPlainText -Force
    $CloudCredential = New-Object System.Management.Automation.PSCredential($userNameCloud,$EncryptPwdCloud)
    }

    $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CloudCredential -Authentication Basic -AllowRedirection
    Import-PSSession $EXOSession
})


$GUIBtn_Create.add_Click({
    
    $RemoteHostName =$GUIRemoteHostName.Text
    $TargetDeliveryDomain = $GUITargetDeliveryDomain.Text

    $userNameOnPrem = $GUIField_User_OnPrem.Text
    $PwdOnPrem = $GUIField_Pwd_OnPrem.Password

    $userNameCloud = $GUIField_User_OnLine.Text
    $PwdCloud = $GUIField_Pwd_OnLine.Password

    If (!$userNameOnPrem -or !$PwdOnPrem -or !$userNameCloud -or !$PwdCloud) {
    [System.Windows.MessageBox]::Show("Please enter valid credentials..","Empty credentials",([System.Windows.MessageBoxButton]::OK),([System.Windows.MessageBoxImage]::Warning))
    } else {
    $EncryptPwdOnPrem = $PwdOnPrem | ConvertTo-SecureString -AsPlainText -Force
    $OnPremCredential = New-Object System.Management.Automation.PSCredential($userNameOnPrem,$EncryptPwdOnPrem)

    $EncryptPwdCloud = $PwdCloud | ConvertTo-SecureString -AsPlainText -Force
    $CloudCredential = New-Object System.Management.Automation.PSCredential($userNameCloud,$EncryptPwdCloud)
            }

    $MigrationMailbox = @()
    $MigrationMailbox = Get-Recipient -ResultSize unlimited| where {$_.RecipientType -eq "MailUser" } | Select DisplayName,UserPrincipalname,ExchangeGuid,PrimarySmtpAddress,ArchiveStatus | Out-GridView -Title "Select Mailbox for migration to Exchnage Online" -PassThru
    $MigrationMailbox | foreach {
                                    New-MoveRequest $_.PrimarySmtpAddress -Remote -RemoteHostName $RemoteHostName -TargetDeliveryDomain $TargetDeliveryDomain -RemoteCredential $OnPremCredential -LargeItemLimit 1000 -BadItemLimit 1000 -AcceptLargeDataLoss:$true
                                    Write-Host "Migration request successfully created for user "$_.DisplayName" ("$_.PrimarySmtpAddress")"
                                }
})


$GUIBtn_MoveRequestStatistic.add_Click({

$GUIlistView.Items.Clear()
$MoveRequestStat = Get-MoveRequest | Get-MoveRequestStatistics
$MoveRequestStat | % {$GUIlistView.Items.Add([pscustomobject]@{VMName=("User NAme");Status=("user1");Email=("user1@domain.com")})}
})

$GUIBtn_Cancel.add_Click({
    #$MainWindow.Close()

    #$GUIlistView.Items.Add([pscustomobject]@{VMName=("User NAme");Status=("user1");Email=("user1@domain.com")})
    #$GUIlistView.Items.Add([pscustomobject]@{VMName=("User2 NAme2");Status=("user2");Email=("user2@domain.com")})
    #Get-hosts | 
    
    #write-host $TargetDeliveryDomain
    Get-Cred
    write-host $OnPremCredential


    #$EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CloudCredential -Authentication Basic -AllowRedirection

#      Import-PSSession $EXOSession


})


$GUIBtn_Ok.Add_Click({
                        #While ($vmpicklistView.SelectedItems.count -gt 0) {
                        #$vmpicklistView.Items.RemoveAt($vmpicklistView.SelectedIndex)
                        #$selected = $vmpicklistView.SelectedItems
                        
                        $a = @()
                        #$GUIlistView.SelectedItems | foreach {$a=$_.Email; $a}
                         #$GUIlistView.Items($vmpicklistView.SelectedIndex).ToString()
                        $a = $GUIlistView.SelectedItems.Email
                        Write-Host $a
                                #}
                      })


#Get-Mailboxes


# Load GUI Window
$MainWindow.WindowStartupLocation = "CenterScreen"
$MainWindow.ShowDialog() | Out-Null

# Check if Window is closed




