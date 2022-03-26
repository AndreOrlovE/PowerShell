param(
    [Parameter(Position=0,mandatory=$true)]
    [string]$cmd,
    [Parameter(Position=1,mandatory=$false)]
    [string]$server,
    [Parameter(Position=2,mandatory=$false)]
    [string]$name,
    [Parameter(Position=3,mandatory=$false)]
    $config = "$PSScriptRoot\config.xml",
    [switch]$detailed = $false 
)

#v 3.9.1
#$config = "c:\Temp\config.xml"
$log = "$(Split-Path $config -Parent)\gameOz.log"

###########################################
# ICONS
###########################################
$ico_refresh_path = "$PSScriptRoot\ico\icon_refresh.png"
$ico_wu_path = "$PSScriptRoot\ico\icon_wu.png"

[xml]$xmlConfig = Get-Content -Path ($config)

$env_color = $xmlConfig.config.system.env.color
#$version = $xmlConfig.config.system.version
$version = "v3.9.2"
$env_name =  $xmlConfig.config.system.env.'#text'
$group = $xmlConfig.config.system.env.group
$warning = $xmlConfig.config.system.env.warning
$win_text = "GameOz $version - $env_name"

###########################################
# EMAIL
###########################################
$email_server = $xmlConfig.config.system.email.server
$email_to = $xmlConfig.config.system.email.'#text'
$email_from = $xmlConfig.config.system.email.from
$email_subj = $xmlConfig.config.system.email.subj
$email_enabled = $xmlConfig.config.system.email.enabled

###########################################
# WINDOWS UPDATE
###########################################
$wu_enabled = $xmlConfig.config.system.wu.enabled

###########################################
# XML
###########################################
function read_services($xmlConfig, $server, $name)
{
    $serviceArr = @()
    
    # Build Array of Servers
    $index = 0

    # filter
    if($server) { $xmlConfig = $xmlConfig.config.services.server| Where-Object {$_.name -eq $server } }
    else { $xmlConfig = $xmlConfig.config.services.server }
    if($name) { $xmlConfig = $xmlConfig | Where-Object {$_.service.name -eq $name }}
    
    foreach($xml in $xmlConfig) 
    {
        $serviceArr += @{
            Desc = $xml.service.'#text'
            Name = $xml.service.name
            Server = $xml.name
            Attr1 = $xml.service.start_wait
            Attr2 = $xml.service.stop_wait
            Index = $index
        }
        $index ++                
    }
    return $serviceArr
}
function read_tasks($xmlConfig, $server, $name)
{
    $taskArr = @()
    
    # Build Array of Servers
    $index = 0

    # filter
    if($server) { $xmlConfig = $xmlConfig.config.tasks.server| Where-Object {$_.name -eq $server } }
    else { $xmlConfig = $xmlConfig.config.tasks.server }
    if($name) { $xmlConfig = $xmlConfig | Where-Object {$_.task.name -eq $name }}
    
    foreach($xml in $xmlConfig) 
    {
        $taskArr += @{
            Desc = $xml.task.'#text'
            Name = $xml.task.name
            Server = $xml.name
            Attr1 = $xml.rdpuser
            Index = $index
        }
        $index ++                
    }
    return $taskArr
}
function read_firewall($xmlConfig, $server, $name)
{
    $firewallArr = @()
    
    # Build Array of Servers
    $index = 0

    # filter
    if($server) { $xmlConfig = $xmlConfig.config.firewall.server| Where-Object {$_.name -eq $server } }
    else { $xmlConfig = $xmlConfig.config.firewall.server }
    if($name) { $xmlConfig = $xmlConfig | Where-Object {$_.rule.name -eq $name }}
    
    foreach($xml in $xmlConfig) 
    {
        $firewallArr += @{
            Desc = $xml.rule.'#text'
            Name = $xml.rule.name
            Server = $xml.name
            Index = $index
        }
        $index ++                
    }
    return $firewallArr
}
function read_processes($xmlConfig, $server, $name)
{
    $processArr = @()
    
    # Build Array of Servers
    $index = 0

    # filter
    if($server) { $xmlConfig = $xmlConfig.config.processes.server| Where-Object {$_.name -eq $server } }
    else { $xmlConfig = $xmlConfig.config.processes.server }
    if($name) { $xmlConfig = $xmlConfig | Where-Object {$_.process.name -eq $name }}
    
    foreach($xml in $xmlConfig) 
    {
        $processArr += @{
            Desc = $xml.process.'#text'
            Name = $xml.process.name
            Server = $xml.name
            Index = $index
        }
        $index ++                
    }
    return $processArr
}
function read_servers($xmlConfig, $server)
{
    $serverArr = @()
    
    # Build Array of Servers
    $index = 0

    # filter
    if($server) { $xmlConfig = $xmlConfig.config.servers.server| Where-Object {$_.name -eq $server } }
    else { $xmlConfig = $xmlConfig.config.servers.server }
    
    foreach($xml in $xmlConfig) 
    {
        $serverArr += @{
            #Desc = $xml.'#text'
            Name = $xml.'#text'
            Server = $xml.name
            Index = $index
        }
        $index ++                
    }
    return $serverArr
}
function read_folders($xmlConfig, $name)
{
    $folderArr = @()

    $index = 0    
    

    # Building Array of Folders

    if($name) { $xmlConfig = $xmlConfig.config.folders.group | Where-Object {$_.name -eq $name } }
    else { $xmlConfig = $xmlConfig.config.folders.group }

    foreach($xml in $xmlConfig) 
    {
        $index2 = 0
        foreach($xml2 in $xml.folder)
        {
            $folderArr += @{
                Desc = $xml2.'#text'
                Server = $xml2.server
                Path = $xml2.path
                Name = $xml.name
                # and now we have a problem (regexp)
                Exc1 = $xml.folders_exclude -replace ",\s+", ","
                Exc2 = $xml.files_exclude -replace ",\s+", ","
                Id = $index2
                Index = $index
            }
            $index2++
        }
        $index++                
    }
    return $folderArr
}
function read_buttons($xmlConfig)
{
    $buttonArr = @()
    # Build Array of Buttons
    $index = 0
    $xmlConfig = $xmlConfig.config.buttons.button
    foreach($xml in $xmlConfig) 
    {
        $buttonArr += @{
            Name = $xml.'#text'
            Cmd = $xml.cmd
            Confirm = $xml.confirm
            Index = $index
        }
        $index ++                
    }
    return $buttonArr
}
###########################################

###########################################
# WORKFLOW
###########################################
workflow workflow_common
{
    Param (
        $data,
        $cmd
    )

    foreach -parallel ($record in $data)
    {
        try
        {
            InlineScript
            {
                $server = $Using:record.Server
                $name = $Using:record.Name
                $index = $Using:record.Index
                $desc = $Using:record.Desc
                $attr1 = $Using:record.Attr1
                $attr2 = $Using:record.Attr2
                $cmd = $Using:cmd

                $status = $null

                if($cmd -eq "gettasks")
                {
                    [string]$status = (Get-ScheduledTask -TaskName $name -ErrorAction Stop).State

                    # check RDP
                    $attr2 = $false
                    if($attr1)
                    {                        
                        $rdps = (qwinsta /server:localhost) | foreach { (($_  -replace '\s{2,}', ','))} | convertfrom-csv
                        if($rdps)
                        {
                            
                            ForEach ($rdp in $rdps) 
                            {
                                if($rdp.USERNAME -eq $attr1 -and $rdp.STATE -eq "Active") { $attr2 = $true }
                            }
                        }
                    }
                    else
                    {
                        $attr2 = $true
                    }
                }
                elseif($cmd -eq "getps")
                { 
                    [string]$status = $PSVersionTable.PSVersion
                }
                elseif($cmd -eq "getprocesses")
                {
                    [string]$status = ((Get-Process -Name $name -ErrorAction Stop).Id)
                }
                elseif($cmd -eq "getservices")
                {
                    [string]$status = (Get-Service -Name $name  -ErrorAction Stop).Status
                }               
                elseif($cmd -eq "getservers")
                {
                    $status = (Get-WmiObject -Query "SELECT LastBootUpTime FROM Win32_OperatingSystem").LastBootUpTime
                    $status = ( (get-date)-([System.Management.ManagementDateTimeconverter]::ToDateTime($status)) )
                    $status = [string]$status.Days + "d" + [string]$status.Hours + "h" + [string]$status.Minutes + "m"
                }
                elseif($cmd -eq "getfw")
                {
                    [string]$status = (Get-NetFirewallRule -DisplayName $name).Enabled
                    switch ($status)
                    {
                        $True { $status = "Opened" } 
                        $False { $status = "Closed" }
                    }
                }
                elseif($cmd -eq "getps")
                { 
                    [string]$status = $PSVersionTable.PSVersion
                }
                elseif($cmd -eq "startservices")
                { 
                    Start-Service -Name $name # Always works with WAIT Flag
                }
                elseif($cmd -eq "stopservices")
                { 
                    if($attr2 -eq "true")
                    {
                        Stop-Service -Name $name
                    }
                    else
                    {
                        Stop-Service -Name $name -NoWait                        
                    }
                }
                elseif($cmd -eq "setfw")
                { 
                    switch ((Get-NetFirewallRule -DisplayName $name).Enabled)
                    {
                        "True" { $status = "False" } 
                        "False" { $status = "True" } 
                    }
                    Set-NetFirewallRule -DisplayName $name -Enabled $status
                }
                elseif($cmd -eq "starttasks")
                {
                    [string]$status = (Start-ScheduledTask -TaskName $name).State
                }
                elseif($cmd -eq "stoptasks")
                {
                    [string]$status = (Stop-ScheduledTask -TaskName $name).State
                }
                elseif($cmd -eq "reboot")
                {
                    Restart-Computer -computer localhost -force
                }
                elseif($cmd -eq "getwu")
                {
                    # "6.3" Windows Server 2012 R2
                    # "6.2" Windows Server 2012
                    $status = "n/a"
                    $desc = "n/a"
                    if(([environment]::OSVersion.Version).Major -eq 6)
                    {
                        if(Get-Module -ListAvailable -Name PSWindowsUpdate)
                        {
                            Import-Module PSWindowsUpdate                    
                            $status = "$((Get-WindowsUpdate).Count)"
                        }
                        
                        [string]$desc = (Get-ScheduledTask -TaskName "WUInstall" -ErrorAction SilentlyContinue).State
                    }
                }
                elseif($cmd -eq "startwu")
                {
                    [string]$status = (Start-ScheduledTask -TaskName "WUInstall" -ErrorAction SilentlyContinue).State
                }
                $return = @{
                    Server = $server
                    Name = $name
                    Index = $index
                    Desc = $desc
                    Status = $status
                    Attr1 = $attr1
                    Attr2 = $attr2
                }
                return $return

            } -PSComputerName $record.Server
        }
        catch
        {
            InlineScript
            {
                $status = "n/a"
            
                $server = $Using:record.Server
                $name = $Using:record.Name
                $index = $Using:record.Index
                $desc = $Using:record.Desc

                $return = @{
                    Server = $server
                    Name = $name
                    Index = $index
                    Desc = $desc
                    Status = $status
                }
                return $return
            }              
        }
    }
}
workflow workflow_folders
{
    Param (
        $data
    )

    foreach -parallel ($record in $data)
    {         
        
        try
        {
            InlineScript
            {
                $server = $Using:record.Server
                $name = $Using:record.Name
                $desc = $Using:record.Desc
                $path = $using:record.Path
                $index = $Using:record.Index
                $id = $Using:record.Id
                $folders_excluded = ($Using:record.Exc1).Split(",")
                $files_excluded = ($Using:record.Exc2).Split(",")

                $folder_size = 0
                $folder_files = 0

                # Symlink Folders
                $folders = Get-ChildItem  -Directory -Path $path -Recurse  -ErrorAction silentlycontinue  | Where { $_.LinkType -eq "SymbolicLink" } | Select-Object FullName
                $folders_excluded_symlinks = @()
                $folders_excluded_symlinks.Clear()
                foreach ($folder in $folders)
                {
                    $folders_excluded_symlinks += "$($folder.FullName)\*"
                }

                $files = Get-ChildItem -File -Path $path -Recurse -Exclude $files_excluded -ErrorAction silentlycontinue |  where { !$_.PSIsContainer } | where { $FullName = $_.FullName; $Length = $_.Length; -not @( $folders_excluded_symlinks | Where { $FullName -like $_ }) }
                foreach ($file in $files)
                {
                    $flag_exclude = 0
                    foreach ($folder_excluded in $folders_excluded)
                    {
                        $folder_excluded = $folder_excluded.Replace("\", "\\")
                        if($file.FullName -match "\\$folder_excluded\\")
                        {
                            $flag_exclude = 1
                        }
                    }
                    if ($flag_exclude -eq 0 )
                    {
                        $folder_size = $folder_size + $file.Length
                        $folder_files = $folder_files + 1
                    }
                }
                $folder_size = $folder_size / 1MB

                $return = @{
                    Server = $server
                    Name = $name
                    Index = $index
                    Desc = $desc
                    Status = $status
                    Id = $id
                    Path = $path
                    Size = $folder_size
                    Quantity = $folder_files
                }
                return $return

            } -PSComputerName $record.Server
        }
        catch
        {
            InlineScript
            {
                $server = $Using:record.Server
                $name = $Using:record.Name
                $desc = $Using:record.Desc
                $path = $using:record.Path
                $index = $Using:record.Index
                $id = $Using:record.Id
                $status = "n/a"


                $return = @{
                    Server = $server
                    Name = $name
                    Index = $index
                    Desc = $desc
                    Status = $status
                    Id = $id
                    Path = $folder_size
                    Size = $attr2
                    Quantity = $folder_files
                }
                return $return
            }
        }         
    }
}
###########################################

###########################################
# OUTPUT 
###########################################
function write_out($message, $status)
{
    $color = status_color $status

    $native_color = $host.ui.RawUI.ForegroundColor
    $host.ui.RawUI.ForegroundColor = $color
    write-output $message
    $host.ui.RawUI.ForegroundColor = $native_color
}
function status_color($status)
{
    switch($status)
    {
        "n/a" {$color = "DarkGray"}
        "Running" {$color = "Green"}
        "Stopped" {$color = "Red"}
        "StopPending" {$color = "Yellow"}
        "StartPending" {$color = "Yellow"}
        "Down" {$color = "Red"}
        "Closed" {$color = "Red"}
        "Opened" {$color = "Green"}
        "Error" {$color = "Red"}
        "Ok" {$color = "Green"}
        "Ready" {$color = "Yellow"}
        default {$color = "Green"}
    }
    return $color
}
function write_log ($message)
{
	try
	{
	    $cur_time =  Get-Date -format "yyyy-MM-dd HH:mm:ss"			
        $cur_time +": " + $message >> $log
	}
	catch
	{
		Write-Host $_.Exception.Message 0
	}
}
function send_email ($message, $caption)
{
    if($email_enabled -eq 1)
    {
        $message = "<table align='center' border=1 bordercolor='black' width='70%' style='border-collapse:collapse;'><tr align='center'><td colspan=4 >$caption</td></tr>`
        $message`
        </table>"

        send-mailmessage -from $email_from -to $email_to -subject $email_subj -smtpServer $email_server -BodyAsHtml $message # -Body $emailMess
        #write_log "send-mailmessage -from $email_from -to $email_to -subject $email_subj -smtpServer $email_server -BodyAsHtml $message"
        $message = "EMAIL: $email_to"
        write_log $message
    }
}
###########################################

###########################################
# FUNCTIONS WIN
###########################################
function win_mode
{
    Add-Type -assembly System.Windows.Forms

    if((read_tasks $xmlConfig $server $name))
    {
        $win_w = 1360
        $win_h = 670 
    }
    else
    {
        $win_w = 1140
        $win_h = 670
    }

    $x = 55
    $y = 15
    $ico_refresh = [System.Drawing.Image]::Fromfile($ico_refresh_path)

    # Main Window
    $win_main                	= New-Object System.Windows.Forms.Form
    $win_main.StartPosition  	= "CenterScreen"
    $win_main.Text           	= $win_text
    $win_main.Width          	= $win_w
    $win_main.Height         	= $win_h
    $win_main.ControlBox	   	= 1
    $win_main.Font           	= New-Object System.Drawing.Font("Arial",12)

    # Panel
    $panel			= new-object System.Windows.Forms.Panel
    $panel.Size 		= New-Object System.Drawing.Size(($win_w-240), ($win_h-230))
    $panel.Location 		= "5, 40"
    $panel.BackColor 		= "#C2C2C2"
    $panel.AutoScroll 	= $True
    $panel.Add_Mousehover( {$this.focus()})
    #$panel.add_mousehover({scrollMouseHover})


    $env_line_label 		= New-Object System.Windows.Forms.Label
    $env_line_label.Location  	= New-Object System.Drawing.Point((5),($win_h-180))
    $env_line_label.Width 		= ($win_w-240)
    $env_line_label.Height 	= 5
    $env_line_label.BackColor 	= $env_color

    ######################
    # FW BLOCK
    ######################
    # fw checkboxes
    $firewall_checkbox_all          	= New-Object System.Windows.Forms.CheckBox
    $firewall_checkbox_all.Location  	= New-Object System.Drawing.Point(($x),($y + 3))
    $firewall_checkbox_all.Checked     = 0
    $firewall_checkbox_all.Size 	= New-Object System.Drawing.Size(15,15)
    $firewall_checkbox_all.TabIndex = 1
    $firewall_checkbox_all.Add_CheckStateChanged({
        if ($firewall_checkbox_all.Checked -eq 1)
        {
            foreach($firewall_checkbox in $firewall_checkbox_arr) { $firewall_checkbox.Checked = 1 }
        }
        else
        {
            foreach($firewall_checkbox in $firewall_checkbox_arr) { $firewall_checkbox.Checked = 0 }
        }    
    })

    # fw caption
    $firewall_caption = New-Object System.Windows.Forms.Label
    $firewall_caption.Location  = New-Object System.Drawing.Point(($x + 20),($y + 3))
    $firewall_caption.Width = 55
    $firewall_caption.Height = 15 
    $firewall_caption.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3)              
    $firewall_caption.Text = "Firewall"

    # fw refresh    
    $fw_refresh_button   = New-Object System.Windows.Forms.Button
    $fw_refresh_button.Location = New-Object System.Drawing.Point(($x + 80), $y)
    $fw_refresh_button.Text     = ""
    $fw_refresh_button.Width 	= 21
    $fw_refresh_button.Height 	= 21
    $fw_refresh_button.TabIndex = 2
    $fw_refresh_button.BackgroundImage = $ico_refresh
    $fw_refresh_button.add_click({win_firewall})
    $fw_refresh_button.Autosize	= 1

    ######################
    # SERVICES BLOCK
    ######################
    # services checkboxes
    $services_checkbox_all          	= New-Object System.Windows.Forms.CheckBox
    $services_checkbox_all.Location  	= New-Object System.Drawing.Point(($x+210),($y + 3))
    $services_checkbox_all.Checked     = 0
    $services_checkbox_all.Size 	= New-Object System.Drawing.Size(15,15)
    $services_checkbox_all.TabIndex = 3
    $services_checkbox_all.Add_CheckStateChanged({
        if ($services_checkbox_all.Checked -eq 1)
        {
            foreach($services_checkbox in $services_checkbox_arr) { $services_checkbox.Checked = 1 }
        }
        else
        {
            foreach($services_checkbox in $services_checkbox_arr) { $services_checkbox.Checked = 0 }
        }    
    })

    # sp caption
    $sp_caption = New-Object System.Windows.Forms.Label
    $sp_caption.Location  = New-Object System.Drawing.Point(($x + 240),($y + 3))
    $sp_caption.Width = 30
    $sp_caption.Height = 15 
    $sp_caption.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3)              
    $sp_caption.Text = "S:P"

    # sp refresh 
    $sp_refresh_button   = New-Object System.Windows.Forms.Button
    $sp_refresh_button.Location = New-Object System.Drawing.Point(($x+280), $y)
    $sp_refresh_button.Text     = ""
    $sp_refresh_button.Width 	= 21
    $sp_refresh_button.Height 	= 21
    $sp_refresh_button.TabIndex = 4
    $sp_refresh_button.BackgroundImage = $ico_refresh
    $sp_refresh_button.add_click({win_services $group ; win_processes $group })
    $sp_refresh_button.Autosize	= 1
    #$fw_refresh_button.TabIndex = 2

    ######################
    # TASKS BLOCK
    ######################
 
    if((read_tasks $xmlConfig $server $name))
    {
        # tasks checkboxes
        $tasks_checkbox_all          	= New-Object System.Windows.Forms.CheckBox
        $tasks_checkbox_all.Location  	= New-Object System.Drawing.Point(($x+440),($y + 3))
        $tasks_checkbox_all.Checked     = 0
        $tasks_checkbox_all.Size 	= New-Object System.Drawing.Size(15,15)
        $tasks_checkbox_all.TabIndex = 5
        $tasks_checkbox_all.Add_CheckStateChanged({
        if ($tasks_checkbox_all.Checked -eq 1)
        {
            foreach($tasks_checkbox in $tasks_checkbox_arr ) 
            { 
                if($tasks_checkbox.Enabled)
                {
                    $tasks_checkbox.Checked = 1 
                }
            }
        }
        else
        {
            foreach($tasks_checkbox in $tasks_checkbox_arr) 
            {
                if($tasks_checkbox.Enabled)
                {
                    $tasks_checkbox.Checked = 0 
                }
            }
        }    
    })

        # tasks caption
        $tasks_caption = New-Object System.Windows.Forms.Label
        $tasks_caption.Location  = New-Object System.Drawing.Point(($x+470),($y + 3))
        $tasks_caption.Width = 40
        $tasks_caption.Height = 15 
        $tasks_caption.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3)              
        $tasks_caption.Text = "Tasks"

        # tasks refresh
        $tasks_refresh_button   = New-Object System.Windows.Forms.Button
        $tasks_refresh_button.Location = New-Object System.Drawing.Point(($x+520), $y)
        $tasks_refresh_button.Text     = ""
        $tasks_refresh_button.Width 	= 21
        $tasks_refresh_button.Height 	= 21
        $tasks_refresh_button.TabIndex = 6
        $tasks_refresh_button.BackgroundImage = $ico_refresh
        $tasks_refresh_button.add_click({win_tasks})
        $tasks_refresh_button.Autosize	= 1
    }

    ######################
    # SERVERS BLOCK
    ######################
    # servers checkboxes
    $servers_checkbox_all          	= New-Object System.Windows.Forms.CheckBox    
    $servers_checkbox_all.Checked     = 0
    $servers_checkbox_all.Size 	= New-Object System.Drawing.Size(15,15)
    $servers_checkbox_all.TabIndex = 7
    $servers_checkbox_all.Add_CheckStateChanged({
        if ($servers_checkbox_all.Checked -eq 1)
        {
            foreach($servers_checkbox in $servers_checkbox_arr) { $servers_checkbox.Checked = 1 }
        }
        else
        {
            foreach($servers_checkbox in $servers_checkbox_arr) { $servers_checkbox.Checked = 0 }
        }    
    })

    # servers caption
    $servers_caption = New-Object System.Windows.Forms.Label
    $servers_caption.Width = 60
    $servers_caption.Height = 15 
    $servers_caption.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3)              
    $servers_caption.Text = "Servers"

    # servers refresh
    $servers_refresh_button   = New-Object System.Windows.Forms.Button
    $servers_refresh_button.Text     = ""
    $servers_refresh_button.Width 	= 21
    $servers_refresh_button.Height 	= 21
    $servers_refresh_button.TabIndex = 8
    $servers_refresh_button.BackgroundImage = $ico_refresh
    $servers_refresh_button.Autosize	= 1


    if((read_tasks $xmlConfig $server $name))
    {
        $servers_checkbox_all.Location =      New-Object System.Drawing.Point(($x+615),($y+3))
        $servers_caption.Location  =          New-Object System.Drawing.Point(($x+635),$y)
        $servers_refresh_button.Location =    New-Object System.Drawing.Point(($x+695),$y)
        $servers_refresh_button.add_click({win_servers ($x+590) ($y); win_servers_wu ($x+590) ($y)})
        #win_servers ($x+590) ($y)  
    }
    else
    {
        $servers_checkbox_all.Location =      New-Object System.Drawing.Point(($x+410),($y + 3))
        $servers_caption.Location  =          New-Object System.Drawing.Point(($x+430),$y)
        $servers_refresh_button.Location =    New-Object System.Drawing.Point(($x+490),$y)
        $servers_refresh_button.add_click({win_servers ($x+385) ($y); win_servers_wu ($x+385) ($y)})
        #win_servers ($x+385) ($y)
    }

    ######################
    # FOLDERS BLOCK
    ######################  
    # folders caption
    $folders_caption = New-Object System.Windows.Forms.Label
    $folders_caption.Width = 60
    $folders_caption.Height = 15 
    $folders_caption.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3)              
    $folders_caption.Text = "Folders"

    $folders_refresh_button   = New-Object System.Windows.Forms.Button
    $folders_refresh_button.Text     = ""
    $folders_refresh_button.Width 	= 21
    $folders_refresh_button.Height 	= 21
    $folders_refresh_button.BackgroundImage = $ico_refresh
    $folders_refresh_button.Autosize	= 1
    $folders_refresh_button.TabIndex = 9

    if((read_tasks $xmlConfig $server $name))
    {
        $folders_caption.Location  =          New-Object System.Drawing.Point(($x+860),$y)
        $folders_refresh_button.Location =    New-Object System.Drawing.Point(($x+920),$y)
        $folders_refresh_button.add_click({win_folders ($x+825) ($y)})
        #win_folders ($x+825) ($y)    
    }
    else
    {
        $folders_caption.Location  =          New-Object System.Drawing.Point(($x+670),$y)
        $folders_refresh_button.Location =    New-Object System.Drawing.Point(($x+730),$y)
        $folders_refresh_button.add_click({win_folders ($x+630) ($y)})
        #win_folders ($x+630) ($y)
    }
    ######################
    # REFRESH ALL
    ###################### 
    $refresh_all_button			= New-Object System.Windows.Forms.Button
    
    $refresh_all_button.Text            = "Refresh All"
    $refresh_all_button.Width 		= 150
    $refresh_all_button.Height 		= 35
    $refresh_all_button.add_click(
                        {   win_firewall 
                            win_services $group 
                            win_processes $group 
                            if((read_tasks $xmlConfig $server $name))
                            {
                                win_tasks 
                                win_servers ($x+590) ($y)
                                win_folders ($x+825) ($y)
                            }
                            else
                            {
                                win_servers ($x+385) ($y)
                                win_folders ($x+630) ($y)
                            }
                        }
    )
    $refresh_all_button.Autosize        = 1
    $refresh_all_button.TabIndex        = 10
    $refresh_all_button.Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)
    if((read_tasks $xmlConfig $server $name))
    {
        $refresh_all_button.Location        = New-Object System.Drawing.Point(($x+860),($y+500))
    }
    else
    {
        $refresh_all_button.Location        = New-Object System.Drawing.Point(($x+650),($y+500))
    }

    ######################
    # PSVERSION BLOCK
    ###################### 
    # psversion caption
    $psversion_caption = New-Object System.Windows.Forms.Label
    $psversion_caption.Width = 70
    $psversion_caption.Height = 15 
    $psversion_caption.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3)              
    $psversion_caption.Text = "PS Version"
    if((read_tasks $xmlConfig $server $name))
    {
        $psversion_caption.Location  =          New-Object System.Drawing.Point(($x+1120),$y)
        win_psversion ($x+1060) ($y+20)
    }
    else
    {
        $psversion_caption.Location  =          New-Object System.Drawing.Point(($x+920),$y)
        win_psversion ($x+850) ($y+20)
    }

    ######################
    # BUTTONS CUSTOM BLOCK
    ######################
    if((read_tasks $xmlConfig $server $name))
    {
        win_buttons ($x+1100) ($y)
    }
    else
    {
        win_buttons ($x+880) ($y)
    } 
    

    # Firewall
    #win_firewall
    # Services
    #win_services $group
    # Processes
    #win_processes $group
    # Tasks
    #write-host (read_tasks $xmlConfig $server $name)
    if((read_tasks $xmlConfig $server $name))
    {
        #win_tasks
    }    



    $win_main.Controls.Add($firewall_checkbox_all)
    $win_main.Controls.Add($services_checkbox_all)
    $win_main.Controls.Add($tasks_checkbox_all)
    $win_main.Controls.Add($servers_checkbox_all)

    $win_main.Controls.Add($firewall_caption)
    $win_main.Controls.Add($sp_caption)
    $win_main.Controls.Add($tasks_caption)
    $win_main.Controls.Add($servers_caption)
    $win_main.Controls.Add($folders_caption)
    $win_main.Controls.Add($psversion_caption)

    $win_main.Controls.Add($folders_refresh_button)
    $win_main.Controls.Add($fw_refresh_button)
    $win_main.Controls.Add($sp_refresh_button)
    $win_main.Controls.Add($tasks_refresh_button)
    $win_main.Controls.Add($servers_refresh_button)
    $win_main.Controls.Add($folders_refresh_button)
    $win_main.Controls.Add($refresh_all_button) 


    $win_main.Controls.Add($panel)
    $win_main.Controls.Add($env_line_label)

    $win_main.Add_Shown({$win_main.Activate()})
    $win_main.ShowDialog() | Out-Null
}
function win_firewall
{
    $x = 30
    $y = 15
    $tab_index = 100

    foreach($firewall_label in $script:firewall_label_arr) { $panel.Controls.Remove($firewall_label) }
    foreach($firewall_checkbox in $script:firewall_checkbox_arr) { $panel.Controls.Remove($firewall_checkbox) }
    $panel.AutoScroll 	= $False
    $panel.AutoScroll 	= $True

    $script:firewall_label_arr = @()
    $script:firewall_checkbox_arr = @()

    #$script:firewall_label_arr.Clear()
    #script:firewall_checkbox_arr.Clear()

    $cmd = "getfw"        
    $results = workflow_common -data (read_firewall $xmlConfig $server $name) -cmd $cmd
    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}

    foreach($result in $results)
    {
        $script:firewall_label_arr += New-Object System.Windows.Forms.Label
        $script:firewall_label_arr[$($result.Index)].Location  = New-Object System.Drawing.Point($x,($y+$($result.Index)*25))
        $script:firewall_label_arr[$($result.Index)].Width = 15
        $script:firewall_label_arr[$($result.Index)].Height = 15            
        $script:firewall_label_arr[$($result.Index)].BackColor = (status_color $result.Status)
        
        $script:firewall_checkbox_arr += New-Object System.Windows.Forms.CheckBox
        $script:firewall_checkbox_arr[$($result.Index)].Location      = New-Object System.Drawing.Point(($x+20),($y+$($result.Index)*25))
        $script:firewall_checkbox_arr[$($result.Index)].Text          = $result.Desc
        $script:firewall_checkbox_arr[$($result.Index)].Size = New-Object System.Drawing.Size(150,15) 
        $script:firewall_checkbox_arr[$($result.Index)].Font = "lucida console,8"
        $script:firewall_checkbox_arr[$($result.Index)].TabIndex      = ($tab_index + $result.Index)
        $script:firewall_checkbox_arr[$($result.Index)].Name = @($result.Server, $result.Name)  
        
        $panel.Controls.Add($script:firewall_label_arr[$($result.Index)])
        $panel.Controls.Add($script:firewall_checkbox_arr[$($result.Index)])
        $message = "firewall $($result.Server) $($result.Name) - $($result.Status)"
        write_log "WIN: $message"
    }

    ######################
    # SET FW BUTTON
    ######################
    $script:firewall_set_button			= New-Object System.Windows.Forms.Button
    $script:firewall_set_button.Location        = New-Object System.Drawing.Point(($x-5),($y+500))
    $script:firewall_set_button.Text            = "Set Firewall"
    $script:firewall_set_button.Width 		= 150
    $script:firewall_set_button.Height 		= 35
    $script:firewall_set_button.add_click({ win_firewall_set})
    $script:firewall_set_button.Autosize        = 1
    $script:firewall_set_button.TabIndex        = ($tab_index + $result.Index + 1)
    $script:firewall_set_button.Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)

    if(!$results)
    {
        $script:firewall_set_button.Enabled = $false
    }

    $win_main.Controls.Add($script:firewall_set_button)

}
function win_firewall_set
{
    foreach($firewall_checkbox in ($script:firewall_checkbox_arr| Where-Object {$_.Checked -eq $true}) )
    {
        $server = (($firewall_checkbox.Name).Split(" "))[0]
        #$name = (($firewall_checkbox.Name).Split(" "))[1]
        $name = ($firewall_checkbox.Name).Substring(($firewall_checkbox.Name).IndexOf(" ")+1, ($firewall_checkbox.Name).Length-($firewall_checkbox.Name).IndexOf(" ")-1)
        $cmd = "setfw"   
        
        $message = "$server $name $cmd"
        write_log "WIN: $message"
           
        $results = workflow_common -data (read_firewall $xmlConfig $server $name) -cmd $cmd
        $results = $null
    }
    $server = $null
    $name = $null
    win_firewall
}
function win_services($group)
{
    $x = 240
    $y = 15
    $script:y_services = $y   
    $tab_index = 200

    foreach($services_label in $script:services_label_arr) { $panel.Controls.Remove($services_label) }
    foreach($services_checkbox in $script:services_checkbox_arr) { $panel.Controls.Remove($services_checkbox) }
    $panel.AutoScroll 	= $False
    $panel.AutoScroll 	= $True

    $script:services_label_arr = @()
    $script:services_checkbox_arr = @()
    #$script:services_label_arr.Clear()
    #$script:services_checkbox_arr.Clear()
  
    $cmd = "getservices"        
    $results = workflow_common -data (read_services $xmlConfig $server $name) -cmd $cmd
 
    if($results)
    {
        if($group -eq 1)
        {
            $k = 0
            $results = $results | Group-Object { $_.Server } | Sort-Object -Property Name
            for($i=0;$i -lt $results.Count; $i++)
            {
                $script:services_checkbox_arr += New-Object System.Windows.Forms.CheckBox
                $script:services_checkbox_arr[$i].Location      = New-Object System.Drawing.Point(($x+20),($y+$i*50))
                $script:services_checkbox_arr[$i].Text          = $results[$i].Name
                $script:services_checkbox_arr[$i].Size = New-Object System.Drawing.Size(150,15) 
                $script:services_checkbox_arr[$i].Font = "lucida console,8"
                $script:services_checkbox_arr[$i].TabIndex      = ($tab_index + $i)
                $script:services_checkbox_arr[$i].Name = $results[$i].Name

                for($j=0;$j -lt ($results[$i]).Count; $j++)
                {
                    $script:services_label_arr += New-Object System.Windows.Forms.Label
                    $script:services_label_arr[$k].Location  = New-Object System.Drawing.Point(($x+40+$j*20),($y+20+$i*50))
                    $script:services_label_arr[$k].Width = 15
                    $script:services_label_arr[$k].Height = 15 
                    $script:services_label_arr[$k].BackColor = (status_color $results[$i].Group[$j].Status)
                    $script:services_label_arr[$k].Font = "lucida console,8"
                    $script:services_label_arr[$k].TextAlign = "MiddleCenter"
                    $script:services_label_arr[$k].Text = ($results[$i].Group[$j].Desc)[0] # first letter

                    $message = "service $($results[$i].Name) $($results[$i].Group[$j].Name) - $($results[$i].Group[$j].Status)"
                    write_log "WIN: $message"

                    $panel.Controls.Add($script:services_label_arr[$k])
                    $k++                
                }      
                $panel.Controls.Add($script:services_checkbox_arr[$i])

            }
            $script:y_services =  $y+$i*50
        }
        else
        {
            $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}

            foreach($result in $results)
            {

                $script:services_label_arr += New-Object System.Windows.Forms.Label
                $script:services_label_arr[$($result.Index)].Location  = New-Object System.Drawing.Point($x,($y+$($result.Index)*25))
                $script:services_label_arr[$($result.Index)].Width = 15
                $script:services_label_arr[$($result.Index)].Height = 15            
                $script:services_label_arr[$($result.Index)].BackColor = (status_color $result.Status)
        
                $script:services_checkbox_arr += New-Object System.Windows.Forms.CheckBox
                $script:services_checkbox_arr[$($result.Index)].Location      = New-Object System.Drawing.Point(($x+20),($y+$($result.Index)*25))
                $script:services_checkbox_arr[$($result.Index)].Text          = $result.Desc
                $script:services_checkbox_arr[$($result.Index)].Size = New-Object System.Drawing.Size(150,15) 
                $script:services_checkbox_arr[$($result.Index)].Font = "lucida console,8"
                $script:services_checkbox_arr[$($result.Index)].Tag = $result.Index
                $script:services_checkbox_arr[$($result.Index)].TabIndex      = ($tab_index + $result.Index)
                $script:services_checkbox_arr[$($result.Index)].Name = @($result.Server, $result.Name)  
        
                $message = "service $($result.Server) $($result.Name) - $($result.Status)"
                write_log "WIN: $message"

                $panel.Controls.Add($script:services_label_arr[$($result.Index)])
                $panel.Controls.Add($script:services_checkbox_arr[$($result.Index)])
            }
            $script:y_services =  $y+$($result.Index)*25+20
        }
    }

    ######################
    # START SERVICES BUTTON
    ######################
    $script:services_start_button			= New-Object System.Windows.Forms.Button
    $script:services_start_button.Location        = New-Object System.Drawing.Point(($x),($y+500))
    $script:services_start_button.Text            = "Start Services"
    $script:services_start_button.Width 		= 150
    
    $script:services_start_button.Height 		= 35
    $script:services_start_button.add_click({ win_services_set "startservices" $group } )
    $script:services_start_button.Autosize        = 1
    $script:services_start_button.TabIndex        = ($tab_index + $result.Index + 1)
    $script:services_start_button.Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)

    ######################
    # STOP SERVICES BUTTON
    ######################
    $script:services_stop_button			= New-Object System.Windows.Forms.Button
    $script:services_stop_button.Location        = New-Object System.Drawing.Point(($x),($y+550))
    $script:services_stop_button.Text            = "Stop Services"
    $script:services_stop_button.Width 		= 150
    $script:services_stop_button.Height 		= 35
    $script:services_stop_button.add_click({ win_services_set "stopservices" $group} )
    $script:services_stop_button.Autosize        = 1
    $script:services_stop_button.TabIndex        = ($tab_index + $result.Index + 2)
    $script:services_stop_button.Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)

    if(!$results)
    {
        $script:services_start_button.Enabled = $false
        $script:services_stop_button.Enabled = $false
    }

    $win_main.Controls.Add($script:services_stop_button) 
    $win_main.Controls.Add($script:services_start_button)
    
}
function win_services_set($cmd, $group)
{
    $confirm = "Yes"

    if($warning -eq 1)
    {
        $confirm = $null
        if($cmd -eq "stopservices")  {$text = "Stop Services"}
        if($cmd -eq "startservices") {$text = "Start Services"}
        $confirm = [System.Windows.Forms.MessageBox]::Show("$text ?", $text , "YesNo","Warning")
    }
 
    if($confirm -eq "Yes")
    {
        if($cmd -eq "stopservices")
        {     
            $script:services_checkbox_arr = $script:services_checkbox_arr | Sort-Object -Property Tag -Descending
        }

        foreach($services_checkbox in ($script:services_checkbox_arr| Where-Object {$_.Checked -eq $true}) )
        {
            $server = (($services_checkbox.Name).Split(" "))[0]
            if($group -eq 1)
            {
                $name = $null
            }
            else
            {
                $name = ($services_checkbox.Name).Substring(($services_checkbox.Name).IndexOf(" ")+1, ($services_checkbox.Name).Length-($services_checkbox.Name).IndexOf(" ")-1)
            }

            $message = "$server $name $cmd"
            write_log "WIN: $message"
  
            if($cmd -eq "stopservices")
            {
                #$results = workflow_common -data (read_services $xmlConfig $server $name) -cmd $cmd; $results = $null
                $data = read_services $xmlConfig $server $name
                #$data = $data | Sort-Object @{Expression={$_.Index}; Ascending=$false}
                        
                foreach($record in $data)
                {
                    $results = workflow_common -data $record -cmd $cmd
                    $results = $null
                    if($record.attr2 -notin ("true", "false") -and $record.attr2)
                    {
                        Start-Sleep -s $record.attr2
                    }
                }
            }
            elseif($cmd -eq "startservices")
            {
                #$results = workflow_common -data (read_services $xmlConfig $server $name) -cmd $cmd; $results = $null
                $data = read_services $xmlConfig $server $name

                foreach($record in $data)
                {                           
                    if($record.attr1 -eq "true")
                    {
                        $results = workflow_common -data $record -cmd $cmd
                        $results = $null
                    }
                    else
                    {
                        $results = Invoke-Command -ComputerName $data.Server  -ErrorAction Stop -ArgumentList $record.Name -ScriptBlock {param($service) Start-Service -Name $service } -AsJob 
                        $results = $null
                        if($record.attr1 -ne "false" -and $record.attr1)
                        {
                            Start-Sleep -s $record.attr1
                        }
                    }
                }
            }    
        }

        $server = $null
        $name = $null
        win_services $group
    }
}
function win_processes($group)
{
    $x = 240
    if($script:y_services -eq 15) # no services
    {
        $y = 15  
    }
    else
    {
        $y = $script:y_services+30  
    }

    $tab_index = 300

    foreach($servers_label in $script:servers_label_arr) { $panel.Controls.Remove($servers_label) }
    foreach($processes_label in $script:processes_label_arr) { $panel.Controls.Remove($processes_label) }
    foreach($processes_label_count in $script:processes_label_count_arr) { $panel.Controls.Remove($processes_label_count) }
    

    $script:servers_label_arr = @()
    $script:processes_label_count_arr = @()
    $script:processes_label_arr = @()
    
    $cmd = "getprocesses"        
    $results = workflow_common -data (read_processes $xmlConfig $server $name) -cmd $cmd

    if($group -eq 1)
    {
        $k = 0
        $results = $results | Group-Object { $_.Server } | Sort-Object -Property Name
        for($i=0;$i -lt $results.Count; $i++)
        {
            $script:servers_label_arr += New-Object System.Windows.Forms.Label
            $script:servers_label_arr[$i].Location      = New-Object System.Drawing.Point(($x+10),($y+$i*50))
            $script:servers_label_arr[$i].Width = 150
            $script:servers_label_arr[$i].Height = 15 
            $script:servers_label_arr[$i].Font = "lucida console,8"
            $script:servers_label_arr[$i].Text = $results[$i].Name

            for($j=0;$j -lt ($results[$i]).Count; $j++)
            {            
                $status = $null
                $count =  $(($($results[$i].Group[$j].Status)).Split(" ")).Count        
                if($results[$i].Group[$j].Status -eq "n/a"){$status = "Error"; $count = 0}
                
                $script:processes_label_count_arr += New-Object System.Windows.Forms.Label
                $script:processes_label_count_arr[$k].Location  = New-Object System.Drawing.Point(($x+$j*40),($y+20+$i*50))
                $script:processes_label_count_arr[$k].Width = 25
                $script:processes_label_count_arr[$k].Height = 15 
                #$script:processes_label_count_arr[$j].BackColor = (status_color $status)
                $script:processes_label_count_arr[$k].Font = "lucida console,8"
                $script:processes_label_count_arr[$k].TextAlign = "MiddleRight"
                $script:processes_label_count_arr[$k].Text = "$($count):" # count of process
                $panel.Controls.Add($script:processes_label_count_arr[$k]) 

                $script:processes_label_arr += New-Object System.Windows.Forms.Label
                $script:processes_label_arr[$k].Location  = New-Object System.Drawing.Point(($x+25+$j*40),($y+20+$i*50))
                $script:processes_label_arr[$k].Width = 15
                $script:processes_label_arr[$k].Height = 15 
                $script:processes_label_arr[$k].BackColor = (status_color $status)
                $script:processes_label_arr[$k].Font = "lucida console,8"
                $script:processes_label_arr[$k].TextAlign = "MiddleCenter"
                $script:processes_label_arr[$k].Text = ($results[$i].Group[$j].Desc)[0] # first letter
                $panel.Controls.Add($script:processes_label_arr[$k])

                $message = "process $($results[$i].Name) $($results[$i].Group[$j].Name) - $count"
                write_log "WIN: $message"

                $k++                    
                      
            }      
            $panel.Controls.Add($script:servers_label_arr[$i])
        }
    }
    else
    {
        $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}

        foreach($result in $results)
        {
            $k = 0
            foreach($status in ($result.Status).Split(" ") )
            {           
                $k++
                $status = $null              
                if($result.Status -eq "n/a"){$status = "Error"}

                $script:processes_label_count_arr += New-Object System.Windows.Forms.Label
                $script:processes_label_count_arr[$($result.Index)].Location  = New-Object System.Drawing.Point(($x),($y+$($result.Index)*25))
                $script:processes_label_count_arr[$($result.Index)].Width = 15
                $script:processes_label_count_arr[$($result.Index)].Height = 15   
                $script:processes_label_count_arr[$($result.Index)].Text = "•"
                $script:processes_label_count_arr[$($result.Index)].Padding = 0
                $script:processes_label_count_arr[$($result.Index)].TextAlign = "MiddleCenter"
                $script:processes_label_count_arr[$($result.Index)].Font = "lucida console,28"
                $script:processes_label_count_arr[$($result.Index)].ForeColor = (status_color $status)

                $script:processes_label_arr += New-Object System.Windows.Forms.Label
                $script:processes_label_arr[$($result.Index)].Location  = New-Object System.Drawing.Point(($x+20),($y+$($result.Index)*25))
                $script:processes_label_arr[$($result.Index)].Width = 150
                $script:processes_label_arr[$($result.Index)].Height = 15   
                $script:processes_label_arr[$($result.Index)].Text = $($result.Desc) + $(if($k -ne 1){" [$k]"})
                $script:processes_label_arr[$($result.Index)].Padding = 0
                $script:processes_label_arr[$($result.Index)].Font = "lucida console,8"            

                $message = "process $($result.Server) $($result.Name) - $k"
                write_log "WIN: $message"
                
                $panel.Controls.Add($script:processes_label_arr[$($result.Index)])
                $panel.Controls.Add($script:processes_label_count_arr[$($result.Index)])

            }
        }
    }
}
function win_tasks
{
    $x = 470
    $y = 15
    $tab_index = 400

    foreach($task_label in $script:tasks_label_arr) { $panel.Controls.Remove($task_label) }
    foreach($task_checkbox in $script:tasks_checkbox_arr) { $panel.Controls.Remove($task_checkbox) }
    $panel.AutoScroll 	= $False
    $panel.AutoScroll 	= $True

    $script:tasks_label_arr = @()
    $script:tasks_checkbox_arr = @()

    $cmd = "gettasks"        
    $results = workflow_common -data (read_tasks $xmlConfig $server $name) -cmd $cmd
    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}

    foreach($result in $results)
    {
        $script:tasks_label_arr += New-Object System.Windows.Forms.Label
        $script:tasks_label_arr[$($result.Index)].Location  = New-Object System.Drawing.Point($x,($y+$($result.Index)*25))
        $script:tasks_label_arr[$($result.Index)].Width = 15
        $script:tasks_label_arr[$($result.Index)].Height = 15            
        $script:tasks_label_arr[$($result.Index)].BackColor = (status_color $result.Status)
        
        $script:tasks_checkbox_arr += New-Object System.Windows.Forms.CheckBox
        $script:tasks_checkbox_arr[$($result.Index)].Location      = New-Object System.Drawing.Point(($x+20),($y+$($result.Index)*25))
        $script:tasks_checkbox_arr[$($result.Index)].Text          = $result.Desc
        $script:tasks_checkbox_arr[$($result.Index)].Size = New-Object System.Drawing.Size(150,15) 
        $script:tasks_checkbox_arr[$($result.Index)].Font = "lucida console,8"
        $script:tasks_checkbox_arr[$($result.Index)].Enabled = $result.Attr2
        $script:tasks_checkbox_arr[$($result.Index)].TabIndex      = ($tab_index + $result.Index)
        $script:tasks_checkbox_arr[$($result.Index)].Name = @($result.Server, $result.Name)  

        $message = "task $($result.Server) $($result.Name) - $($result.Status)"
        write_log "WIN: $message"
        
        $panel.Controls.Add($script:tasks_label_arr[$($result.Index)])
        $panel.Controls.Add($script:tasks_checkbox_arr[$($result.Index)])
    }

    ######################
    # START TASKS BUTTON
    ######################
    $script:tasks_start_button			= New-Object System.Windows.Forms.Button
    $script:tasks_start_button.Location        = New-Object System.Drawing.Point(($x),($y+500))
    $script:tasks_start_button.Text            = "Start Task"
    $script:tasks_start_button.Width 		= 150
    $script:tasks_start_button.Height 		= 35
    $script:tasks_start_button.add_click({ win_tasks_set "starttasks"})
    $script:tasks_start_button.Autosize        = 1
    $script:tasks_start_button.TabIndex        = ($tab_index + $result.Index + 1)
    $script:tasks_start_button.Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)

    ######################
    # STOP TASKS BUTTON
    ######################
    $script:tasks_stop_button			= New-Object System.Windows.Forms.Button
    $script:tasks_stop_button.Location        = New-Object System.Drawing.Point(($x),($y+550))
    $script:tasks_stop_button.Text            = "Stop Tasks"
    $script:tasks_stop_button.Width 		= 150
    $script:tasks_stop_button.Height 		= 35
    $script:tasks_stop_button.add_click({ win_tasks_set "stoptasks"} )
    $script:tasks_stop_button.Autosize        = 1
    $script:tasks_stop_button.TabIndex        = ($tab_index + $result.Index + 2)
    $script:tasks_stop_button.Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)

    $win_main.Controls.Add($script:tasks_start_button)    
    $win_main.Controls.Add($script:tasks_stop_button)    
}
function win_tasks_set($cmd)
{
    $confirm = "Yes"

    if($warning -eq 1)
    {
        $confirm = $null
        $confirm = [System.Windows.Forms.MessageBox]::Show("Run Tasks ?", "Run Tasks" , "YesNo","Warning")
    }
 
    if($confirm -eq "Yes")
    {
        foreach($tasks_checkbox in ($script:tasks_checkbox_arr| Where-Object {$_.Checked -eq $true}) )
        {
            $server = (($tasks_checkbox.Name).Split(" "))[0]
            #$name = (($tasks_checkbox.Name).Split(" "))[1]
            $name = ($tasks_checkbox.Name).Substring(($tasks_checkbox.Name).IndexOf(" ")+1, ($tasks_checkbox.Name).Length-($tasks_checkbox.Name).IndexOf(" ")-1)

            $message = "$server $name $cmd"
            write_log "WIN: $message"

            $results =         workflow_common -data (read_tasks $xmlConfig $server $name) -cmd $cmd
            $results = $null
        }
        $server = $null
        $name = $null
        win_tasks
    }
}
function win_servers($x, $y)
{
    $tab_index = 500

    foreach($server_checkbox in $script:servers_checkbox_arr) { $panel.Controls.Remove($server_checkbox) }
    $panel.AutoScroll 	= $False
    $panel.AutoScroll 	= $True
    

    $script:servers_checkbox_arr = @()

    $cmd = "getservers"        
    $results = workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd
    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}

    foreach($result in $results)
    {       
        $script:servers_checkbox_arr += New-Object System.Windows.Forms.CheckBox
        if($wu_enabled -eq 1)
        {
            $script:servers_checkbox_arr[$($result.Index)].Location      = New-Object System.Drawing.Point(($x+20),($y+$($result.Index)*40)) #25
        }
        else
        {
            $script:servers_checkbox_arr[$($result.Index)].Location      = New-Object System.Drawing.Point(($x+20),($y+$($result.Index)*25))
        }
        $script:servers_checkbox_arr[$($result.Index)].Text          = "$($result.Name) - $($result.Status)"
        $script:servers_checkbox_arr[$($result.Index)].Size = New-Object System.Drawing.Size(190,15) 
        $script:servers_checkbox_arr[$($result.Index)].Font = "lucida console,8"
        #$script:servers_checkbox_arr[$($result.Index)].BackColor = "Green"
        $script:servers_checkbox_arr[$($result.Index)].TabIndex      = ($tab_index + $result.Index)
        $script:servers_checkbox_arr[$($result.Index)].Name = $result.Server

        $panel.Controls.Add($script:servers_checkbox_arr[$($result.Index)])
        $message = "server $($result.Name) - $($result.Status)"
        write_log "WIN: $message"  
    }

    ######################
    # REBOOT BUTTON
    ######################
    $script:servers_reboot_button			= New-Object System.Windows.Forms.Button
    $script:servers_reboot_button.Location        = New-Object System.Drawing.Point(($x+40),($y+500))
    $script:servers_reboot_button.Text            = "Reboot"
    $script:servers_reboot_button.Width 		= 150
    $script:servers_reboot_button.Height 		= 35
    $script:servers_reboot_button.add_click({ win_servers_set "reboot"})
    $script:servers_reboot_button.Autosize        = 1
    $script:servers_reboot_button.TabIndex        = ($tab_index + $result.Index + 1)
    $script:servers_reboot_button.Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)

    $win_main.Controls.Add($script:servers_reboot_button)      
}
function win_servers_wu($x, $y)
{
    if($wu_enabled -eq 1)
    {
        $tab_index = 600
        $ico_wu = [System.Drawing.Image]::Fromfile($ico_wu_path)

        foreach($server_wu_label in $script:server_wu_label_arr) { $panel.Controls.Remove($server_wu_label) }
        foreach($server_wu_ico_label in $script:server_wu_ico_label_arr) { $panel.Controls.Remove($server_wu_ico_label) }
        foreach($server_wu_result_label in $script:server_wu_result_label_arr) { $panel.Controls.Remove($server_wu_result_label) }
        $panel.AutoScroll 	= $False
        $panel.AutoScroll 	= $True
    
        $script:server_wu_label_arr = @()
        $script:server_wu_ico_label_arr = @()
        $script:server_wu_result_label_arr = @()

        $cmd = "getwu"        
        $results = workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd
        $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}

        foreach($result in $results)
                                                                                                                                                    {       
        $script:server_wu_ico_label_arr += New-Object System.Windows.Forms.Label   
        $script:server_wu_label_arr += New-Object System.Windows.Forms.Label 
        $script:server_wu_result_label_arr += New-Object System.Windows.Forms.Label   

        $script:server_wu_ico_label_arr[$($result.Index)].Location  = New-Object System.Drawing.Point(($x+25),($y+20+$($result.Index)*40))
        $script:server_wu_ico_label_arr[$($result.Index)].Width = 16
        $script:server_wu_ico_label_arr[$($result.Index)].Height = 16 
        $script:server_wu_ico_label_arr[$($result.Index)].BackgroundImage = $ico_wu
        if($result.Desc -eq "Running")
        {
            $script:server_wu_ico_label_arr[$($result.Index)].BackColor = "Cyan"
        }

        $script:server_wu_label_arr[$($result.Index)].Location  = New-Object System.Drawing.Point(($x+45),($y+20+$($result.Index)*40))
        $script:server_wu_label_arr[$($result.Index)].Width = 90
        $script:server_wu_label_arr[$($result.Index)].Height = 15 
        $script:server_wu_label_arr[$($result.Index)].Font = New-Object System.Drawing.Font("Microsoft Sans Serif",7,0,3)
        $script:server_wu_label_arr[$($result.Index)].ForeColor = "#1A5599"
        $script:server_wu_label_arr[$($result.Index)].Text = "Windows Updates : "

        $script:server_wu_result_label_arr[$($result.Index)].Location  = New-Object System.Drawing.Point(($x+135),($y+20+$($result.Index)*40))
        $script:server_wu_result_label_arr[$($result.Index)].Width = 30
        $script:server_wu_result_label_arr[$($result.Index)].Height = 15 
        $script:server_wu_result_label_arr[$($result.Index)].Font = New-Object System.Drawing.Font("Microsoft Sans Serif",7,0,3)
        switch($($results[$($result.Index)].Status))
        {
            "n/a"  {$script:server_wu_result_label_arr[$($result.Index)].ForeColor = "Gray"}
            0      {$script:server_wu_result_label_arr[$($result.Index)].ForeColor = "#1A5599"}
            default{$script:server_wu_result_label_arr[$($result.Index)].ForeColor = "Red"} 
        }
        
        $script:server_wu_result_label_arr[$($result.Index)].Text = "$($results[$($result.Index)].Status)"

        $panel.Controls.Add($script:server_wu_label_arr[$($result.Index)])
        $panel.Controls.Add($script:server_wu_ico_label_arr[$($result.Index)])
        $panel.Controls.Add($script:server_wu_result_label_arr[$($result.Index)])
        $message = "server $($result.Server) : $($result.Status)"
        write_log "WIN: $message"  
    } 
    
        ######################
        # INSTALL WU BUTTON
        ######################
        $script:servers_wui_button			= New-Object System.Windows.Forms.Button
        $script:servers_wui_button.Location        = New-Object System.Drawing.Point(($x+40),($y+550))
        $script:servers_wui_button.Text            = "Install WU"
        $script:servers_wui_button.Width 		= 150
        $script:servers_wui_button.Height 		= 35
        $script:servers_wui_button.add_click({ win_servers_wu_set "startwu"})
        $script:servers_wui_button.Autosize        = 1
        $script:servers_wui_button.TabIndex        = ($tab_index + $result.Index + 1)
        $script:servers_wui_button.Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)

        $win_main.Controls.Add($script:servers_wui_button)
    } 
}
function win_servers_set($cmd)
{
    $confirm = "Yes"

    if($warning -eq 1)
    {
        $confirm = $null
        $confirm = [System.Windows.Forms.MessageBox]::Show("Reboot Servers ?", "Reboot Servers" , "YesNo","Warning")
    }
 
    if($confirm -eq "Yes")
    {
        foreach($servers_checkbox in ($script:servers_checkbox_arr| Where-Object {$_.Checked -eq $true}) )
        {
            $server = $servers_checkbox.Name

            $message = "$server $cmd"
            write_log "WIN: $message"

            $results =         workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd
            $results = $null
        }
        $server = $null
        #win_servers
    }
}
function win_servers_wu_set($cmd)
{
    foreach($servers_checkbox in ($script:servers_checkbox_arr| Where-Object {$_.Checked -eq $true}) )
    {
        $server = $servers_checkbox.Name

        $message = "$server $cmd"
        write_log "WIN: $message"

        $results =         workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd
        $results = $null
    }
    $server = $null
    #win_servers_wu


    if((read_tasks $xmlConfig $server $name))
    {
        win_servers_wu ($x+590) ($y)
    }
    else
    {
        win_servers_wu ($x+385) ($y)
    }
}
function win_folders($x, $y)
{
    $tab_index = 700

    foreach($folder_name_label in $script:folder_name_label_arr) { $panel.Controls.Remove($folder_name_label) }
    foreach($folder_info_label in $script:folder_info_label_arr) { $panel.Controls.Remove($folder_info_label) }
    foreach($folder_line_label in $script:folder_line_label_arr) { $panel.Controls.Remove($folder_line_label) }
    $panel.AutoScroll 	= $False
    $panel.AutoScroll 	= $True

    $script:folder_name_label_arr = @()
    $script:folder_info_label_arr = @()
    $script:folder_line_label_arr = @()
    
    $results = workflow_folders -data (read_folders $xmlConfig $name)
    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}, @{Expression={$_.Id}; Ascending=$True}

    #foreach($result in $results)
    for($i=0;$i -lt $results.Count; $i++)
    {
        
        $script:folder_name_label_arr += New-Object System.Windows.Forms.Label   
        $script:folder_info_label_arr += New-Object System.Windows.Forms.Label 
        
        if($results[$i].Id -eq 0)
        {
            # first line
            if($i -ne 0){$y = $y+10}
            $main_index = $result.Index
            $main_size = $results[$i].Size
            $main_count = $results[$i].Quantity
            
                            
            $script:folder_name_label_arr[$i].Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8,0,3)  
            $script:folder_info_label_arr[$i].Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8,0,3)                  
        }
        else
        {        
            $script:folder_name_label_arr[$i].Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8,2,3)
            $script:folder_info_label_arr[$i].Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8,0,3)

            if( $main_size -eq $($results[$i].Size) -and $main_count -eq $($results[$i].Quantity) )
            {
                $script:folder_name_label_arr[$i].ForeColor = "#3B3B3B"
                $script:folder_info_label_arr[$i].ForeColor = "#3B3B3B"

            }
            else
            {
                $script:folder_name_label_arr[$i].ForeColor = "Red"
                $script:folder_info_label_arr[$i].ForeColor = "Red"
            }
        }

        $script:folder_name_label_arr[$i].Location  = New-Object System.Drawing.Point($x,($y+$i*20))
        $script:folder_name_label_arr[$i].Width = 100
        $script:folder_name_label_arr[$i].Height = 15 
        $script:folder_name_label_arr[$i].Text = $results[$i].Desc + ": " 

        $script:folder_info_label_arr[$i].Location  = New-Object System.Drawing.Point(($x+100),($y+$i*20))
        $script:folder_info_label_arr[$i].Width = 110
        $script:folder_info_label_arr[$i].Height = 15 
        $script:folder_info_label_arr[$i].Text = "$($results[$i].Quantity)  [$([math]::Round($($results[$i].Size),2)) MB]" 
      
        $message = "folder $($results[$i].Server) $($results[$i].Path) - $($results[$i].Quantity), $($results[$i].Size) B"
        write_log "WIN: $message"

        $panel.Controls.Add($script:folder_name_label_arr[$i])
        $panel.Controls.Add($script:folder_info_label_arr[$i])

    }
}
function win_psversion($x, $y)
{
    $script:servers_ps_arr = @()
    $script:version_ps_arr = @()

    $cmd = "getps"        
    $results = workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd
    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}

    foreach($result in $results)
    {       
        $script:servers_ps_arr += New-Object System.Windows.Forms.Label
        $script:servers_ps_arr[$($result.Index)].Location      = New-Object System.Drawing.Point(($x+20),($y+10+$($result.Index)*15))
        $script:servers_ps_arr[$($result.Index)].Text          = "$($result.Name):"
        $script:servers_ps_arr[$($result.Index)].Size = New-Object System.Drawing.Size(100,15) 
        $script:servers_ps_arr[$($result.Index)].Font = "Microsoft Sans Serif,7"

        $script:version_ps_arr += New-Object System.Windows.Forms.Label
        $script:version_ps_arr[$($result.Index)].Location      = New-Object System.Drawing.Point(($x+120),($y+10+$($result.Index)*15))
        $script:version_ps_arr[$($result.Index)].Text          = "$($result.Status)"
        $script:version_ps_arr[$($result.Index)].Size = New-Object System.Drawing.Size(70,15) 
        $script:version_ps_arr[$($result.Index)].Font = "Microsoft Sans Serif,7"

        $message = "psversion $($result.Server) - $($result.Status)" 
        write_log "WIN: $message"

        $win_main.Controls.Add($script:servers_ps_arr[$($result.Index)])
        $win_main.Controls.Add($script:version_ps_arr[$($result.Index)])
    }    
}
function win_buttons($x, $y)
{
    $tab_index = 800
    $script:buttons_arr = @()   
    $results = read_buttons $xmlConfig

    foreach($result in $results)
    {        

        $script:buttons_arr			+= New-Object System.Windows.Forms.Button
        if(($results.Index).Count -eq 1)
        {
            $script:buttons_arr[$result.Index].Location        = New-Object System.Drawing.Point(($x),($y+500))
        }
        else
        {
            $script:buttons_arr[$result.Index].Location        = New-Object System.Drawing.Point(($x),( ($y+550) - ($result.Index)*35 - ($result.Index)*15 ) )                      
        }
        $script:buttons_arr[$result.Index].Height 		= 35
        $script:buttons_arr[$result.Index].Text            = $result.Name
        $script:buttons_arr[$result.Index].Width 		= 150 
        $script:buttons_arr[$result.Index].Tag =       $($result.Index)  
        $script:buttons_arr[$result.Index].add_click({ win_buttons_set $this.Tag} )
        $script:buttons_arr[$result.Index].Autosize        = 1
        $script:buttons_arr[$result.Index].TabIndex        = ($tab_index + $result.Index)
        $script:buttons_arr[$result.Index].Font 		= New-Object System.Drawing.Font("Microsoft Sans Serif",11,1,3)  

        $message = "button $($result.Name) - '$($result.Cmd)'" 
        write_log "WIN: $message"            
        $win_main.Controls.Add($script:buttons_arr[$result.Index])      
    }
}
function win_buttons_set ($index)
{
    # IF THERE IS ONLY 1 BUTTON
    if(($xmlConfig.config.buttons.button).Count -ge 2)
    {
        $xmlConfig = ($xmlConfig.config.buttons.button)[$index] # | Where-Object {$_.name -eq $index } }
    }
    else
    {
        $xmlConfig = ($xmlConfig.config.buttons.button)
    }
    $cmd = $($xmlConfig.cmd)
    $confirm = $($xmlConfig.confirm)
    $name = $($xmlConfig.'#text')

    if($confirm -eq 1)
    {
        $confirm = $null
        $confirm = [System.Windows.Forms.MessageBox]::Show("Run $($cmd) ?", "$($name)" , "YesNo","Warning")
    }
    else { $confirm = "Yes" }

    if($confirm -eq "Yes")
    {
        try
        {
            $message = "button $name : Invoke-Command -ScriptBlock { Invoke-Expression $($cmd) }" 
            write_log "WIN: $message"  

            $return = Invoke-Command -ScriptBlock { Invoke-Expression $($cmd) }
            
            if($return)
            {
                [System.Windows.Forms.MessageBox]::Show($return, $($name) , "Ok","Information")
                $message = "button result : $return" 

            }
            else
            {
                [System.Windows.Forms.MessageBox]::Show("Completed", $($name) , "Ok","Information")
                $message = "button result : completed"
            }
            write_log "WIN: $message"
        }
        catch
        {
            write-host $_.Exception.Message
            $message = "[ERROR] $($_.Exception.Message)"
            write_log $message
            [System.Windows.Forms.MessageBox]::Show("Something went wrong", $($name) , "Ok","Error")
        }
    }    
}
###########################################


switch($cmd)
{
    "getservers"   {$results = workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd}
    "getservices"  {$results = workflow_common -data (read_services $xmlConfig $server $name) -cmd $cmd}
    "getprocesses" {$results = workflow_common -data (read_processes $xmlConfig $server $name) -cmd $cmd}
    "gettasks"     {$results = workflow_common -data (read_tasks $xmlConfig $server $name) -cmd $cmd}
    "getfw"        {$results = workflow_common -data (read_firewall $xmlConfig $server $name) -cmd $cmd}
    "getps"        {$results = workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd}    
    "setfw"        {$results = workflow_common -data (read_firewall $xmlConfig $server $name) -cmd $cmd; $results = $null}
    "stopservices" {
                        #$results = workflow_common -data (read_services $xmlConfig $server $name) -cmd $cmd; $results = $null
                        $data = read_services $xmlConfig $server $name
                        $data = $data | Sort-Object @{Expression={$_.Index}; Ascending=$false}
                        
                        foreach($record in $data)
                        {
                            $results = workflow_common -data $record -cmd $cmd
                            $results = $null
                            if($record.attr2 -notin ("true", "false") -and $record.attr2)
                            {
                                Start-Sleep -s $record.attr2
                            }
                        }
                   }
    "startservices"{
                        #$results = workflow_common -data (read_services $xmlConfig $server $name) -cmd $cmd; $results = $null
                        $data = read_services $xmlConfig $server $name

                        foreach($record in $data)
                        {                           
                            if($record.attr1 -eq "true")
                            {
                                $results = workflow_common -data $record -cmd $cmd
                                $results = $null
                            }
                            else
                            {
                                #write_log "Invoke-Command -ComputerName $($record.Server)  -ErrorAction Stop -ArgumentList $($record.Name) -ScriptBlock {param($service) Start-Service -Name $service } -AsJob "
                                $results = Invoke-Command -ComputerName $record.Server  -ErrorAction Stop -ArgumentList $record.Name -ScriptBlock {param($service) Start-Service -Name $service } -AsJob 
                                $results = $null
                                if($record.attr1 -ne "false" -and $record.attr1)
                                {
                                    Start-Sleep -s $record.attr1
                                }
                            }
                        }
                   }
    "starttasks"   {$results = workflow_common -data (read_tasks $xmlConfig $server $name) -cmd $cmd; $results = $null}
    "stoptasks"    {$results = workflow_common -data (read_tasks $xmlConfig $server $name) -cmd $cmd; $results = $null}
    "reboot"       {$results = workflow_common -data (read_servers $xmlConfig $server $name) -cmd $cmd; $results = $null}
    "getwu"        {$results = workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd}
    "startwu"      {$results = workflow_common -data (read_servers $xmlConfig $server) -cmd $cmd}
    "getfolders"   {$results = workflow_folders -data (read_folders $xmlConfig $name)}
    "help"
    {
        if($detailed)
        {
            write-output "`tgetservers - get list of servers. Use getservers [server] for single server"
            write-output "`tgetservices - get list of services. Use getservices [server] [service] for single service"
            write-output "`tgetprocesses - get list of processes. Use getprocesses [server] [process] for single process"
            write-output "`tgettasks - get list of tasks. Use gettasks [server] [task] for single task"
            write-output "`tgetfw - get list of firewall rules. Use getfw [server] [rule] for single rule"
            write-output "`tgetps - get ps version of servers. Use getps [server] for single output"
            write-output "`tsetfw - on or off firewall rules. Use setfw [server] [rule] for single rule"
            write-output "`tstopservices - stop all services. Use stopservices [server] [service] for single service"
            write-output "`tstartservices - start all services. Use startservices [server] [service] for single service"
            write-output "`tstarttasks - start all tasks. Use starttasks [server] [task] for single task"
            write-output "`tstoptasks - stop all tasks. Use stoptasks [server] [task] for single task"
            write-output "`treboot - reboot all servers. Use reboot [server] for single server"
            write-output "`tgetfolders - get size and files count in folders. Use getfolders [name] for single output"
        }
        else
        {
            write-output "`tgetservers`n`tgetservices`n`tgetprocesses`n`ttgettasks`n`tgetfw`n`tgetps`n`tsetfw`n`tstopservices`n`tstartservices`n`tstarttasks`n`tstoptasks`n`treboot`n`tgetfolders"
        }
    }
}

###########################################
# OUTPUT
###########################################
if($cmd -eq "win")
{
    write_log "WIN: started"
    win_mode    
}
elseif($cmd -eq "getfolders")
{
    write_log "CMD: getfolders"

    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}, @{Expression={$_.Id}; Ascending=$True}
    foreach($result in $results)
    {
        if($result.Id -eq 0)
        {
            $main_index = $result.Index
            $main_size = $result.Size
            $main_count = $result.Quantity        
            $message = "*"
            $status = "Ok"
            
        }
        else
        {
            if( ($main_size -eq $($result.Size)) -and $main_count -eq $($result.Quantity)){ $message = "+" ; $status = "Ok" }
            else { $message = "-" ; $status = "Error" }
        }

        $email_message = "$email_message <tr align='center'><td> $message[$($result.Name)] </td><td> $($result.Server) </td><td> $($result.Desc) </td><td> S:$([math]::Round($($result.Size),2)) Q:$($result.Quantity) </td></tr>"
        $message = "$message[$($result.Name)] $($result.Server) $($result.Desc) : S:$([math]::Round($($result.Size),2)) Q:$($result.Quantity)"
        
        write_log $message
        write_out $message $status 
        
    }
 
    send_email $email_message "Folders"
}
elseif($cmd -eq "getprocesses")
{
    write_log "CMD: getprocesses"

    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}
    foreach($result in $results)
    {
        if( $(($($result.Status)).Split(" ")).Count -gt 1 )
        {
            for($i=0;$i -lt $(($($result.Status)).Split(" ")).Count;$i++)
            {
                $status = $(($($result.Status)).Split(" "))[$i]
                $message = "[$($result.Index)][$i] $($result.Server) : $($result.Name) - $status"
                $email_message = "$email_message <tr align='center'><td> $i </td><td> $($result.Server) </td><td>  $($result.Name) </td><td> $status </td></tr>"
        
                write_log $message
                if($detailed -eq $true)
                {
                    #$message = "$($result.Name) `t $status"
                    write_out $message $status
                }                
            }
            
            if($detailed -eq $false)
            {
                $message = "$($result.Name) [$($(($($result.Status)).Split(" ")).Count)] `t OK"
                write_out $message $status
            }            
        }
        else
        {
            $status = $(($($result.Status)).Split(" "))

            $message = "[$($result.Index)] $($result.Server) : $($result.Name) - $status"
            $email_message = "$email_message <tr align='center'><td></td><td> $($result.Server) </td><td>  $($result.Name) </td><td> $status </td></tr>"

            write_log $message
            if($detailed -eq $false)
            {
                if($status -eq "n/a")
                {
                    $message = "$($result.Name) `t $status"
                }
                else
                {
                    $message = "$($result.Name) `t OK"
                }
            }
            write_out $message $status
        }
    }
    send_email $email_message "Processes"
}
elseif($cmd -eq "getwu")
{
    write_log "CMD: getwu"

    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}
    if($results)
    {
        foreach($result in $results)
        {
            #$result
            #$message = "[$($result.Index)] $($result.Server) : $($result.Name) - $($result.Desc) $($result.Status) $($result.Attr1) $($result.Attr2)"
            $message = "[$($result.Index)] $($result.Server) : $($result.Name) - $($result.Desc) `t $($result.Status)"
            $email_message = "$email_message <tr align='center'><td> $($result.Index) </td><td> $($result.Server) </td><td>  $($result.Desc) </td><td> $($result.Status) </td></tr>"

            $status = $($result.Status)
            write_log $message
            if($detailed -eq $false)
            {
                $message = "$($result.Server) : $($result.Desc) - $($result.Status)"
            }
            write_out $message $status
        }

        send_email $email_message ($cmd.Substring(3, $cmd.Length -3 )).ToUpper()
    }
}
else
{
    write_log "CMD: $cmd"

    $results = $results | Sort-Object @{Expression={$_.Index}; Ascending=$True}
    if($results)
    {
        foreach($result in $results)
        {
            #$message = "[$($result.Index)] $($result.Server) : $($result.Name) - $($result.Desc) $($result.Status) $($result.Attr1) $($result.Attr2)"
            $message = "[$($result.Index)] $($result.Server) : $($result.Name) - $($result.Desc) `t $($result.Status)"
            $email_message = "$email_message <tr align='center'><td> $($result.Index) </td><td> $($result.Server) </td><td>  $($result.Desc) </td><td> $($result.Status) </td></tr>"

            $status = $($result.Status)
            write_log $message
            if($detailed -eq $false)
            {
                $message = "$($result.Desc) `t $($result.Status)"
            }
            write_out $message $status
        }

        send_email $email_message ($cmd.Substring(3, $cmd.Length -3 )).ToUpper()
    }
}
###########################################