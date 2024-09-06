<#PSScriptInfo

.VERSION 1.0.0

.GUID aec62915-eace-4acb-b971-5de59ed8eb9f

.AUTHOR @cgeneske

.COPYRIGHT Copyright (c) 2024 Craig Geneske

.LICENSEURI https://github.com/cgeneske/CyberArkPSMClientImporter/blob/main/LICENSE.md

.PROJECTURI https://github.com/cgeneske/CyberArkPSMClientImporter

#>

<#
.SYNOPSIS
Interactive tool for importing Sysinternals Remote Desktop Connection Manager (RDCMan) RDG files, into 
PSMClient (Connection Manager) RDP CustomView tree.  Latest solution and full README are available at: 
https://github.com/cgeneske/CyberArkPSMClientImporter

.DESCRIPTION
With the addition of the RDP CustomView to the CyberArk PSMClient connection manager, it is now possible to create
an organized tree of server entries very similar to most popular RDP connection managers available today.
Windows administrators that were previously leveraging Sysinternals Remote Desktop Connection Manager (RDCMan) with
an established set of predefined RDG/Group/Server tree(s), can leverage this script to import these trees quickly
and easily into CybrArk PSMClient's RDP CustomView.

.EXAMPLE
CyberArk_PSMClient_Importer.ps1

.INPUTS
None - Script will prompt interactively

.OUTPUTS
None

.NOTES
AUTHOR:
Craig Geneske

VERSION HISTORY:
1.0.0   9/6/2024   - Initial Release

DISCLAIMER:
This solution is provided as-is - it is not supported by CyberArk nor an official CyberArk solution.
#>

using namespace System.Collections.Generic 

################################################## LOADING TYPES #################################################
#region Types
Add-Type -AssemblyName System.Windows.Forms

#endregion

#################################################### FUNCTIONS ###################################################
#region Functions

Function Write-Log {
    <#
    .SYNOPSIS
        Writes a consistently formatted log entry to stdout and a log file
    .DESCRIPTION
        This function is designed to provide a way to consistently format log entries and extend them to
        one or more desired outputs (i.e. stdout and/or a log file).  Each log entry consists of three main
        sections:  Date/Time, Event Type, and the Event Message.  This function is also extended to output
        a standard header during script invocation and footer at script conclusion.
    .PARAMETER Type
        Sets the type of event message to be output.  This must be a member of the defined ValidateSet:
        INF [Informational], WRN [Warning], ERR [Error].
    .PARAMETER Message
        The message to prepend to the log event
    .PARAMETER Header
        Prints the log header
    .PARAMETER Footer
        Prints the log footer
    .EXAMPLE
        [FUNCTION CALL]     : Write-Log -Type INF -Message "This is an informational log message"
        [FUNCTION RESULT]   : 02/09/2023 09:43:25 | [INF] | This is an informational log message
    .NOTES
        Author: Craig Geneske
    #>
    Param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('INF','WRN','ERR')]
        [string]$Type,

        [Parameter(Mandatory = $false)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [switch]$Header,

        [Parameter(Mandatory = $false)]
        [switch]$Footer
    )

    $eventColor = [System.Console]::ForegroundColor
    if ($Header) {
        if ([Environment]::UserInteractive) {
            $eventString = @"
#############################################################################################################
#                                                                                                           #
#                                    CyberArk PSMClient Importer                                            #
#                                                                                                           #
#############################################################################################################
"@
        }
        else {
            $eventString = ""
        }

        $eventString += "`n`n-----------------------> BEGINNING SCRIPT @ $(Get-Date -Format "MM/dd/yyyy HH:mm:ss") <-----------------------`n"
        $eventColor = "Cyan"
    }
    elseif ($Footer) {
        $eventString = "`n-------------------------> ENDING SCRIPT @ $(Get-Date -Format "MM/dd/yyyy HH:mm:ss") <-------------------------`n"
        $eventColor = "Cyan"
    }
    else {
        $eventString =  $(Get-Date -Format "MM/dd/yyyy HH:mm:ss") + " | [$Type] | " + $Message
        switch ($Type){
            "WRN" { $eventColor = "Yellow"; Break }
            "ERR" { $eventColor = "Red"; Break }
        }
    }

    #Console Output (Interactive)
    Write-Host $eventString -ForegroundColor $eventColor
}

Function Set-ServerNodes {
    <#
    .SYNOPSIS
        Walks the RDG XML tree and replicates server nodes to the PSMClient XML tree
    .DESCRIPTION
        Traverses the RDG XML tree recursively to replicate an identical nested group structure within PSMClient's
        XML tree.  Custom display names and target host names are all preserved.
    .PARAMETER StartNode
        The starting XML node in the RDG XML tree for considering group traversal or server node content for copy
    .PARAMETER TargetNode
        The target XML node in the PSMClient XML tree that should receive server node content from the RDG XML tree
    .EXAMPLE
        Set-ServerNodes -StartNode $rdgXmlDoc.RDCMan.file -TargetNode $psmClientRdgRoot
    .NOTES
        Author: Craig Geneske
    #>

    Param (
        [System.Xml.XmlNode]$StartNode,
        [System.Xml.XmlNode]$TargetNode
    )

    if ($StartNode.group) {
        foreach ($groupNode in $StartNode.group) {
            $groupEle = $psmXmlDoc.CreateElement("item")
            $groupEle.SetAttribute("name", $groupNode.properties.name)
            $groupEle.SetAttribute("text", $groupNode.properties.name)
            $groupEle.SetAttribute("imageindex", "0")
            $nextTargetNode = $TargetNode.AppendChild($groupEle)
            Set-ServerNodes -StartNode $groupNode -TargetNode $nextTargetNode
        }
    } else {
        foreach ($serverNode in $StartNode.server) {
            $displayName = $serverNode.properties.displayName
            if (!$displayName) {
                $displayName = $serverNode.properties.name
            }
            $serverEle = $psmXmlDoc.CreateElement("item")
            $serverEle.SetAttribute("name", $serverNode.properties.name)
            $serverEle.SetAttribute("text", $displayName)
            $serverEle.SetAttribute("imageindex", "2")
            $TargetNode.AppendChild($serverEle) | Out-Null
        }
    }    
}

#endregion

###################################################### MAIN ######################################################
#region Main

Clear-Host
Write-Log -Header

try {
    Write-Log -Type INF -Message "Please select the RDG File(s) for import..."
    [List[string]] $rdgFiles = @()
    $rdgOpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $rdgOpenFileDialog.Filter = "RDCMan Groups|*.rdg"
    $rdgOpenFileDialog.InitialDirectory = [System.Environment]::GetFolderPath("MyDocuments")
    $rdgOpenFileDialog.Multiselect = $true
    $rdgOpenFileDialog.Title = "Please select one or more RDG File(s) for import..."
    if ($rdgOpenFileDialog.ShowDialog() -eq "OK") {
        $rdgFiles.AddRange($rdgOpenFileDialog.FileNames)
        Write-Log -Type INF -Message "You selected the following files:"
        foreach ($file in $rdgFiles) {
            Write-Log -Type INF -Message "--> $file"
        }
    } else {
        Write-Log -Type ERR -Message "RDG file selection has been cancelled, aborting script"
        throw
    }

    Write-Log -Type INF -Message "Validating RDG File(s)..."
    [List[xml]]$rdgXmlTreeCandidates = @()
    foreach ($file in $rdgFiles) {
        $rdgXmlDoc = [xml]$(Get-Content -Path $file)
        if (!$rdgXmlDoc.RDCMan) {
            Write-Log -Type WRN -Message "The RDG file [$file] exists but is not the expected format, and will be ignored"
            continue
        }
        Write-Log -Type INF "RDG file [$file] is valid"
        $rdgXmlTreeCandidates.Add($rdgXmlDoc)
        $rdgXmlDoc = $null
    }

    if (!$rdgXmlTreeCandidates) {
        Write-Log -Type ERR -Message "No RDG files are valid, aborting script"
        throw
    }

    Write-Log -Type INF -Message "Please select your PSMConnectionManager.exe..."
    $exeOpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $exeOpenFileDialog.Filter = "PSM Connection Manager|PSMConnectionManager.exe"
    $exeOpenFileDialog.InitialDirectory = "C:\"
    $exeOpenFileDialog.Title = "Please select your PSMConnectionManager.exe..."
    if ($exeOpenFileDialog.ShowDialog() -eq "OK") {
        $PSMClientPath = Split-Path -Parent $exeOpenFileDialog.FileName
    } else {
        Write-Log -Type ERR -Message "PSM Connection Manager selection has been cancelled, aborting script"
        throw
    }

    $psmClientXmlPath = $PSMClientPath + "\Custom.xml"

    Write-Log -Type INF -Message "Validating PSMClient CustomView XML File..."
    if (Test-Path -Path $psmClientXmlPath -PathType Leaf) {
        $psmXmlDoc = [xml]$(Get-Content -Path $psmClientXmlPath)
        if (!$psmXmlDoc.CustomView) {
            Write-Log -Type ERR -Message "The PSMClient Custom.xml file was found but is not the expected format"
            throw
        }
        Write-Log -Type INF "PSMClient's CustomView (Custom.xml) is valid"
    } else {
        Write-Log -Type INF -Message "The PSMClient CustomView (Custom.xml) does not exist, one will be created"
        $psmXmlDoc = New-Object -TypeName System.Xml.XmlDocument
        $xmlDeclNode = $psmXmlDoc.CreateXmlDeclaration("1.0", "UTF-8", $null)
        $psmXmlDoc.AppendChild($xmlDeclNode) | Out-Null
        $custViewEle = $psmXmlDoc.CreateElement("CustomView")
        $custViewRoot = $psmXmlDoc.AppendChild($custViewEle)
        $rootEle = $psmXmlDoc.CreateElement("item")
        $rootEle.SetAttribute("name", "Root")
        $rootEle.SetAttribute("text", "Root")
        $rootEle.SetAttribute("imageindex", "1")
        $custViewRoot.AppendChild($rootEle) | Out-Null
    }

    [List[xml]]$rdgXmlTrees = @()
    foreach ($rdgXml in $rdgXmlTreeCandidates) {
        if ($psmXmlDoc.CustomView.item.HasChildNodes) {
            $shouldSkip = $false
            foreach ($node in $psmXmlDoc.CustomView.item.item) {
                if (($node.name -match $rdgXml.RDCMan.file.properties.name) -and ($node.imageindex -match "0")) {
                    Write-Log -Type WRN -Message "An existing node in the PSMClient CustomView tree has been identified by the same name as the RDG root in [$($rdgXml.RDCMan.file.properties.name)], RDG will be ignored"
                    $shouldSkip = $true
                    break
                }
            }
            if (!$shouldSkip) {
                $rdgXmlTrees.Add($rdgXml)
            }
        } else {
            $rdgXmlTrees.Add($rdgXml)
        }
    }

    if (!$rdgXmlTrees) {
        Write-Log -Type ERR -Message "No unique RDG file candidates remain, aborting script"
        throw
    }

    #Creating new root folder for RDG tree in PSMClient CustomView
    foreach ($rdgXmlTree in $rdgXmlTrees) {
        Write-Log -Type INF -Message "Creating new root folder for RDG tree [$($rdgXmlTree.RDCMan.file.properties.name)] in PSMClient CustomView..."
        $rdgFileEle = $psmXmlDoc.CreateElement("item")
        $rdgFileEle.SetAttribute("name", $rdgXmlTree.RDCMan.file.properties.name)
        $rdgFileEle.SetAttribute("text", $rdgXmlTree.RDCMan.file.properties.name)
        $rdgFileEle.SetAttribute("imageindex", "0")
        $psmClientRdgRoot = $psmXmlDoc.CustomView.item.AppendChild($rdgFileEle)
    
        #Walking RDG tree to populate PSMClient CustomView
        Write-Log -Type INF -Message "Walking RDG tree to populate PSMClient CustomView for RDG tree [$($rdgXmlTree.RDCMan.file.properties.name)]..."
        Set-ServerNodes -StartNode $rdgXmlTree.RDCMan.file -TargetNode $psmClientRdgRoot
    }

    #Backup and Save PSMClient CustomView
    Write-Log -Type INF -Message "Preparing to write out final PSMClient CustomView..."
    if (Test-Path -Path $psmClientXmlPath -PathType Leaf) {
        Write-Log -Type INF -Message "Custom.xml already exists, creating backup copy..."
        try {
            Copy-Item -Path $psmClientXmlPath -Destination $($psmClientXmlPath + ".bak__" + (Get-Date -Format "MM-dd-yyyy_HHmmss"))
        } catch {
            Write-Log -Type ERR -Message "Unable to create backup copy of Custom.xml, aborting"
            throw
        }
    }
    try {
        Write-Log -Type INF "Attempting to Save PSMClient CustomView..."
        $psmXmlDoc.Save($psmClientXmlPath)
    } catch {
        Write-Log -Type ERR -Message "Unable to save PSMClient CustomView --> $($_.Exception.Message)"
        throw
    }
    Write-Log -Type INF -Message "Script has completed successfully!"
} catch {
    if ($_.Exception.Message -notmatch "ScriptHalted") {
        Write-Log -Type ERR -Message "An unexpected error occured --> $($_.Exception.Message)"
    }
} finally {
    $error.Clear()
    Write-Log -Footer
    Write-Host -NoNewline -ForegroundColor Yellow "Press any key to exit..."
    $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') | Out-Null
    Clear-Host
    exit
}

#endregion