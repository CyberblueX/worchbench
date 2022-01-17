<#
.SYNOPSIS
    Konvertiert alte "Binäre Office Dokument"-Formate in aktuelle Datei-Formate um.
.DESCRIPTION
    Konvertiert alte "Binäre Office Dokument"-Formate in aktuelle Datei-Formate um.
.EXAMPLE

    PS C:\> .\Convert-OfficeBinaryDocuments.ps1
    Dokumente im Ordner in welchem das Script aktuell gespeichert ist werden konvertiert.

    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -Recurse
    Dokumente im Ordner in welchem das Script aktuell gespeichert ist sowie dessen Unterordner werden konvertiert.

    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -Path "C:\ToConvert"
    Dokumente im Ordner "C:\ToConvert" werden konvertiert.

    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -TargetPath "C:\Converted"
    Die konvertierten Dokumente werden im Ordner "C:\Converted" gespeichert.

    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -Path "C:\ToConvert" -BackupPath "C:\OriginalBackups"
    Dokumente im Ordner "C:\ToConvert" werden konvertiert. Zudem werden die Original-Dokumente in den Ordner "C:\OriginalBackups" verschoben.

    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -CreatePDF
    Erzeugt zusätzlich einen PDF Ausdruck der Quelldatei.
    
    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -KeepFileTime
    Übernimmt die Zeitangaben für "Zuletzt bearbeitet" auf die konvertieren Dokumente. Dies gilt nicht für PDFs.
    
    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -Force
    Überschreibt ohne Rückfrage Dateien am Zielort, wenn diese bereits vorhanden sind.

    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -Force
    Zeigt eine Rückfrage wenn Dateien am Zielort bereits vorhanden sind.
    
    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -ShowFileList
    Vor Beginn wird zur Information eine Auflistung alle gefundenen Dokumente angezeigt.
    
    PS C:\> .\Convert-OfficeBinaryDocuments.ps1 -ShowFileSelection
    Vor Beginn wird eine Auswahl aller gefunden Dokumente angezeigt. Über eine Mehrfachauswahl (STRG / Umschalt) können einzelne Dokumente ausgewählt werden.

.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    Office Binary Document Converter
    
    Funktionen
    --------------
    Allgemein:
    + Auswahl einzelner Dateien
    + PDFs erzeugen
    + Original Dateidatum behalten
    + Original Dateien in ausgewählten BackupOrdner verschieben

    - Dokumente mit Passwörtern werden nicht unterstützt
    - Wenn das Zieldokument bereits vorhanden ist wird die konvertierung übersprungen. Es wird auch keine PDF erzeugt, falls angefordert.
    - PDFs werden immer überschrieben.

    Excel:
    +
    
    Word:
    + Bei PDFs werden Lesezeichen entsprechend der Überschriften erzeugt

    PowerPoint:
    + Es werden 2 PDFs erzeugt. Eine mit Vollbild-Slides und eine _Handout mit der Ansicht 3 Slides & Notizbereich

    Sources:
    https://administrator.de/contentid/365694#comment-1266357
    https://docs.microsoft.com/de-de/office/vba/api/excel.xlfileformat
    https://social.msdn.microsoft.com/Forums/en-US/abd9b628-4ba2-4f0b-aab7-e5caf1602a83/powershell-store-entire-workbook-as-a-pdf-file?forum=exceldev
    https://www.experts-exchange.com/articles/7237/Mass-remove-known-password-from-Word-files.html
    https://stackoverflow.com/questions/60342174/word-exportasfixedformat
    https://gist.github.com/allenyllee/5d7c4a16ae0e33375e4a6d25acaeeda2

#>

#region Depencies and Inputs
#Requires -Version 3.0

[CmdletBinding()]
param (
    # Specifies a path to one or more locations.
    [Parameter(Mandatory=$false,
               Position=0,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true,
               HelpMessage="Pfad zum Ordner in welchem die Quelldateien sind. Standard: Aktueller Ordner des Scripts.")]
    [ValidateNotNullOrEmpty()]
    [Alias("PSPath")]
    [string[]]
    $Path = $PSScriptRoot,
    #$Path = $env:USERPROFILE,

    [Parameter(Mandatory=$false,
               Position=1,
               ValueFromPipelineByPropertyName=$true,
               HelpMessage="Ausgabe der neuen Dokumente in separaten Pfad umleiten.")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $TargetPath,

    [Parameter(HelpMessage="Wenn gesetzt werden alle Original-Dokumente in den BackupPfad verschoben.")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $BackupPath,

    [Parameter(HelpMessage="Sollen untergeordnete Verzeichnisse auch bearbeitet werden? Standard: false")]
    [switch]
    $Recurse,

    [Parameter(HelpMessage="Sollen zusätzlich PDF Dateien erstellt werden? Standard: false")]
    [switch]
    $CreatePDF,

    [Parameter(HelpMessage="Sollen die neu erzeugten Dokumente das ursprüngliche Änderungsdatum behalten? (Gilt nicht für PDF) Standard: false")]
    [switch]
    $KeepFileTime,

    [Parameter(HelpMessage="Soll vor begin eine Auflistung aller gefundenen Dokumente angezeigt werden? Standard: false")]
    [switch]
    $ShowFileList,

    [Parameter(HelpMessage="Sollen nur bestimmte gefundene Dokumente konvertiert werden? Standard: false")]
    [switch]
    $ShowFileSelection,

    [Parameter(HelpMessage="Sollen bereits vorhandene Dateien überschrieben werden? Standard: false")]
    [switch]
    $Force,

    [Parameter(HelpMessage="Soll eine Abfrage vor dem Überschreiben einzelner Dokumente erfolgen? Standard: false")]
    [switch]
    $Confirm

)

# add switches
if ($Confirm) {
    $Force = $true
}

# checks for wrong switch kombinations
if ($TargetPath -and $BackupPath) {
    Write-Host "Fehler: Verwenden Sie -TargetPath und -BackupPath nicht gemeinsam." -ForegroundColor Yellow
    Write-Host "" -ForegroundColor Yellow
    Write-Host "ENTWEDER:  die konvertierten Dokumente werden in eine neue Verzeichnisstrucktur gespiegelt      (-TargetPath)" -ForegroundColor Yellow
    Write-Host "ODER:      die Originaldokumente werden nach Konvertierung in ein Backupverzeichnis verschoben  (-BackupPath)" -ForegroundColor Yellow
    return
}

if ($ShowFileList -and $ShowFileSelection) {
    Write-Host "Fehler: Verwenden Sie -ShowFileList und -ShowFileSelection nicht gemeinsam." -ForegroundColor Yellow
    Write-Host "" -ForegroundColor Yellow
    Write-Host "ENTWEDER:  eine Gesamt-Auflistung zur Information wird angezeigt                               (-ShowFileList)" -ForegroundColor Yellow
    Write-Host "ODER:      eine Gesamt-Auflistung bei welcher Dateien AUSGEWÄHLT werden MÜSSEN wird angezeigt  (-ShowFileSelection)" -ForegroundColor Yellow
    return
}

# Check Input Vars
if (!(Test-Path $Path)) {
    Write-Host "Fehler: Das angegebene Quellverzeichnis '$Path' ist nicht vorhanden." -ForegroundColor Yellow
    return
} elseif ((Get-Item -Path $Path) -isnot [System.IO.DirectoryInfo]) {
    Write-Host "Fehler: Das angegebene Quellverzeichnis '$Path' ist kein Ordner." -ForegroundColor Yellow
    return
} else {
    [System.IO.DirectoryInfo]$Path = Get-Item -Path $Path
}

# create targetpath if needed
if ($TargetPath) {
    if (!(Test-Path -Path $TargetPath)) {
        [System.IO.DirectoryInfo]$TargetPath = New-Item -ItemType "Directory" -Path $TargetPath -Force
    } elseif ((Get-Item -Path $TargetPath) -isnot [System.IO.DirectoryInfo]) {
        Write-Host "Fehler: Das angegebene Zielverzeichnis '$TargetPath' ist kein Ordner." -ForegroundColor Yellow
        return
    } else {
        [System.IO.DirectoryInfo]$TargetPath = Get-Item -Path $TargetPath
    }    
}

# create backuppath if needed
if ($BackupPath) {
    if (!(Test-Path -Path $BackupPath)) {
        [System.IO.DirectoryInfo]$BackupPath = New-Item -ItemType "Directory" -Path $BackupPath -Force
    } elseif ((Get-Item -Path $BackupPath) -isnot [System.IO.DirectoryInfo]) {
        Write-Host "Fehler: Das angegebene Backupverzeichnis '$BackupPath' ist kein Ordner." -ForegroundColor Yellow
        return
    } else {
        [System.IO.DirectoryInfo]$BackupPath = Get-Item -Path $BackupPath
    }
}
#endregion

#region Defaults & Vars

# Filter for file extensions
$extensions_word        = ".doc", ".dot"
$extensions_excel       = ".xls", ".xlt"
$extensions_powerpoint  = ".ppt", ".pot", ".pps"
$extensions             = $extensions_word + $extensions_excel + $extensions_powerpoint

# Filter for FileList
$size = @{label="Size(MB)";expression={[math]::Round($_.length/1MB,2)}}
$directory = @{label="Relative Paths";expression={"." + $_.Directory.tostring().replace($Path.Fullname, '')}}
$filelist_filter = "CreationTime", "LastWriteTime", $directory, "Name", "Extension", $size, "Fullname"

# Specials for Apps
$mpar = [System.Reflection.Missing]::Value
$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type] 

#endregion

#region Get Files

# Read all files
$files_read = Get-ChildItem -Path $Path -Recurse:$Recurse

# filter by extensions
# Note: using -Include was slower than filter later
$files_filtered = @($files_read | Where-Object {$_.Extension -in $extensions}) | Sort-Object -Property Directory,Name

# check
if (!$files_filtered) {
    Write-Host "Fehler: Keine kompatiblen Dokumente im gewähltlen Verzeichnis: '$($Path.Fullname)'." -ForegroundColor Yellow
    return
}
#endregion

#region Show FileLists or FileSelection
if ($ShowFileList) {
    $continue = $files_filtered | Select-Object -Property $filelist_filter | Out-GridView -Title "Gefundene Dokumente, fortfahren?" -PassThru
    if (!$continue) {
        Write-Host "Abbruch durch Benutzer." -ForegroundColor Yellow
        return            
    }
}

if ($ShowFileSelection) {
    $files_preselected = $files_filtered | Select-Object -Property $filelist_filter | Out-GridView -PassThru -Title "Bitte wählen Sie (Mehrfachauswahl mit STRG/UMSCHALT) die gewünschten Dokumente aus:"
    if (!$files_preselected) {
        Write-Host "Fehler: Keine Dokumente ausgewählt." -ForegroundColor Yellow
        return
    } else {
        $files_selected = @()
        foreach ($selection in $files_preselected) {
            $files_selected += $files_filtered | Where-Object {$_.Fullname -eq $selection.Fullname}
        }
    }
} else {
    $files_selected = $files_filtered
}

if (!$files_selected) {
    Write-Host "Fehler: Keine kompatiblen Dokumente im gewähtlen Verzeichnis: '$($Path.Fullname)' ." -ForegroundColor Yellow
    return
}

$files_word =           @($files_selected | Where-Object {$_.Extension -in $extensions_word})
$files_excel =          @($files_selected | Where-Object {$_.Extension -in $extensions_excel})
$files_powerpoint =     @($files_selected | Where-Object {$_.Extension -in $extensions_powerpoint})
#$files_all = $files_word + $files_excel + $files_powerpoint

#endregion

#region Create temp folder
$tmpWorkPath = Join-Path -Path $env:TEMP -ChildPath "Convert-OfficeBinaryDocuments"

if (Test-Path -Path $tmpWorkPath) {
    Remove-Item -Path $tmpWorkPath -Force -Recurse   
}
[System.IO.DirectoryInfo]$WorkPath = New-Item -ItemType "Directory" -Path $tmpWorkPath -Force
#endregion

#region Converting Excel Documents
if ($files_excel){
    Write-Host "Bearbeite Excel Dokumente:" -ForegroundColor Magenta
    
    # Prepare Excel for conversion
    $xlsApp = New-Object -ComObject Excel.Application
    $xlsApp.DisplayAlerts = $false
    $xlsApp.ScreenUpdating = $false
    $xlsApp.Visible = $false
    
    # For each file
    foreach ($file in $files_excel) {
        Write-Host "Konvertiere " -NoNewline
        Write-Host "$($file.Fullname -replace "^$([regex]::escape($Path.Fullname))")" -NoNewline -ForegroundColor Cyan
        Write-Host " ... " -NoNewline
        
        try{
            # Open Workbook
            $xlsWorkbooks = $xlsApp.Workbooks.Open($file.Fullname,$false,$true)

            # Check for Macros
            $HasVBProject = $xlsWorkbooks.HasVBProject
            
            # Set Default Output format
            $TargetExtension = '.xlsx'
            $TargetFormatID = 51
            
            # determine file extension and set output format
            switch($file.Extension){
                '.xls' {
                    $TargetExtension    = @{$true='.xlsm';$false='.xlsx'}[$HasVBProject]
                    $TargetFormatID     = @{$true=52;$false=51}[$HasVBProject]
                }
                '.xlt' {
                    $TargetExtension    = @{$true='.xltm';$false='.xltx'}[$HasVBProject]
                    $TargetFormatID     = @{$true=53;$false=54}[$HasVBProject]
                }
            }

            if ($TargetPath) {
                # generate and create TargetFolder for output file
                $TargetFolder = $file.Directory.Fullname -replace "^$([regex]::escape($Path.Fullname))",$TargetPath.Fullname
                if (!(Test-Path $TargetFolder)) {
                    $objTargetFolder = New-Item -ItemType Directory $TargetFolder -Force 
                } else {
                    $objTargetFolder = Get-Item $TargetFolder
                }
                # new filename in targetdirecory
                [string]$TargetFile = Join-Path -Path $objTargetFolder.FullName -ChildPath ($file.Basename + $TargetExtension)
                [string]$TargetFilePDF = Join-Path -Path $objTargetFolder.FullName -ChildPath ($file.Basename + ".pdf")
            } else {
                [string]$TargetFile = Join-Path -Path $file.Directory.FullName -ChildPath ($file.Basename + $TargetExtension)
                [string]$TargetFilePDF = Join-Path -Path $file.Directory.FullName -ChildPath ($file.Basename + ".pdf")            
            }
            [string]$TargetFileTmp = Join-Path -Path $WorkPath.FullName -ChildPath ($file.Basename + $TargetExtension)
            
            # Check if target already exists and if Force is set
            if (!($Force -or $Confirm) -and (Test-Path $TargetFile)) {
                Write-Host "Skipped. " -ForegroundColor Yellow -NoNewline
                Write-Host "(Zieldokument existiert bereits. Verwende -Force zum überschreiben oder -Confirm für eine einzelne Abfrage.)"
                continue
            }

            # Save file as new format
            $xlsWorkbooks.SaveAs($TargetFileTmp,[ref]$TargetFormatID)
            
            # Create PDF if required
            if ($CreatePDF) {
            #     $xlsWorkbooks.ActiveSheet.PageSetup.Orientation = 2
            #     $xlsApp.PrintCommunication = $false
            #     $xlsWorkbooks.ActiveSheet.PageSetup.FitToPagesTall = $false
            #     $xlsWorkbooks.ActiveSheet.PageSetup.FitToPagesWide = 1
            #     $xlsApp.PrintCommunication = $true
            #     $xlsWorkbooks.Saved = $true 
            #    "saving $filepath" 
                $xlsWorkbooks.ExportAsFixedFormat([ref]$xlFixedFormat::xlTypePDF,[ref]$TargetFilePDF)
            }

            # close document
            if ($xlsWorkbooks){
                $xlsWorkbooks.Close($false)
                $xlsWorkbooks = $null
            }

            if ($KeepFileTime) {
                $objTargetFile = Get-Item $TargetFileTmp
                $objTargetFile.LastAccessTimeUtc = $file.LastAccessTimeUtc
                $objTargetFile.LastWriteTimeUtc = $file.LastWriteTimeUtc
            }
            
            # Move file to target
            Move-Item -Path $TargetFileTmp -Destination $TargetFile -Force:$Force -Confirm:$Confirm

            if ($BackupPath) {
                # create BackupFolder
                $BackupFolder = $file.Directory.Fullname -replace "^$([regex]::escape($Path.Fullname))",$BackupPath.Fullname
                if (!(Test-Path $BackupFolder)) {
                    $objBackupFolder = New-Item -ItemType Directory $BackupFolder -Force 
                } else {
                    $objBackupFolder = Get-Item $BackupFolder
                }
                [string]$BackupFile = Join-Path -Path $objBackupFolder.FullName -ChildPath $file.Name
                Move-Item -Path $file.Fullname -Destination $BackupFile -Force:$Force -Confirm:$Confirm
            } 

            # Check if new file exists
            if (Test-Path $TargetFile){Write-Host 'OK.' -ForegroundColor Green}

        } catch {
            # Error occured
            Write-Host "Fehler!" -ForegroundColor Red
            Write-Error "$($_.Exception.Message)"
        } finally {
            # close workbook
            if ($xlsWorkbooks){
                $xlsWorkbooks.Close($false)
                $xlsWorkbooks = $null
            }
        }
    }
    # Quit Excel and tidy up
    $xlsApp.DisplayAlerts = $true
    $xlsApp.Screenupdating = $true
    $xlsApp.Quit() | out-null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlsApp) | out-null
}
#endregion

#region Converting Word Documents
if ($files_word) {
    Write-Host "Bearbeite Word Dokumente:" -ForegroundColor Magenta

    # Prepare Word for conversion
    $wordApp = New-Object -Com Word.Application
    $wordApp.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]0
    $wordApp.Screenupdating = $false
    $wordApp.Visible = $false
    
    # For each file
    foreach ($file in $files_word) {
        
        Write-Host "Konvertiere " -NoNewline
        Write-Host "$($file.Fullname -replace "^$([regex]::escape($Path.Fullname))")" -NoNewline -ForegroundColor Cyan
        Write-Host " ... " -NoNewline

        try{
            # Open Document
            $document = $wordApp.Documents.Open($file.Fullname,$false,$true)
            
            # Check for Macros
            $HasVBProject = $document.HasVBProject
            
            # Set Default Output format
            $TargetExtension = '.docx'
            $TargetFormatID = 12
            
            # determine file extension and set output format
            switch($file.Extension){
                '.doc' {
                    $TargetExtension =  @{$true='.docm';$false='.docx'}[$HasVBProject]
                    $TargetFormatID =   @{$true=13;$false=12}[$HasVBProject]
                }
                '.dot' {
                    $TargetExtension =  @{$true='.dotm';$false='.dotx'}[$HasVBProject]
                    $TargetFormatID =   @{$true=15;$false=14}[$HasVBProject]
                }
            }

            if ($TargetPath) {
                # generate and create TargetFolder for output file
                $TargetFolder = $file.Directory.Fullname -replace "^$([regex]::escape($Path.Fullname))",$TargetPath.Fullname
                if (!(Test-Path $TargetFolder)) {
                    $objTargetFolder = New-Item -ItemType Directory $TargetFolder -Force 
                } else {
                    $objTargetFolder = Get-Item $TargetFolder
                }
                # new filename in targetdirecory
                [string]$TargetFile = Join-Path -Path $objTargetFolder.FullName -ChildPath ($file.Basename + $TargetExtension)
                [string]$TargetFilePDF = Join-Path -Path $objTargetFolder.FullName -ChildPath ($file.Basename + ".pdf")
            } else {
                [string]$TargetFile = Join-Path -Path $file.Directory.FullName -ChildPath ($file.Basename + $TargetExtension)
                [string]$TargetFilePDF = Join-Path -Path $file.Directory.FullName -ChildPath ($file.Basename + ".pdf")            
            }
            [string]$TargetFileTmp = Join-Path -Path $WorkPath.FullName -ChildPath ($file.Basename + $TargetExtension)

            # Convert to new format and enable all features
            $document.Convert()

            # Check if target already exists and if Force is set
            if (!($Force -or $Confirm) -and (Test-Path $TargetFile)) {
                Write-Host "Skipped. " -ForegroundColor Yellow -NoNewline
                Write-Host "(Zieldokument existiert bereits. Verwende -Force zum überschreiben oder -Confirm für eine einzelne Abfrage.)"
                continue
            }
            
            # Save file as new format
            $document.SaveAs2([ref][system.object]$TargetFileTmp,[ref]$TargetFormatID)
            
            # Create PDF if required
            if ($CreatePDF) {
                #$document.SaveAs([ref][system.object]$TargetFilePDF, [ref]17)
                $document.Activate()
                $document.ExportAsFixedFormat2($TargetFilePDF, [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF, $false, [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint, [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument, 0, 0, [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent, $true, $true, [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateHeadingBookmarks)
            }

            # close document
            if ($document){
                $document.Close($false)
                $document = $null
            }

            if ($KeepFileTime) {
                $objTargetFile = Get-Item $TargetFileTmp
                $objTargetFile.LastAccessTimeUtc = $file.LastAccessTimeUtc
                $objTargetFile.LastWriteTimeUtc = $file.LastWriteTimeUtc
            }

            # Move file to target
            Move-Item -Path $TargetFileTmp -Destination $TargetFile -Force:$Force -Confirm:$Confirm

            if ($BackupPath) {
                # create BackupFolder
                $BackupFolder = $file.Directory.Fullname -replace "^$([regex]::escape($Path.Fullname))",$BackupPath.Fullname
                if (!(Test-Path $BackupFolder)) {
                    $objBackupFolder = New-Item -ItemType Directory $BackupFolder -Force 
                } else {
                    $objBackupFolder = Get-Item $BackupFolder
                }
                [string]$BackupFile = Join-Path -Path $objBackupFolder.FullName -ChildPath $file.Name
                Move-Item -Path $file.Fullname -Destination $BackupFile -Force:$Force -Confirm:$Confirm
            } 

            # Check if new file exists
            if (Test-Path $TargetFile){Write-Host 'OK.' -ForegroundColor Green}

        } catch {
            # Error occured
            Write-Host "Fehler!" -ForegroundColor Red
            Write-Error "$($_.Exception.Message)"
        } finally {
            # close document
            if ($document){
                $document.Close($false)
                $document = $null
            }
        }
    }
    # Quit Word and tidy up
    $wordApp.DisplayAlerts = -1
    $wordApp.Screenupdating = $true
    $wordApp.Quit() | out-null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | out-null
}

#endregion

#region Converting PowerPoint Documents
if ($files_powerpoint){
    Write-Host "Bearbeite PowerPoint Dokumente:" -ForegroundColor Magenta
    
    # Prepare PowerPoint for conversion
    [Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Powerpoint") > $null
    [Reflection.Assembly]::LoadWithPartialname("Office") > $null # need this or powerpoint might not close
    $ppApp = New-Object "Microsoft.Office.Interop.Powerpoint.ApplicationClass"
    #$ppApp = New-Object -Com Powerpoint.Application
    $ppApp.DisplayAlerts = 1
    
    Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

    # For each file
    foreach($file in $files_powerpoint){

        Write-Host "Konvertiere " -NoNewline
        Write-Host "$($file.Fullname -replace "^$([regex]::escape($Path.Fullname))")" -NoNewline -ForegroundColor Cyan
        Write-Host " ... " -NoNewline

        try{
            # Open Document
            $presentation = $ppApp.Presentations.Open($file.Fullname,$true,$mpar,0)
            
            # Check for Macros
            $HasVBProject = $presentation.HasVBProject

            # Set Default Output format
            $TargetExtension = '.pptx'
            $TargetFormatID = 12
            
            # determine file extension and set format number
            switch($file.Extension){
                '.ppt' {
                    $TargetExtension =  @{$true='.pptm';$false='.pptx'}[$HasVBProject]
                    $TargetFormatID =   @{$true=25;$false=24}[$HasVBProject]
                }
                '.pot' {
                    $TargetExtension =  @{$true='.potm';$false='.potx'}[$HasVBProject]
                    $TargetFormatID =   @{$true=27;$false=26}[$HasVBProject]
                }
                '.pps' {
                    $TargetExtension =  @{$true='.ppsm';$false='.ppsx'}[$HasVBProject]
                    $TargetFormatID =   @{$true=29;$false=28}[$HasVBProject]
                }
            }

            if ($TargetPath) {
                # generate and create TargetFolder for output file
                $TargetFolder = $file.Directory.Fullname -replace "^$([regex]::escape($Path.Fullname))",$TargetPath.Fullname
                if (!(Test-Path $TargetFolder)) {
                    $objTargetFolder = New-Item -ItemType Directory $TargetFolder -Force 
                } else {
                    $objTargetFolder = Get-Item $TargetFolder
                }
                # new filename in targetdirecory
                [string]$TargetFile = Join-Path -Path $objTargetFolder.FullName -ChildPath ($file.Basename + $TargetExtension)
                [string]$TargetFilePDF = Join-Path -Path $objTargetFolder.FullName -ChildPath ($file.Basename + ".pdf")
                [string]$TargetFilePDF2 = Join-Path -Path $objTargetFolder.FullName -ChildPath ($file.Basename + "_HandOut" + ".pdf")
            } else {
                [string]$TargetFile = Join-Path -Path $file.Directory.FullName -ChildPath ($file.Basename + $TargetExtension)
                [string]$TargetFilePDF = Join-Path -Path $file.Directory.FullName -ChildPath ($file.Basename + ".pdf")
                [string]$TargetFilePDF2 = Join-Path -Path $file.Directory.FullName -ChildPath ($file.Basename + "_HandOut" + ".pdf")              
            }
            [string]$TargetFileTmp = Join-Path -Path $WorkPath.FullName -ChildPath ($file.Basename + $TargetExtension)

            # Check if target already exists and if Force is set
            if (!($Force -or $Confirm) -and (Test-Path $TargetFile)) {
                Write-Host "Skipped. " -ForegroundColor Yellow -NoNewline
                Write-Host "(Zieldokument existiert bereits. Verwende -Force zum überschreiben oder -Confirm für eine einzelne Abfrage.)"
                continue
            }
            
            # Save file as new format
            $presentation.SaveAs([ref]$TargetFileTmp,[ref]$TargetFormatID)

            # Create PDF if required
            if ($CreatePDF) {
                #$presentation.ExportAsFixedFormat2($TargetFilePDF, [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF, [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentPrint, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder]::ppPrintHandoutHorizontalFirst, ([Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputSlides,[Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputFourSlideHandouts))
                
                $fixedFormatType = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF
                $intent = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentPrint
                $frameSlides = [Microsoft.Office.Core.MsoTriState]::msoTrue
                $handoutOrder = [Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder]::ppPrintHandoutVerticalFirst
                $outputType = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputSlides
                $outputType2 = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputThreeSlideHandouts
                $printHiddenSlides = [Microsoft.Office.Core.MsoTriState]::msoFalse
                $printRange = $presentation.PrintOptions.Ranges.Add(1, $presentation.Slides.Count)
                $rangeType = [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintAll
    
                # String	The name of the slide show.
                $slideShowName = "Slideshow Name"
                
                # Boolean	Whether the document properties should also be exported. The default is False.
                $includeDocProperties = $false
                
                # Boolean	Whether the IRM settings should also be exported. The default is True.
                $keepIRMSettings = $true
                
                # Boolean	Whether to include document structure tags to improve document accessibility. The default is True.
                $docStructureTags = $true
                
                # Boolean	Whether to include a bitmap of the text. The default is True.
                $bitmapMissingFonts = $true
                
                # Boolean	Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is False.
                $useISO19005_1 = $false
                
                # Boolean	Whether the resulting document should include associated pen marks.
                $includeMarkup = $true
                
                # Variant	A pointer to an Office add-in that implements the IMsoDocExporter COM interface and allows calls to an alternate implementation of code. The default is a null pointer.
                $externalExporter = $null
                
                ##
                ## Publishes as PDF or XPS.
                ##
                ## vba - difference between ExportAsFixedFormat2 and ExportAsFixedFormat? - Stack Overflow
                ## https://stackoverflow.com/questions/37585025/difference-between-exportasfixedformat2-and-exportasfixedformat
                ##
                
                # ExportAsFixedFormat2 can include pen markups
                # Presentation.ExportAsFixedFormat2 Method (PowerPoint)
                # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/presentation-exportasfixedformat2-method-powerpoint
                $presentation.ExportAsFixedFormat2($TargetFilePDF, $fixedFormatType, $intent, $frameSlides, $handoutOrder, $outputType, $printHiddenSlides, $printRange, $rangeType, $slideShowName, $includeDocProperties, $keepIRMSettings, $docStructureTags, $bitmapMissingFonts, $useISO19005_1, $includeMarkup)
                $presentation.ExportAsFixedFormat2($TargetFilePDF2, $fixedFormatType, $intent, $frameSlides, $handoutOrder, $outputType2, $printHiddenSlides, $printRange, $rangeType, $slideShowName, $includeDocProperties, $keepIRMSettings, $docStructureTags, $bitmapMissingFonts, $useISO19005_1, $includeMarkup)
            }

            # close document
            if ($presentation){
                $presentation.Close()
                $presentation = $null
            }

            if ($KeepFileTime) {
                $objTargetFile = Get-Item $TargetFileTmp
                $objTargetFile.LastAccessTimeUtc = $file.LastAccessTimeUtc
                $objTargetFile.LastWriteTimeUtc = $file.LastWriteTimeUtc
            }

            # Move file to target
            Move-Item -Path $TargetFileTmp -Destination $TargetFile -Force:$Force -Confirm:$Confirm

            if ($BackupPath) {
                # create BackupFolder
                $BackupFolder = $file.Directory.Fullname -replace "^$([regex]::escape($Path.Fullname))",$BackupPath.Fullname
                if (!(Test-Path $BackupFolder)) {
                    $objBackupFolder = New-Item -ItemType Directory $BackupFolder -Force 
                } else {
                    $objBackupFolder = Get-Item $BackupFolder
                }
                [string]$BackupFile = Join-Path -Path $objBackupFolder.FullName -ChildPath $file.Name
                Move-Item -Path $file.Fullname -Destination $BackupFile -Force:$Force -Confirm:$Confirm
            } 

            # Check if new file exists
            if (Test-Path $TargetFile){Write-Host 'OK.' -ForegroundColor Green}


        } catch {
            # Error occured
            Write-Host "Fehler!" -ForegroundColor Red
            Write-Error "$($_.Exception.Message)"
        } finally {
            # close presentation
            if ($presentation){
                $presentation.Close()
                $presentation = $null
            }
        }
    }
    # Quit Powerpoint and tidy up
    $ppApp.DisplayAlerts = 2
    $ppApp.Quit() | out-null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | out-null
    [System.GC]::Collect();
    [System.GC]::WaitForPendingFinalizers();
    [System.GC]::Collect();
    [System.GC]::WaitForPendingFinalizers();
}
#endregion

#region CleanUp
if (Test-Path -Path $tmpWorkPath) {
    Remove-Item -Path $tmpWorkPath -Force -Recurse
}
#endregion


#region Tests
<#

Measure-Command {
    $files_read = Get-ChildItem -Path $path -Recurse
    
    $files_word =           $files_read | Where-Object {$_.Extension -in ".doc", ".dot"}
    $files_excel =          $files_read | Where-Object {$_.Extension -in ".xls", ".xlt"}
    $files_powerpoint =     $files_read | Where-Object {$_.Extension -in ".ppt", ".pot", ".pps"}
    
    } | ft -AutoSize

$wordApp_app = New-Object -ComObject Word.Application
#$TargetFormatID = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument

 | ForEach-Object {
    write-host $_.FullName
    $documentx_filename = "$($_.DirectoryName)\$($_.BaseName).docx"
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $documentument = $wordApp_app.Documents.Open($_.FullName)
    $documentument.SaveAs([ref]$documentx_filename, [ref]12)
    $documentument.SaveAs([ref]$pdf_filename, [ref]17)
    $documentument.Close()
}
$wordApp_app.Quit()
#>
#endregion