<#
.Synopsis
   Convert a file in avi (i.e. video format) to mp4
.Description
   As synopsis, plus the function depends on having the Handbrake CLI installed at the location $handBrakeFolder\HandBrakeCLI.exe
   The variables $presetSwitch, $presetValue, $verboseSwitch take values of meaning to Handbrake.
   Right now, these are hard-coded for simplicity in preference to flexibility.
.Example
   AviToMp4 "c:\temp\source.avi" "c:\temp\target.mp4" -h $handbrakeFolder
   AviToMp4 -sourceAvi $source -targetMp4 $target -h $handbrakeFolder
#>
function AviToMp4 ($sourceAvi, $targetMp4, $handbrakeFolder) {
    $cmd = "$handbrakeFolder\HandBrakeCLI.exe"
    $presetSwitch = "-Z"
    $presetValue = "Fast 1080p30"
    #$presetValue = "HQ 1080p30 Surround"
    $verboseSwitch = "--verbose=1"
    $sourceAvi
    & $cmd $presetSwitch $presetValue -i $sourceAvi -o $targetMp4 $verboseSwitch
}

<#
.Synopsis
   Copy the timestamp from a source AVI file to the converted equivalent MP4 file.
.Description
   The purpose is so that I can easily associate the AVI and the equivalent MP4 file by timestamp, 
   when sorting or searching.
.Example
   Copy-SourceTimeStampToTarget "c:\temp\source.avi" "c:\temp\target.mp4"
#>
function Copy-SourceTimeStampToTarget ($sourceAvi, $targetMp4) {
    $srcTime = Get-Item  $sourceAvi
    $targetTime = Get-Item $targetMp4
    $targetTime.LastWriteTime = $srcTime.LastWriteTime
    $targetTime.CreationTime = $srcTime.CreationTime
}

<#
.Synopsis
   Given the named folder, count the number of files of the named type
.Example
   Get-FileTypeCount "c:\temp" "mp4"
#>
function Get-FileTypeCount ($folder, $extension) {
    $set = (Get-ChildItem $folder -Filter "*.$extension") | Measure-Object
    $setCount = $set.Count
    "In folder [$folder], there are [$setCount] files of type [$extension]"
    start-sleep 2
}

<#
.Synopsis
   Given the full path of a file, return its name only, including extension.
   For the example below, "source.avi" would be returned.
.Example
   Get-FileNameFromFullPath "c:\temp\source.avi"
#>
function Get-FileNameFromFullPath ($file) {
    Split-Path -Path $file -Leaf
}

<#
.Synopsis
   Given the full path of a file, return all parent folders, but not the filename.
   For the example below, "c:\temp" would be returned.
.Example
   Remove-FileNameFromFullPath "c:\temp\source.avi"
#>
function Remove-FileNameFromFullPath ($file) {
    [System.IO.Path]::GetDirectoryName($file)
}

<#
.Synopsis
   With the assumption that the passed file is of type video, return the 
   (COM-based) formatted string that gives a video's duration.
.Example
   Get-VideoDuration "c:\temp\source.avi"
#>
function Get-VideoDuration ($fullPath) {
    $LengthColumn = 27
    $objShell = New-Object -ComObject Shell.Application 
    $objFolder = $objShell.Namespace($(Remove-FileNameFromFullPath $fullPath))
    $objFile = $objFolder.ParseName($(Get-FileNameFromFullPath $fullPath))
    $objFolder.GetDetailsOf($objFile, $LengthColumn)
}

<#
.Synopsis
   Convert an avi to mp4 file. It wraps AviToMp3 and provides metrics
   about the conversion.
.Example
   Convert-FromAviToMp4File "c:\temp\source.avi" "c:\temp\target.mp4"
#>
function Convert-FromAviToMp4File ($aviSource, $mp4target, $handbrakeFolder) {
    $source
    $target
    start-sleep 2
    $startTime = Get-Date
    AviToMp4 -sourceAvi $source -targetMp4 $target -h $handbrakeFolder
    $endTime = Get-Date
    $secondsToConvert = [int] ($endTime - $startTime).TotalSeconds
    Get-Date | Out-File -Append $logFile
    "Completed conversion of [$source] in [$secondsToConvert] seconds" | Out-File -Append $logFile 
    $sourceSizeInMb = [int] ((Get-Item -Path $source).Length/1MB)  
    $targetSizeInMb = [int] ((Get-Item -Path $target).Length/1MB)  
    "sourceSizeInMb: [$sourceSizeInMb]" | Out-File -Append $logFile 
    "targetSizeInMb: [$targetSizeInMb]" | Out-File -Append $logFile 

    $conversionRate = $sourceSizeInMb/$secondsToConvert
    "Conversion rate was [$conversionRate]MB per second" | Out-File -Append $logFile 
    $compressionRatio = $sourceSizeInMb/$targetSizeInMb
    "Compression ratio was $compressionRatio ([$sourceSizeInMb]/[$targetSizeInMb])" | Out-File -Append $logFile 
    $aviDuration =  Get-VideoDuration -fullPath $source
    $mp4Duration =  Get-VideoDuration -fullPath $target
    "Avi duration: [$aviDuration]" | Out-File -Append $logFile 
    "MP4 duration: [$mp4Duration] (these should match)" | Out-File -Append $logFile 


    Get-Date | Out-File -Append $logFile

    $baseName | clip.exe 
    start-sleep 5
}

<#
.Synopsis
   [ENTRY POINT]
   Convert a batch of avi files to mp4 format with the same name and location, apart from the extension.
.Description
   Given a folder with a set of avi files, convert each of those files to mp4 format, in the same folder.
   If the target mp4 name already exists, then skip.
   Following the conversion the file creation and amendment times get copied from the source/avi to the target/mp4.
   And finally, the target name (minus extension) is copied to the clipboard so that you can check the details in Explorer/search.
   The rootnames rename the same, only the extension changes.
   Dependencies: Handbrake CLI already installed in the given location.
   See - https://handbrake.fr/downloads2.php
.Example
   Convert-AviBatchToMp4 -v "G:\VideosCollection\TheRest" -h "C:\temp\HandBrakeCLI-1.0.7-win-x86_64"
#>
function Convert-AviBatchToMp4 ($videoFolder, $handbrakeFolder) {
   $logFile = "$currDir/Conversion.$(Get-Random).log"

   Set-Location $videoFolder
   Get-FileTypeCount -folder $videoFolder -extension "avi" | Out-File -Append $logFile 
   Get-FileTypeCount -folder $videoFolder -extension "mp4" | Out-File -Append $logFile 

   $aviList = gci -Filter *.avi
   #$aviList = gci -Filter "2000-05-30 19.03.13 2001 Em grandad Wells paper plane.avi"

   $aviList | ForEach-Object {
       $currentAvi = $_
       $baseName = $currentAvi.BaseName
       $source = "$videoFolder/$baseName.avi"
       $target = "$videoFolder/$baseName.mp4"

       if (!( Test-Path $target)) {
           Convert-FromAviToMp4File $source $target $handbrakeFolder
       }
       Copy-SourceTimeStampToTarget -sourceAvi $source -targetMp4 $target
   }
}

#cmd.exe '$handBrakeDir\HandBrakeCLI.exe -Z "Fast 1080p30" -i "G:\VideosCollection\Videos\2002-03-03 15.12.42 emma4.avi" -o "C:\MpegUtil\first2.mp4" --verbose=1'
