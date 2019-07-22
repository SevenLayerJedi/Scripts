function Change-NameLength([string]$pdfDir, [int]$nameLength){
  if ($($pdfDir.length) -lt 1 -OR -$($nameLength -lt 1)){
    Write-host ""
    Write-host " [USAGE]"
    Write-host "     Change-NameLength -pdfDir `"C:\users\bob\pdf_dir`" -nameLength 254"
    Write-host ""
    return
  }else{
    ######################################
    # .variables                         #
    ######################################


    # Directory of pdfs
    # Not tested but probally wont work on unc paths
    #$pdfDir = "C:\toolshed\shorten_pdf_name"

    # Length to shorten to
    #$nameLength = 254


    ######################################
    # .functions                         #
    ######################################


    function shorten_subject([string]$messageSubject){
      if ($($messageSubject.length) -gt $($nameLength)){
        try{
          $newSubject = $messageSubject[0..$($nameLength - 5)] -join ""
        }catch{
          Write-host " [+] DANGER, DANGER DANGER, Will Robinson!"
        }
      }else{
        $newSubject = $messageSubject
      }
      return $newSubject
    }


    ######################################
    # .main                              #
    ######################################

    # Get all the PDFs in the directory
    # This does not do recursive lookup
    Write-host " [+] GETTING PDFs" -foregroundcolor green
    $allPdfFiles = gci -Path $pdfDir | ?{$_.Extension -eq ".pdf"}

    # Run through a loop and change all the names
    # so the length is equal to the nameLength variable
    foreach($f in $allPdfFiles){
      if ($($f.name.length - 4) -gt $($nameLength)){
        Write-host " [+] CHANGING NAME: $($f.name[0..32] -join '')" -foregroundcolor green
        $shortSubject = shorten_subject -messageSubject $($f.name)
        $oldName = $f.fullname
        $newName = $f.directoryname + "\$($shortSubject).pdf"
        # Write-host "  [+] OLD LENGTH: $($f.name.length)" -foregroundcolor green
        # Write-host "  [+] NEW LENGTH: $($shortSubject.length + 4)" -foregroundcolor green

        if(![System.IO.File]::Exists($newName)){
          Rename-Item "$($oldName)" "$($newName)" -Force
        }else{
          $newName = $($f.name[0..$($nameLength - 10)]) -join ""
          $newName = $f.directoryname + "\$($newName)" + "(" + [string]$(random -Minimum 10000 -Maximum 99999) + ").pdf"
          #Write-host "  [+] RENAME THIS: $($oldName)" -foregroundcolor green
          #Write-host "  [+] TO THIS: $($newName)" -foregroundcolor green
          Rename-Item "$($oldName)" "$($newName)" -Force
        }
        # Write-host "  [+] RENAME THIS: $($oldName)" -foregroundcolor green
        # Write-host "  [+] TO THIS: $($newName)" -foregroundcolor green
      }else{
        Write-host " [+] NO CHANGE: $($f.name[0..32] -join '')" -foregroundcolor cyan
        Write-host "  [+] NAME LENGTH: $($f.name.length)" -foregroundcolor cyan
      }
    }
  }
}

