param([string]$path)

class ExcelFile{
  [string]$fullPath
  [System.Object]$app 
  [System.Object]$book

  ExcelFile([string] $inputPath){
    $this.fullPath = Convert-Path($inputPath)
    $this.app = New-Object -ComObject Excel.application
    $this.app.Visible = $false
  }

  open(){
    try{
      $this.book = $this.app.Workbooks.Open($this.fullPath)
    }catch{
      Write-Host -foreground Red "Cannot open file. `: " + $this.fullPath
      exit
    }
  }

  close(){
    $this.app.DisplayAlerts = $false
    $this.book.close()
    $this.app.DisplayAlerts = $true
    $this.app.quit()
  }

  [System.Collections.Generic.List[string]] getStyles(){
    $styleNames = New-Object 'System.Collections.Generic.List[string]'
    foreach($style in $this.book.Styles){

      if(-not $style.BuiltIn){
        $styleNames.Add($style.Name)
      }      
    }
      return $styleNames
  }

[System.Collections.Generic.List[string]] getNames(){
  $visibleNames = New-Object 'System.Collections.Generic.List[string]'
  $visibleNames.Add("-- Visible names")
  $unvisibleNames = New-Object 'System.Collections.Generic.List[string]'
  $unvisibleNames.Add("-- Unvisible names")
  foreach($name in $this.book.Names){

    if($name.visible){
      $visibleNames.Add($name.Name + " ... refersto: " + $name.refersto)
    }else{
      $unvisibleNames.Add($name.Name + " ... refersto: " + $name.refersto)
    }

  }
  return $visibleNames + $unvisibleNames
}

  [System.Collections.Generic.List[string]] getHeaderFooter(){
    $headerFooter = New-Object 'System.Collections.Generic.List[string]'
    foreach($sheet in $this.book.Sheets){
      foreach($page in $sheet.pagesetup){
        $headerFooter.Add("---- Sheet=" + $sheet.name)
        $headerFooter.Add("Header left= " + $page.LeftHeader + "`r`n" + "Header center= " + $page.CenterHeader + "`r`n" + "Header right=" + $page.RightHeader)
        $headerFooter.Add("Footer left= " + $page.LeftFooter + "`r`n" + "Footer center= " + $page.CenterFooter + "`r`n" + "Footer right=" + $page.RightFooter)
      }
    }
    return $headerFooter
  }
}

if($path -eq ""){
  Write-Host -foreground red "The argument needs to be exist path. e.g. powershell ./excel-cheker.ps1 .\file.xlsx"
}else{
  $file = New-Object ExcelFile($path)
  $file.open()

  Write-Host -foreground Yellow ("`r`n******** file : " + $path)
  
  Write-Host -foreground Cyan "`r`n------- Styles"
  foreach($outputStyle in $file.getStyles()){
      Write-Host -foreground Cyan $outputStyle
  }
  Write-Host -foreground Magenta "`r`n------- Names"
  foreach($outputName in $file.getNames()){
      Write-Host -foreground Magenta $outputName
  }
  Write-Host -foreground Green "`r`n------- Headers And Footers"
  foreach($outputValue in $file.getHeaderFooter()){
      Write-Host -foreground Green $outputValue
  }
  $file.close()
  Write-Host -foreground Yellow "`r`n******** END"
}
