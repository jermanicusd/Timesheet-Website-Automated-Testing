
$username = "";
$pwdfile = "..\File.txt";
$loginUrl = "https://.com";
$logoutUrl = "";
$addTimeSheet = "url";
$StartT1 = "8";
$StartT2 = "AM";
$break = "60";
$EndT1 = "5";
$EndT2 = "PM";

# Function to return a password from an encrypted file 
Function get-password([string]$credentialsfile) { 
    
    #Check to see if the file exists 
    if (-not (Test-Path $credentialsfile)){ 
        
        #If not, then prompt user for the credential 
        $creds = Get-Credential 
        
        #Get the password part 
        $encpassword = $creds.password 
        
        # Convert it from secure string and save it to the specified file 
        $encpassword |ConvertFrom-SecureString |Set-Content $credentialsfile} 

    else { 

        #If the file exists, get the content and convert it back to secure string 
        $encpassword = get-content $credentialsfile | convertto-securestring 
    } 
    
    # Use the Marshal classes to create a pointer to the secure string in memory 
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($encpassword) 
    
    # Change the value at the pointer back to unicode (i.e. plaintext) 
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) 
    remove-variable encpassword,ptr 
    
    # Return the decrypted password 
    return $pass 
} 

#Get date
$date = (Get-Date).ToString("MM-dd-yyyy")

#Open IE Page ($ie.visible = $false to run in background)
$ie = New-Object -com internetexplorer.application;
$ie.visible = $true;
$ie.navigate($loginUrl);
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }    #wait for browser idle

#Assign document to $doc
$doc = $ie.document

#login
try { 
    # Find the username field and set the value to that of our variable 
    $doc.IHTMLDocument3_getElementsByName("loginid").ie8_item(0).value = "$username";

    # Find the password field and set the value to that of the result 
    # of a call to the get-password function with the paramter defined at top 
    $doc.IHTMLDocument3_getElementsByName("password").ie8_item(0).value = (get-password $pwdfile)

    # Find and click the submit button 
    $doc.IHTMLDocument3_getElementByID("SignIn").click()

    # Wait until login is complete 
    while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; } 

} catch {$null}

#TimeSheet Page
$ie.navigate($addTimeSheet);
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }

#Start Time
$doc.IHTMLDocument3_getElementById("_fid_7").ie9_value = $StartT1
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
$doc.IHTMLDocument3_getElementById("_fid_9").ie9_value = $StartT2
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }

#Break
$doc.IHTMLDocument3_getElementById("_fid_11").value = $break
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }

#End Time
$doc.IHTMLDocument3_getElementById("_fid_13").value = $EndT1
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
$doc.IHTMLDocument3_getElementById("_fid_15").value = $EndT2
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }

#Save
$doc.IHTMLDocument3_getElementById("footerSaveButton").click()

#Close Page
$ie.Quit()

---------------------------------------------------------------------

#Troubleshooting/Testing bits
#Always use the following methods instead of the native ones:
#IHTMLDocument3_getElementsByTagName
#IHTMLDocument3_getElementsByName
#IHTMLDocument3_getElementByID

#Get-Member -InputObject $ie.Document.getElementsByTagName("main")
#($ie.document.getElementbyID("iframe").contentWindow.document.getElementbyID('loginForm')| Where-Object {$_.name -eq "_ssoUser"}).value = "username"

#$#link = @($ie.Document.getElementsByTagName('main')) | Where-Object {$_.innerText -eq 'login'}

#[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11"
#$URI = "https://dell.quickbase.com/db/main?a=signin";
#$html = Invoke-WebRequest -Uri $URI
#($HTML.ParsedHtml.getElementsByTagName(‘’) | Where{ $_.className -eq ‘UserInput WithPadding’ } ).innerText

#$html.ParsedHtml | Get-Member

#$frame = ($ie.document.getElementsByTagName("iframe"))[0]
#$Iframe = $frame.contentWindow.document

#($ie.document.IHTMLDocument3_getElementsByName("loginid") |select -first 1).value = $username
