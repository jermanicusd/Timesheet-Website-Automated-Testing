
$username = "";
$pwdfile = "..\File.txt";
$loginUrl = "https://";
$logoutUrl = "https://";
$addTimeSheet = "https://";
$currentTimeSheet = "https://";
$hours = "8";

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

#Get day of week
$date = Get-Date
$dayofMonth = $date.day
$dayofWeek = $date.DayOfWeek
$dayofMonthP = ($date.day +1)
$dayofMonthM = ($date.day -1)

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
    $doc.IHTMLDocument3_getElementsByName("username").ie8_item(0).value = "$username";

    # Find the password field and set the value to that of the result 
    # of a call to the get-password function with the paramter defined at top 
    $doc.IHTMLDocument3_getElementsByName("password").ie8_item(0).value = (get-password $pwdfile)

    # Find and click the submit button 
    $doc.IHTMLDocument3_getElementById("button_ok").click()

    # Wait until login is complete 
    while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; } 

} catch {$null}

#Check if Monday to create new timesheet
if ($dayofWeek -eq "Monday") 
    { 
    #Load Create Timesheet Page
    $ie.navigate($addTimeSheet);
    while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }    #wait for browser idle

    #Click Save to create, Date is auto set
    $doc.IHTMLDocument3_getElementById("button_save").click()
     }

#browse to current timesheet
$ie.navigate($currentTimeSheet);
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 5; }    #wait for browser idle

#Enter hours for current day
#$ie.Refresh()
#while ($ie.Busy -eq $true) { Start-Sleep -Seconds 5; }    #wait for browser idle

$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofMonth").ie8_item(0).focus();
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 5; }    #wait for browser idle

$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofWeek").ie8_item(0).value = $hours;
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 5; }    #wait for browser idle

#$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofMonthm").ie8_item(0).select();
#$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofMonthm").ie8_item(0).fireevent("onchange")

#$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofMonthP").item(0).setactive();
#$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofMonthP").item(0).select();
#$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofMonthM").item().setcapture();
#$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofMonth").item(0).click();
#$doc.IHTMLDocument3_getElementsByName("d_r0_$dayofMonth").item(0).select();
#$doc.IHTMLDocument3_getElementsByName("d_r1_$dayofMonth").item(0).onclick();
#$doc.IHTMLDocument3_getElementsByName("d_r1_$dayofMonthM").item(0).onfocus();
#$doc.IHTMLDocument3_getElementsByName("d_r1_$dayofMonth").item(0).onselect();
#$doc.IHTMLDocument3_getElementsByName("d_r1_$dayofMonth").item(0).focus();


#Save
$doc.IHTMLDocument3_getElementById("button_save").click();
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }    #wait for browser idle

#Verify hours saved


#Submit if Friday
if ($dayofWeek -eq "Friday") { 
    $doc.IHTMLDocument3_getElementById("button_submit").click(); 
    while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }    #wait for browser idle

    $doc.IHTMLDocument3_getElementById("ts_submit").click();
    while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }    #wait for browser idle
    }


#Logout
$ie.Navigate($logoutUrl)
while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }    #wait for browser idle

#Close Browser
$ie.Quit()

--------------------------------------------------------

#Troubleshooting/Testing bits
#$ie.document.getElementsByName("username").ie8_item(0).value = "$username";
#Get-Member -InputObject $ie.document.getElementsByName("username") #.ie8_item(0).value = $username;
#$ie.document.getElementsByName("d_r0_$dayofmonth").ie8_item(0).value = "8";
#Get-Member -InputObject $ie.Document.getElementsByName("d_r0_2")
#$p = ($ie.document.getElementsByTagName("iframe"))[0]
#$q = $p.contentWindow.document
#$t = $q.getelementbyname("d_r0_2")
#$ie.Document.body | Out-File -FilePath "C:\web.txt"

#$doc.getElementsByName("username").ie8_item(0).value = "$username";
#$doc.getElementsByName("password").ie8_item(0).value = "$password";
#($doc.getElementsByName("button_ok") |select -first 1).click();

#($ie.document.getElementsByName("home.quick-links.current-timesheet") |select -first 1).click();
