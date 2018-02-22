################################################################################
# Pull-ADLS
#
# Authors:     Justin McCormick, Caleb Gross
#
# Description: Given a valid UTM code, this script uses your CAC certificates
#              to pull CBT records from ADLS. Once you've queried ADLS and
#              pulled the CSV records, this script will parse those files.
#
################################################################################

# Help with debugging.
$ErrorActionPreference = 'Stop'

# Initialize a few variables.
$unsafe_chars   = "[^a-zA-Z0-9\(\)\-\. ]"
$records_folder = "$PWD\Training_Records"
$records_file   = "$records_folder\ADLS_Training.csv"
$courses_file   = "$PWD\courses_tracked.txt"
$certs_folder   = "$PWD\Certificates"

# Locate user's CAC ID certificate.
$cert_thumb = (Get-ChildItem "Cert:\CurrentUser\My\" | Where-Object {$_.FriendlyName -like "*id*"}).Thumbprint

# Create session to store cookies between web requests, and add cookies to session.
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.cookies.add((New-Object System.Net.Cookie -ArgumentList "AFPORTAL_LOGIN_AGREEMENT","accepted","/","af.mil"))
$session.cookies.add((New-Object System.Net.Cookie -ArgumentList "BUILDING","Administration","/","golearn.adls.af.mil"))

function file_exists($file) {
    if (Test-Path "$file") {
        Write-Host "[+] Found $file."
        if ( (get-childitem $file).length -eq 0 ) {
            Write-Host "[-] $file is empty."
        }
        return (get-childitem $file).length
    }
    else {
        Write-Host "[-] $file does not exist."
        return -1
    }
}

# Authenticate to Access.asp.
function authenticate() {
    
    # Authenticate to CACRequest.aspx and store resulting cookie. Verify that authentication was successful.
    $url      = "https://golearn.adls.af.mil/CACAuthentication/CACRequest.aspx"
    $cac_auth = Invoke-WebRequest -Uri $url -CertificateThumbprint $cert_thumb -WebSession $session
    if ($cac_auth -like "*GoLearn - Powered By The ADLS :: Main Page*") {Write-Host "[+] Successful CAC authentication."} 
    else {Write-Host "[-] Error logging in with CAC certificates."; exit}
   
    # Using the cookie from CACRequest.aspx (above), send a simple GET request to Access.asp and scrape resulting organization name and ID.
    $url         = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/Access.asp"
    $referer     = "https://golearn.adls.af.mil/kc/admin/faculty_lounge_ttl.asp?table=&function=faculty_lounge"
    $utm_init    = Invoke-WebRequest -Uri $url -CertificateThumbprint $cert_thumb -WebSession $session -Headers @{'referer' = $referer}
    $org_name    = ($utm_init.AllElements | Where-Object {$_.id -eq "OrgName"}).value
    $org_id      = ($utm_init.AllElements | Where-Object {$_.id -eq "Org_ID"}).value
	Write-Host "[*] Organization Name: $org_name"
    Write-Host "[*] Organization ID:   $org_id"

    # If no prior UTM code has been stored, prompt user for UTM code.
    if ($global:UTM_CODE -eq $null) {
	    $global:UTM_CODE = Read-Host "    Enter UTM code" -AsSecureString
	    $global:UTM_CODE = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($UTM_CODE))
    }
    else {
        Write-Host "[*] Using previously entered UTM code."
    }

    # Using org name, org ID, and UTM code, authenticate to Access.asp. Verify that authentication was successful.
    $data = @{
        agreement='agreed';
        actionDo='Access';
        Org_ID=$org_id;
        OrgName=$org_name;
        FirstTime='False';
        AccessCode=$UTM_CODE;
        }
    $utm_auth = Invoke-WebRequest -Uri $url -CertificateThumbprint $cert_thumb -WebSession $session -Headers @{'referer' = $referer} -Method 'POST' -Body $data
    if ($utm_auth -like "*UTM / UDM - Main Page*") {Write-Host "[+] Successful UTM/UDM authentication."} else {Write-Host "[-] Error authenticationg with supplied UTM code."; exit}

    return @{
        'id'   = $org_id;
        'name' = $org_name;
        }

}

# Build a list of URLs for each tracked course.
function get_course_urls() {

    if ((file_exists $courses_file) -lt 0) {return}

    # Locate file which maps course names to course IDs.
    $courses = Get-Content -Path "$courses_file"

    # Generate URLs for each tracked course.
    $base_url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/OrgGroupReport/OrgGroupReportByCourse_Excel.asp?ReportName=SearchUsersDetails&strAction=SearchUsersDetails&adls_org_id=&blnAllSubOrgs=true&strRankType='AD','CV','CT','DP','OT'"
    $course_urls = @{}
    foreach ($course in $courses) {
        $course_name,$course_num = $course -split "="
        $course_urls.Add("$course_name", $base_url + "&customerorg_ident=" + $org_info.Item('id') + "&CourseID=$course_num")
    }

    return $course_urls
}

# Pull records for each tracked course.
function update_records () {

    if ((file_exists $courses_file) -le 0) {return}

    # Prepare records folder for HTML download.
	if ((file_exists $records_folder) -le 0) {
		Write-Host "[*] Creating folder '$records_folder' to store course completion records."
		New-Item "$records_folder" -Force -ItemType directory | Out-Null
	} else {
		Write-Host "[*] Clearing records from '$records_folder'."
		Remove-Item "$records_folder/*"
	}

    # Download HTML file at each course URL.
    $course_urls = get_course_urls
    $referer     = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/OrgGroupReport/OrgGroupReportByCourse.asp"
    Write-Host "[*] Running queries..."
	Try {
        foreach ($course_name in $course_urls.Keys) {
            $course_url = $course_urls.Item($course_name)
            $outfile    = "$records_folder/" + ($course_name -replace $unsafe_chars,"") + ".html"
            Write-Host "    $course_name"
            Invoke-WebRequest -Uri $course_url -CertificateThumbprint $cert_thumb -WebSession $session -Headers @{'referer' = $referer} -OutFile $outfile
        }
	}
	Catch {
		Write-Host "[-] Error running queries."
	}

    # Load HTML files into dictionary.
	$user_dict = @{}
    foreach ($file in Get-ChildItem "$records_folder/*.html") {

        # Create HTML object from HTML file contents.
		$src = Get-Content $file.FullName -raw
        $html = New-Object -ComObject "HTMLFile"
        $html.IHTMLDocument2_write($src)

        # Get rows from HTML object, removing header (index 0) and footer (index len-1).
		$rows = @($html.getElementsByTagName("tr"))
        $rows = $rows[1..($rows.length-2)]
        
        # Update dictionary with training records from each username in HTML file.
		foreach ($row in $rows) {
			$cells = $row.children
			$user_name = (($cells.item(1)).innerText).Replace(",","")

            # If dictionary does not yet contain username, add the username.
			if (($user_dict.Keys -notcontains $user_name)) {$user_dict += @{$user_name=@{}}}

            # Update username with training records.
            $user_dict[$user_name].Add($file.BaseName,($cells.item(7)).innerText)
		}
	}

    # Begin to build CSV file.
	Write-Host "[*] Compiling final CSV file."
	$table = @()

    # Get sorted list of course names.
	$course_names = @($course_urls.keys)
	$course_names = ($course_names | Sort-Object)
    
    # Build header row.
   	$headers = ""
	$headers_item = @("User")
	foreach ($course_name in $course_names) {
		$headers_item += $course_name -join ","
	}
	$headers += $headers_item -join ","

	# Get dates for each user and append to table.
	foreach ($user in $user_dict.Keys) {

        # Sort user's course names.
		$user_dict.$user.Keys = $user_dict.$user.Keys | Sort-Object
		
        # Get date for each course name and append to current row.
        $row = @($user)
		foreach ($course_name in $course_names) { $row += $user_dict.$user.$course_name }

        # Comma-separate user's row elements, and apppend to table.
		$table += $row -join ","
	}

	# Remove HTML files.
	Write-Host "[*] Removing HTML files."
	if (Test-Path "$records_folder") {Remove-Item "$records_folder/*.html"}

	# Create CSV file.
	try {
		$headers | Out-File "$records_file" -Encoding ascii
		$table | Sort-Object | Out-File "$records_file" -Encoding ascii -Append
	} catch {
		Write-Host "[-] Error saving file. Ensure that the CSV file is closed."
	}

    return

}

# Check a member's course completion dates.
function display_member_dates () {

    if ((file_exists $records_file) -lt 0) {return}
    
    $adls_cont = Import-Csv "$records_file"

    # Query user for user to check.
    $counter = 1
    foreach ($item in $adls_cont) {
        Write-Host "`t[$counter] " $item.User 
        $counter += 1
    }
    $chk_user = Read-Host "`nWhich user would you like to check?"

    # Display user information.
    Clear-Host
    $adls_cont[$chk_user-1] | Format-List
}

# Download course completion certificate for a single member.
function get_single_certificate() {

    if ((file_exists $records_file) -le 0) {return}

    # Query user for member to check.
    $adls_cont = Import-Csv "$records_file"
    Write-Host "`n[*] Member list:"
    $counter = 1
    foreach ($item in $adls_cont) {
        Write-Host "`t[$counter] " $item.User 
        $counter += 1
    }
    $chk_user = Read-Host "`nWhich member would you like to check?"

    # Get member details.
    $user_vars = ($adls_cont[$chk_user-1].User).Split(" ");
    $first_name = $user_vars[1];
    $last_name = $user_vars[0]
    $user_details = get_user_details $last_name $first_name
    $result_no = $user_details[1]
    $user_id = $user_details[3]

    # Get member's list of completed courses and corresponding certificate URLs.
    $url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserRemoval_UserProgress_AJAX.asp?ADLS=&USERS_IDENT=$user_id&CounterVar=$result_no&sid=" + (Get-Random -Minimum 1 -Maximum 10)
    $cert_response = Invoke-WebRequest -Uri $url -CertificateThumbprint $cert_thumb -WebSession $session -Headers @{'referer' = $url}

    # Parse list of certificate URLs.
    $cert_links = @()
    $cert_response -split "`n" | %{if ($_ -like "*certificate*"){(($_ -split '"')[5]) -match ".*strCourseID=(......).*" | Out-Null; $cert_links += $Matches[1] }}

    # Parse list of completed courses.
    $page_array = @()
    $cert_response -split "`n" | %{if (($_ -like "*width:390px*") -or ($_ -like "*certificate*")){$page_array += $_}}
    $cert_name = @()
    $cert_response -split "`n" | %{if($_ -like "*certificate*"){$cert_name += (($page_array[([array]::IndexOf($page_array,$_)+1)]) -split ">")[1] -replace "</div",""  }}

    # Prompt user to select course.
    Write-Host "`n[+] Which course would you like to download a completion certificate for?"
    $c_num = 1
    $cert_name | %{Write-Host "`t[$c_num] $_";$c_num++}
    $menu_sel = [int](Read-Host "`n[-] Input Selection")
    if ($menu_sel -gt $cert_links.Count) {
        Write-Host "[-] Invalid selection."
        break :menu
    }
    $course_id = $cert_links[$menu_sel-1]

    # Download certificate for selected member and course.
    download_certificate $user_id $course_id
}

# Download course completion certificate for all members.
function get_batch_certificates () {

    if ((file_exists $records_file) -le 0) {return}
    
    $adls_cont = Import-Csv "$records_file"
    
    # Prompt user for course selection.
    $links = get_course_urls
    Write-Host "[*] Course list:"
    $counter = 1
    foreach ($course in $links.Keys) {
        Write-Host "`t[$counter] $course"
        $counter++
    }
    $menu_sel = ([int](Read-Host "For which course would you like to download certificates of completion?")-1)

    # Get ID for selected course.
    $c=0; $links.GetEnumerator() | %{if($c -eq $menu_sel){$my_link = $_.Value;$my_cs = $_.Key};$c++}
    Write-Host "[*] Checking for $my_cs..."
    $my_link -match '.*CourseID=(.....[C0123456789]).*' | Out-Null
    $course_id = $Matches[1]

    $url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp"
    $referer = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp?ADLS="

    foreach ($item in $adls_cont) {

        # Get member details.
        $un = $item.User -split " "
        $first_name = $un[1]
        $last_name = $un[0]
        $user_details = get_user_details $last_name $first_name
        $user_id = $user_details[3]

        # Download certificate for selected member and course.
        download_certificate $user_id $course_id
    }
}

# Get a member's detailed information.
function get_user_details($last_name, $first_name) {

    # Get member details.
    $url      = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp"
    $referer  = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp"
    $data     = "FormSubmitted=Submitted&ADLS=&strFname=$first_name&strLname=$last_name&strSSN="
    $user_details_html = Invoke-WebRequest -Uri $url -CertificateThumbprint $cert_thumb -WebSession $session -Headers @{'referer' = $referer} -Method 'POST' -Body $data

    # Parse member details.
    $org_name = $org_info.Item('name').Replace("+"," ")
    $user_details = @()
    ($user_details_html -split "`n") | %{if ($_ -like "*>$org_name<*" -or $_ -like "*ShowDetails*"){$user_details += $_}}
    return $user_details | %{if($_ -like "*>$org_name<*"){($user_details[([int][array]::IndexOf($user_details,$_)-1)] -split ";")[1] -split "'"}}

}

# Download a course completion certificate.
function download_certificate($user_id, $course_id){

    if ((file_exists $certs_folder) -le 0) {
		Write-Host "[*] Creating folder '$certs_folder' to store course completion certificates."
		New-Item "$certs_folder" -Force -ItemType directory | Out-Null
	}

    # Download certificate HTML body.
    $url = "https://golearn.adls.af.mil/kc/certificate/aetc_cert_certif.asp?crs_ident=$course_id&kc_ident=kc0001&login=$user_id"
    $referer = "https://golearn.adls.af.mil/kc/certificate/aetc_cert_certif.asp"
    $cert_response = Invoke-WebRequest -Uri $url -CertificateThumbprint $cert_thumb -WebSession $session -Headers @{'referer' = $referer} -Method 'POST' -Body $data

    # Change local paths to remote paths for constituent images in HTML body.
    Write-Host "[+] Outputting Certificate for $last_name $course_id."
    $image_array = @()
    ($cert_response -split "`n")| %{if ($_ -match '.*="([0-9A-Za-z_.]*(jpg|gif))".*') {if ($image_array -notcontains $matches[1]){$image_array += $matches[1]}}}
    $final_out = ""
    ($cert_response -split "`n") | %{foreach ($image in $image_array){if ($_ -like "*$image*"){$_ = $_.Replace($image,"https://golearn.adls.af.mil/kc/certificate/$image")}}$final_out += $_} 
    $outfile = "$certs_folder\" + "$user_id $course_id.html".Replace(" ","_")
    $final_out | Out-File $outfile -Force
}

# Display menu of course-editing options.
function edit_courses() {

    # Create course tracking file if it doesn't exist already.
    if ((file_exists $courses_file) -lt 0) {
        Write-Host "[+] Creating $courses_file"
        New-Item "$courses_file" -ItemType file | Out-Null
    }

    :edit_menu while(1) {
        Write-Host "`n=== Edit Tracked Courses ===`n"
        Write-Host "`t [1] View currently tracked courses"
        Write-Host "`t [2] Add a course to the tracker"
        Write-Host "`t [3] Remove a course from the tracker"
        Write-Host "`n`t [0] Return to main menu`n"

        switch(Read-Host "Input selection") {
            1 {Clear-Host; Write-Host; view_courses}
            2 {Clear-Host; Write-Host; add_course}
            3 {Clear-Host; Write-Host; remove_course}
            0 {Write-Host "`n[*] Returning to main menu."; return}
            default {Write-Host "`n[-] Invalid selection."; break :edit_menu}
        }
    }
}

# Display currently tracked courses.
function view_courses() {

    $course_file = Get-Content "$courses_file"
    Write-Host "[*] Courses currently being tracked:"
    $course_array = @()
    foreach ($line in $course_file) {
        $course_array += $line
        $temp = $line -split "="
        $cn = $temp[0]
        $cs_num = $temp[1]
        Write-Host "    $cs_num : $cn"
    }
    Write-Host "`n`n"
}

# Add new course to tracked courses.
function add_course() {

    # Search for courses based on search term.
    $search = Read-Host "Enter a search term for the course you would like to add"
    $url        = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/OrgGroupReport/OrgGroupReportByCourse.asp"
    $pg_results = Invoke-WebRequest -Uri $url -CertificateThumbprint $cert_thumb -WebSession $session -Headers @{'referer' = $url}
    $temp_array = @()
    ($pg_results -split "`n") | %{if (($_ -like "*$search*") -and ($_ -like "*option value*")) {$temp_array += $_}}

    # Display list of courses matching search term. Prompt user for course to add.
    $c = 1
    $course_array = @()
    Write-Host "`n[+] Matching courses:"
    foreach ($item in $temp_array) {
        $item -match ".*>(.*)<.*" | Out-Null
        $cn = ($matches[1]) -replace $unsafe_chars,""
        $item -match '.*value="(.*)">.*' | Out-Null
        $cs_num = $matches[1]
        $line = $cn + "=" + $cs_num
        $course_array += $line
        Write-Host "`t[$c] $cn : $cs_num"
        $c++
    }
    Write-Host "`n`t[0] Return to previous menu`n"
    $cs_to_add = [int](Read-Host "Which course would you like to add?")
    if ($cs_to_add -gt $temp_array.Count) {Write-Host "`n[-] Invalid course selection.`n";return}
    if ($cs_to_add -eq '0') {Clear-Host; Write-Host "`n[*] Returning to previous menu.`n";return}

    # Add selected course.
    $exists = 0
    $info = ($course_array[[int]$cs_to_add-1]) -replace "\*",""
    $dat = Get-Content "$courses_file"
    ($dat -split "`n") | %{if ($_ -like "*$info*"){Write-Host "`n[-] Duplicate course found. Unable to add.";$exists = 1} }
    if ($exists -ne 1) {
        Write-Host "`n[+] Adding $info to $courses_file"
        $info | Out-File -Append -Force -FilePath "$courses_file"
    }

    return
}

# Remove course from tracked courses.
function remove_course() {

    $course_file = Get-Content "$courses_file"

    # Prompt user for course to remove.
    Write-Host "Which course would you like to remove?"
    $c = 1
    $course_array = @()
    foreach ($line in $course_file) {
        $course_array += $line
        $temp = $line -split "="
        $cn = $temp[0]
        $cs_num = $temp[1]
        Write-Host "`t[$c] $cn : $cs_num"
        $c++
    }
    Write-Host "`n`t[0] Exit`n"
    $menu_sel = [int](Read-Host "[-] Input Selection")
    if ($menu_sel -gt $course_file.count) {Write-Host "`n[-] Invalid course selection.";return}
    if ($menu_sel -eq '0') {Clear-Host; Write-Host "`n[*] Returning to previous menu.`n";return}

    # Remove course.
    $removed = 0
    $line_to_remove = $course_array[[int]$menu_sel-1]
    $new_array = @()
    $course_file | %{if (($_ -like "*$line_to_remove*") -and ($removed -eq 0)){Write-Host "`n[*] Removing $line_to_remove";$removed = 1}else {$new_array += $_}}
    $new_array | Out-File -Force -FilePath "$courses_file"

    return
}

# TODO: Import a CSV into records file.
function import_csv () {
    if (Test-Path "$records_file") {
        Write-Host "[+] Found ADLS Training File"
        $final_cont = Get-Content "$records_file"
        $dat_array = @()
        Write-Host "`n[+] Which CSV Would you like to import?"
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $PWD
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $adls_csv = $OpenFileDialog.filename
        $adls_content = Get-Content $adls_csv
        if (($final_cont.count) -ne $adls_content.Count) {Write-Host "`n[-] Incorrect Line Count - Create Template First`n";break :menu}
        $count = 1
        foreach ($line in $adls_content) {
            foreach ($fin_line in $final_cont) {
                $area = $line -split ","
                $fin_split = $fin_line -split ","
                if ($count -eq 1) {
                    foreach($in_split in $fin_split) {
                        if ($in_split -like $area[1]) {
                            Write-Host "[+] Duplicate Column Found... Cannot Import CSV`n"
                            return
                        }
                    }
                    $count++
                }
                if ($area[0] -like $fin_split[0]) {
                    $dat_array += ($fin_line + "," + $area[1])
                    Write-Host "[*] Adding" $fin_split[0] "with" $area[1]
                }
            }
        }
        $headers = $dat_array[0]
        $table = ($dat_array | Select-Object -Skip 1)
        $headers | Out-File "$records_file" -Encoding ascii
        $table | Sort-Object | Out-File "$records_file" -Encoding ascii -Append

    } else {
        Write-Host "`n[-] Cannot Find ADLS Training File.  Try Querying First"
    }

}

# TODO: Create CSV template to be imported.
function create_sample () {
    if (Test-Path "$records_file") {
        Write-Host "[+] Found ADLS Training File"
        $adls_cont = Import-Csv "$records_file"
        $template = @()
        Write-Host "[+] Creating Sample Template"
        $col_name = Read-Host "[-] What is the name of the column you are adding?"
        $template += "User,$col_name"
        foreach ($item in $adls_cont) {
            $template += ($item.User + ",Data")
        }
        $template | Out-File "$PWD/Import_Template.csv" -Encoding ascii
        Write-Host "[+] Template stored: Import_Template.csv`n"
    } else {
        Write-Host "[-] Cannot Find ADLS Training File.  Try Querying First"
    }
}


######################
# MAIN PROGRAM START #
######################

Clear-Host
Write-Host
$org_info = authenticate

if (Test-Path "$records_file") {
    $last_write = gci $records_file | Select-Object -ExpandProperty LastWriteTime
    $age = New-TimeSpan -Start (gci $records_file).LastWriteTime -End (Get-Date)
    if     ($age.Days -gt 0)  {$date_string = "$($age.Days) days"}
    elseif ($age.Hours -gt 0) {$date_string = "$($age.Hours) hours"}
    else                      {$date_string = "$($age.Minutes) minutes"}
    Write-Host "`n[*] Records file last updated $date_string ago at $last_write."
    if ($age.Days -ge 1) {
        Write-Host "[*] It's been over a day since you last queried ADLS...consider running option [1] first!"
    }
}
else {
    Write-Host "`n[-] Records file '$records_file' does not exist. Try running a query with option [1]."
}

#    Write-Host "`t[5] Import CSV"
#    Write-Host "`t[6] Create CSV Template to Import"
#        5 {Clear-Host; Write-Host; import_csv}
#        6 {Clear-Host; Write-Host; create_sample}
:main_menu while (1) {
    Write-Host "`n=== Pull-ADLS Main Menu ===`n"
    Write-Host "`t[1] Query ADLS to update course completion dates"
    Write-Host "`t[2] Display course completion dates"
    Write-Host "`t[3] Download course completion certificate for single member"
    Write-Host "`t[4] Download course completion certificate for entire organization"
    Write-Host "`t[5] Edit tracked courses"
    Write-Host "`n`t[0] Exit`n`n"

    switch (Read-Host "Input selection") {
        1 {Clear-Host; Write-Host; update_records}
        2 {Clear-Host; Write-Host; display_member_dates}
        3 {Clear-Host; Write-Host; get_single_certificate}
        4 {Clear-Host; Write-Host; get_batch_certificates}
        5 {Clear-Host; Write-Host; edit_courses}
        0 {exit}
        default {Write-Host "`n[-] Invalid selection."; break :main_menu}
    }
}
