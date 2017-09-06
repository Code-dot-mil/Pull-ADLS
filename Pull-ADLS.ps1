################################################################################
# ADLS CBT Record Puller
#
# Authors:     Justin McCormick, Caleb Gross
#
# Description: Given a valid UTM code, this script uses your CAC certificates
#              to pull CBT records from ADLS. Once you've queried ADLS and
#              pulled the CSV records, this script will parse those files.
#
# Note:        All code for educational purposes only.
#
################################################################################

# help with debugging
$ErrorActionPreference = 'Stop'

# initialize a few variables
$unsafe_chars = "[^a-zA-Z0-9\(\)\-\. ]"
$records_folder = "Training_Records"

# locate user's CAC ID certificate
$cert_thumb = (Get-ChildItem "Cert:\CurrentUser\My\" | Where-Object {$_.FriendlyName -like "*id*"}).Thumbprint

# create session to store cookies between web requests
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# add cookies to session
$session
$session.cookies.add((New-Object System.Net.Cookie -ArgumentList "AFPORTAL_LOGIN_AGREEMENT","accepted","/","af.mil"))
$session.cookies.add((New-Object System.Net.Cookie -ArgumentList "BUILDING","Administration","/","golearn.adls.af.mil"))
$session

# issue web request and return response
function request($url, $referer, $method, $data) {
    $webreq = New-Object SmartCard.webpage
    $webreq.url = $url
    $webreq.referer = $referer
    $webreq.method = $method
    $webreq.accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    $webreq.data = $data
    $webreq.contenttype = "application/x-www-form-urlencoded"
    return [SmartCard.mainclass]::web_req($webreq)
}

# LOGIN FUNCTION FOR ALL REQUESTS
function login_func() {
    
    # authenticate to CACRequest.aspx and store cookie
    $cac_auth = Invoke-WebRequest -Uri "https://golearn.adls.af.mil/CACAuthentication/CACRequest.aspx" -CertificateThumbprint $cert_thumb -WebSession $session

    # check for successful login
    if ($cac_auth -like "*GoLearn - Powered By The ADLS :: Main Page*") {Write-Host "[+] Successfully Logged In"} 
    else {Write-Host "[-] Error Logging in with Credentials"; exit}

    # using cookie from CACRequest.aspx, send a simple GET request to Access.asp
    $utm_url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/Access.asp"
    $utm_referer = "https://golearn.adls.af.mil/kc/admin/faculty_lounge_ttl.asp?table=&function=faculty_lounge"
    $utm_headers = @{
        'referer' = $utm_referer
        }
    $utm_init = Invoke-WebRequest -Uri $utm_url -CertificateThumbprint $cert_thumb -WebSession $session -Headers $headers

    # scrape organization name and ID from prior GET request to Access.asp
    $org_name = ($utm_init.AllElements | Where-Object {$_.id -eq "OrgName"}).value
    $org_id   = ($utm_init.AllElements | Where-Object {$_.id -eq "Org_ID"}).value
    Write-Host "[+] Using $org_id as the Organization ID"
	Write-Host "[+] Using $org_name as the Organization Name"

    # read in UTM code from user
    if ($global:UTM_CODE -eq $null) {
	        $global:UTM_CODE = Read-Host "[-] Input UTM Code" -AsSecureString
	        $global:UTM_CODE = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($UTM_CODE))
        }

    # using org name, org ID, and UTM code, authenticate to Access.asp
    $utm_data = @{
        agreement='agreed';
        actionDo='Access';
        Org_ID=$org_id;
        OrgName=$org_name;
        FirstTime='False';
        AccessCode=$UTM_CODE;
        }
    $utm_auth = Invoke-WebRequest -Uri $utm_url -CertificateThumbprint $cert_thumb -WebSession $session -Headers $utm_headers -Method Post -Body $utm_data
    if ($utm_auth -like "*UTM / UDM - Main Page*") {Write-Host "[+] Successfully Logged In UTM/UDM"} else {Write-Host "[-] Error Logging in with UTM Code"; exit}

    $links = get_course_urls

    return $links

}

# PULL ALL RECORDS FOR CBTS IN links VARIABLE FOR SQUADRON
function query_cbt () {

    # prepare records folder for html download
	if (!(Test-Path "$PWD/$records_folder")){
		Write-Host "`n[+] Creating Folder '$records_folder' to Store Records"
		New-Item "$records_folder" -Force -ItemType directory | Out-Null
	} else {
		Write-Host "`n[+] Deleting All Files in $records_folder"
		Remove-Item "$PWD/$records_folder/*"
	}
	Write-Host "[+] Putting Records in $records_folder"

    # download html file at each course URL
    $course_urls = login_func

	Try {
    
        foreach ($course_name in $course_urls.Keys) {
            $course_url = $course_urls.Item($course_name)
            $outfile = "$PWD/$records_folder/" + ($course_name -replace $unsafe_chars,"") + ".html"
            request $course_url $course_url "GET" "" | Out-File -FilePath $outfile -Force
            }
	}
	Catch {
		Write-Host "[x] Error Caught"
	}

    # load html files into dictionary
	$user_dict = @{}
    foreach ($file in Get-ChildItem "$PWD/$records_folder/*.html") {

        # create html object from html file contents
		$src = Get-Content $file.FullName -raw
        $html = New-Object -ComObject "HTMLFile"
        $html.IHTMLDocument2_write($src)

        # get rows from html object, removing header (index 0) and footer (index len-1)
		$rows = @($html.getElementsByTagName("tr"))
        $rows = $rows[1..($rows.length-2)]
        
        # update dictionary with training records from each username in html file
		foreach ($row in $rows) {
			$cells = $row.children
			$user_name = (($cells.item(1)).innerText).Replace(",","")

            # if dictionary does not yet contain username, add the username
			if (($user_dict.Keys -notcontains $user_name)) {$user_dict += @{$user_name=@{}}}

            # update username with training records
            $user_dict[$user_name].Add($file.BaseName,($cells.item(7)).innerText)
		}
	}

    # build CSV file
	Write-Host "`n[+] Creating Final CSV`n"
	$table = @()

    # get sorted list of course names
	$course_names = @($course_urls.keys)
	$course_names = ($course_names | Sort-Object)
    
    # build header row
   	$headers = ""
	$headers_item = @("User")
	foreach ($course_name in $course_names) {
		$headers_item += $course_name -join ","
	}
	$headers += $headers_item -join ","

	# get dates for each user and append to table
	foreach ($user in $user_dict.Keys) {

        # sort user's course names
		$user_dict.$user.Keys = $user_dict.$user.Keys | Sort-Object
		
        # get date for each course name and append to current row
        $row = @($user)
		foreach ($course_name in $course_names) { $row += $user_dict.$user.$course_name }

        # comma-separate user's row elements, and apppend to table
		$table += $row -join ","
	}

	# remove HTML files
	Write-Host "`n[+] Removing HTML files"
	if (Test-Path "$PWD/$records_folder") {Remove-Item "$PWD/$records_folder/*.html"}

	# create CSV file
	try {
		$headers | Out-File "$PWD/$records_folder/ADLS_Training.csv" -Encoding ascii
		$table | Sort-Object | Out-File "$PWD/$records_folder/ADLS_Training.csv" -Encoding ascii -Append
		Write-Host "[+] Creating Backup File`n"
		Copy-Item "$PWD/$records_folder/ADLS_Training.csv" "$PWD/$records_folder/ADLS_Training_BK.csv" -Force
	} catch {
		Write-Host "[*] Error saving file. Ensure that the CSV file is closed."
	}	

}

# CHECK INDIVIDUAL CBT DATES PER PERSON
function check_cbt () {
    Write-Host "[+] Checking CBT Dates"
    if (Test-Path "$PWD/$records_folder/ADLS_Training.csv") {
        Write-Host "[+] Found ADLS Training File"
        $adls_cont = Import-Csv "$PWD/$records_folder/ADLS_Training.csv"

        Write-Host "`n[+] Which user would you like to check?"
        $counter = 1
        foreach ($item in $adls_cont) {
            Write-Host "`t[$counter] " $item.User 
            $counter += 1
        }
        
        $chk_user = Read-Host "`n[-] Input Selection"
        $adls_cont[$chk_user-1]

    } else {
        Write-Host "[-] Cannot Find ADLS Training File. Try Querying First."
    }
}

# IMPORT CSV INTO FINAL QUERY RESULT
function import_csv () {
    if (Test-Path "$PWD/$records_folder/ADLS_Training.csv") {
        Write-Host "[+] Found ADLS Training File"
        $final_cont = Get-Content "$PWD/$records_folder/ADLS_Training.csv"
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
        $headers | Out-File "$PWD/$records_folder/ADLS_Training.csv" -Encoding ascii
        $table | Sort-Object | Out-File "$PWD/$records_folder/ADLS_Training.csv" -Encoding ascii -Append

    } else {
        Write-Host "`n[-] Cannot Find ADLS Training File.  Try Querying First"
    }

}

# RESTORE BACKUP FUNCTION
function restore_back () {
    if ((Test-Path "$PWD/$records_folder/ADLS_Training.csv") -and (Test-Path "$PWD/$records_folder/ADLS_Training_BK.csv")) {
        Write-Host "[+] Found ADLS Training File & Backup... Restoring"
        Copy-Item "$PWD/$records_folder/ADLS_Training_BK.csv" "$PWD/$records_folder/ADLS_Training.csv" -Force
        Write-Host "`n[+] Backup Restored...`n"
    } else {
        Write-Host "[-] Cannot Find ADLS Training File or Backup.  Try Querying First"
    }
}

# CREATE SAMPLE CSV TO IMPORT
function create_sample () {
    if (Test-Path "$PWD/$records_folder/ADLS_Training.csv") {
        Write-Host "[+] Found ADLS Training File"
        $adls_cont = Import-Csv "$PWD/$records_folder/ADLS_Training.csv"
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

# DOWNLOAD CERTIFICATE FOR INDIVIDUAL USER
function user_certificate() {
    Write-Host "[+] Download Certificate for User"
    if (Test-Path "$PWD/$records_folder/ADLS_Training.csv") {
        Write-Host "[+] Found ADLS Training File"

        $links = login_func

        # CHOOSE USER
        $adls_cont = Import-Csv "$PWD/$records_folder/ADLS_Training.csv"
        Write-Host "`n[+] Which user would you like to check?"
        $counter = 1
        foreach ($item in $adls_cont) {
            Write-Host "`t[$counter] " $item.User 
            $counter += 1
        }

        # GET USER FIRST/LAST NAME
        $chk_user = Read-Host "`n[-] Input Selection"
        $user_vars = ($adls_cont[$chk_user-1].User).Split(" ");$first_name = $user_vars[1];$last_name = $user_vars[0]

        # PROCESS CERT REQUEST
        $webreq = New-Object SmartCard.webpage
        $webreq.url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp"
        $webreq.referer = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp"
        $webreq.method = "POST"
        $webreq.accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        $webreq.data = "FormSubmitted=Submitted&ADLS=&strFname=$first_name&strLname=$last_name&strSSN="
        $webreq.contenttype = "application/x-www-form-urlencoded"
        $chk_resp = [SmartCard.mainclass]::web_req($webreq)
        
        #$org_name_search = $org_name.replace("+"," ")
        #$url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp"
        #$post_params = "FormSubmitted=Submitted&ADLS=&strFname=$first_name&strLname=$last_name&strSSN="
        #$chk_resp = (process_req $url $url "POST" $post_params)

        # FEW LINES BELOW CREATE ARRAY OF CODE TO GRAB INFORMATION SUCH AS USERIDENT AND DETAILS ON USER IN SQUADRON
        $my_temp_array = @()
        $org_name_temp = $org_name.Replace("+"," ") # CORRECT THE ORGANIZATIONAL NAME
        ($chk_resp -split "`n") | %{if ($_ -like "*>$org_name_temp<*" -or $_ -like "*ShowDetails*"){$my_temp_array += $_}} # CREATE ARRAY TO PARSE
        $show_dets = $my_temp_array | %{if($_ -like "*>$org_name_temp<*"){($my_temp_array[([int][array]::IndexOf($my_temp_array,$_)-1)] -split ";")[1] -split "'"}} # GET DETAIL INFORMATION FOR USER
        $COUNTERVAR = $show_dets[1] # NUMBER OF SEARCH RESULT
        $USER_IDENT = $show_dets[3] # USER IDENTITY OF PERSON IN SQUADRON
        
        # PROCESS REQUEST
        $url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserRemoval_UserProgress_AJAX.asp?ADLS=&USERS_IDENT=$USER_IDENT&CounterVar=$COUNTERVAR&sid=" + (Get-Random -Minimum 1 -Maximum 10)
        
        $webreq = New-Object SmartCard.webpage
        $webreq.url = $url
        $webreq.referer = $url
        $webreq.method = "GET"
        $webreq.accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        $webreq.data = ""
        $webreq.contenttype = "application/x-www-form-urlencoded"
        $cert_response = [SmartCard.mainclass]::web_req($webreq)

        #$cert_response = process_req $url $url "GET"

        # CREATE ARRAY OF PAGE OBJECTS WORKING WITH
        $page_array = @()
        $cert_response -split "`n" | %{if (($_ -like "*width:390px*") -or ($_ -like "*certificate*")){$page_array += $_}}
        
        # CREATE CERTIFICATE LINK ARRAY
        $cert_links = @()
        $cert_response -split "`n" | %{if ($_ -like "*certificate*"){(($_ -split '"')[5]) -match ".*strCourseID=(......).*" | Out-Null; $cert_links += $Matches[1] }}

        # CREATE COURSE NAME ARRAY
        $cert_name = @()
        $cert_response -split "`n" | %{if($_ -like "*certificate*"){$cert_name += (($page_array[([array]::IndexOf($page_array,$_)+1)]) -split ">")[1] -replace "</div",""  }}

        # CHOOSE WHICH CERTIFICATE TO DOWNLOAD & PRODUCE MENU
        Write-Host "`n[+] Which Certificate Would You Like?"
        $c_num = 1
        $cert_name | %{Write-Host "`t[$c_num] $_";$c_num++}
        $menu_sel = [int](Read-Host "`n[-] Input Selection")

        # ERROR CHECKING ON MENU SELECTION
        if ($menu_sel -le $cert_links.Count) {
            Write-Host "[-] Selection is $menu_sel"
        } else {
            Write-Host "[*] Error With Selection"
            break :menu
        }

        # CERTIFICATE REQUEST/RESPONSE
        $sel_id = $cert_links[$menu_sel-1]

        download_cert $sel_id $USER_IDENT $last_name

    } else {
        Write-Host "[-] Cannot Find ADLS Training File.  Try Querying First"
    }

}

# DOWNLOAD IMAGES FOR CERTIFICATE TO DYNAMICALLY CREATE
function download_images ($image_array) {
    $cur_dir = Get-Location
	if (!(Test-Path "$cur_dir/Certificates")){
		Write-Host "`n[+] Creating Folder 'Certificates' to Store Images and Certificates"
		New-Item "Certificates" -Force -ItemType directory | Out-Null
	}
	if (!(Test-Path "$cur_dir/Certificates/images")) {
		New-Item "$cur_dir/Certificates/images" -Force -ItemType directory | Out-Null
	}

    Write-Host "`t[+] Downloading images to dynamically create Certificate"

	foreach ($image in $image_array) {
		$url = "https://golearn.adls.af.mil/kc/certificate/$image"
		$wc = New-Object System.Net.WebClient
		$wc.DownloadFile($url, "$cur_dir/Certificates/images/$image")
	}
}

# DOWNLOAD CERTIFICATES FOR SQUADRON FOR SPECIFIC COURSE
function cert_coursenum ($links) {
    Write-Host "[+] Creating List Of Course Numbers"
    $cur_dir = Get-Location
    if (Test-Path "$PWD/$records_folder/ADLS_Training.csv") {
        Write-Host "[+] Found ADLS Training File"
        $adls_cont = Import-Csv "$PWD/$records_folder/ADLS_Training.csv"
        Write-Host "[+] Which Course Would you like to Download All user Certificates For?"

        $counter = 1
        foreach ($course in $links.Keys) {
            Write-Host "`t[$counter] $course"
            $counter++
        }

        $menu_sel = ([int](Read-Host "[-] Input Selection")-1)

        $c=0; $links.GetEnumerator() | %{if($c -eq $menu_sel){$my_link = $_.Value;$my_cs = $_.Key};$c++} # GET LINK OF SELECTION
        Write-Host "[+] Going to Parse For $my_cs Certificates"

        $my_link -match '.*CourseID=(.....[C0123456789])&.*' | Out-Null
        $link_cn = $Matches[1] # COURSE NUMBER

        $links = login_func

        $adls_cont = Import-Csv "$PWD/$records_folder/ADLS_Training.csv"
        $url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp"
        $ref = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp?ADLS="

        foreach ($item in $adls_cont) {
            $un = $item.User -split " " # USERNAME LAST,FIRST ARRAY
            $first_name = $un[1]
            $last_name = $un[0]
            $post_params = "FormSubmitted=Submitted&ADLS=&strFname=$first_name&strLname=$last_name&strSSN="

            Write-Host "`t[+] Searching for User $first_name $last_name" 
            
            $webreq = New-Object SmartCard.webpage
            $webreq.url = $url
            $webreq.referer = $ref
            $webreq.method = "POST"
            $webreq.accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
            $webreq.data = $post_params
            $webreq.contenttype = "application/x-www-form-urlencoded"
            $chk_resp = ([SmartCard.mainclass]::web_req($webreq)) -split "<tr"
            
            
            #$chk_resp = (process_req $url $ref "POST" $post_params) -split "<tr"


            # PROCESS CERT REQUEST
            $org_name_search = $org_name.replace("+"," ")
            $url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/UTMUserLookup.asp"
            $post_params = "FormSubmitted=Submitted&ADLS=&strFname=$first_name&strLname=$last_name&strSSN="
            
            $webreq = New-Object SmartCard.webpage
            $webreq.url = $url
            $webreq.referer = $url
            $webreq.method = "POST"
            $webreq.accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
            $webreq.data = $post_params
            $webreq.contenttype = "application/x-www-form-urlencoded"
            $chk_resp = ([SmartCard.mainclass]::web_req($webreq))
            
            
            # $chk_resp = (process_req $url $url "POST" $post_params)

            # FEW LINES BELOW CREATE ARRAY OF CODE TO GRAB INFORMATION SUCH AS USERIDENT AND DETAILS ON USER IN SQUADRON
            $my_temp_array = @()
            $org_name_temp = $org_name.Replace("+"," ") # CORRECT THE ORGANIZATIONAL NAME
            ($chk_resp -split "`n") | %{if ($_ -like "*>$org_name_temp<*" -or $_ -like "*ShowDetails*"){$my_temp_array += $_}} # CREATE ARRAY TO PARSE
            $show_dets = $my_temp_array | %{if($_ -like "*>$org_name_temp<*"){($my_temp_array[([int][array]::IndexOf($my_temp_array,$_)-1)] -split ";")[1] -split "'"}} # GET DETAIL INFORMATION FOR USER
            $COUNTERVAR = $show_dets[1] # NUMBER OF SEARCH RESULT
            $USER_IDENT = $show_dets[3] # USER IDENTITY OF PERSON IN SQUADRON

            download_cert $link_cn $USER_IDENT $last_name
        }
    } else {
        Write-Host "[-] Cannot Find ADLS Training File.  Try Querying First"
    }
}

# DOWNLOAD CERTIFICATE FUNCTION
function download_cert($sel_id,$USER_IDENT,$last_name){
    $ref = "https://golearn.adls.af.mil/kc/certificate/aetc_cert_certif.asp"
    #$cert_response = process_req "https://golearn.adls.af.mil/kc/certificate/aetc_cert_certif.asp?crs_ident=$sel_id&kc_ident=kc0001&login=$USER_IDENT" $ref GET

    $webreq = New-Object SmartCard.webpage
    $webreq.url = "https://golearn.adls.af.mil/kc/certificate/aetc_cert_certif.asp?crs_ident=$sel_id&kc_ident=kc0001&login=$USER_IDENT"
    $webreq.referer = $ref
    $webreq.method = "GET"
    $webreq.accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    $webreq.data = ""
    $webreq.contenttype = "application/x-www-form-urlencoded"
    $cert_response = [SmartCard.mainclass]::web_req($webreq)





    # DOWNLOAD IMAGES OF CERTIFICATE
    $image_array = @()
    ($cert_response -split "`n")| %{if ($_ -match '.*="([0-9A-Za-z_.]*(jpg|gif))".*') {if ($image_array -notcontains $matches[1]){$image_array += $matches[1]}}}
    $new_ia = $image_array -ne "AETC_CERT.jpg"
    download_images $new_ia

    # TEST FOR COMPLETION
    $complete = 0
    $month_array = @('January','February','March','April','May','June','July','August','September','October','November','December')
    ($cert_response -split "`n") | %{foreach($month in $month_array){if (($_ -like "*$month*") -and ($_ -notlike "*months*")){$complete = 1}}}
    if ($complete -eq 0){Write-Host "[*] Course $sel_id not complete for $last_name"; return}

    # FINAL FILE CREATION / TEST-PATH
    $final_out = ""
    ($cert_response -split "`n") | %{foreach ($image in $image_array){if ($_ -like "*$image*"){$_ = $_.Replace($image,"./images/$image")}}$final_out += $_} 
    Write-Host "[+] Outputting Certificate for $last_name $sel_id"
    $final_out | Out-File "./Certificates/$last_name $sel_id.html" -Force
}

function edit_courses() {
    :edit_menu while(1) {
        Write-Host "`n[+] Edit Courses Tracked"
        Write-Host "`t [1] Add a Course to the Tracker"
        Write-Host "`t [2] Delete a Course from the Tracker"
        Write-Host "`t [3] View Currently Tracked Courses"
        Write-Host "`n`t [4] Return To Main Menu`n"
        $menu_sel = Read-Host "[-] Input Selection"

        switch($menu_sel) {
            1 {add_course}
            2 {remove_course}
            3 {view_courses}
            4 {Write-Host "[+] Returning to Main Menu"; return}
            default {Write-Host "[*] Incorrect Selection"; break :edit_menu}
        }
    }
}

function add_course() {
    Clear-Host
    Write-Host "[+] Search for the Course you are looking to add"
    $search = Read-Host "`t[-] Input Search Term"
    
    # URL REQUEST TO CREATE ARRAY OF COURSES
    $url = "https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/OrgGroupReport/OrgGroupReportByCourse.asp"
    
    $webreq = New-Object SmartCard.webpage
    $webreq.url = $url
    $webreq.referer = $url
    $webreq.method = "GET"
    $webreq.accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    $webreq.data = ""
    $webreq.contenttype = "application/x-www-form-urlencoded"
    $pg_results = [SmartCard.mainclass]::web_req($webreq)
    
    #$pg_results = process_req $url $url "GET"
    $temp_array = @()
    ($pg_results -split "`n") | %{if (($_ -like "*$search*") -and ($_ -like "*option value*")) {$temp_array += $_}}

    # DISPLAY MENU OF ALL COURSE RESPONSES
    $c = 1
    $course_array = @()
    Write-Host "`n[+] Which Course Would You Like To Add"
    foreach ($item in $temp_array) {
        $item -match ".*>(.*)<.*" | Out-Null
        $cn = ($matches[1]) -replace $unsafe_chars,""
        #$cn -match "(.*)\(.*\).*" | Out-Null
        #$cn = $matches[1] # STRIP EVERYTHING BEFORE PARANTHESIS
        $item -match '.*value="(.*)">.*' | Out-Null
        $cs_num = $matches[1]
        $line = $cn + "=" + $cs_num
        $course_array += $line
        Write-Host "`t[$c] $cn : $cs_num"
        $c++
    }
    Write-Host "`n`t[$c] Exit"
    
    # GET SELECTION OF COURSE TO ADD
    $cs_to_add = [int](Read-Host "[-] Input Selection")
    if ($cs_to_add -gt $temp_array.Count) {Write-Host "[+] Exiting";return}

    # ADD TO COURSE FILE/CREATE IF NOT CREATED
    if (!(Test-Path "$pwd/courses_tracked.dat")) {
        Write-Host "[+] Creating Courses_Tracked.dat"
        New-Item -ItemType file -Name "courses_tracked.dat" | Out-Null
    }

    $exists = 0
    $info = ($course_array[[int]$cs_to_add-1]) -replace "\*",""
    $dat = Get-Content "$PWD/courses_tracked.dat"
    ($dat -split "`n") | %{if ($_ -like "*$info*"){Write-Host "[*] Duplicate Course Found`n`t[*] Unable to Add";$exists = 1} }
    if ($exists -ne 1) {
        Write-Host "[+] Adding $info to courses_tracked.dat"
        $info | Out-File -Append -Force -FilePath "$pwd/courses_tracked.dat"
    }
    return
}

function remove_course() {
    Clear-Host
    Write-Host "[-] Remove a Course"

    # TEST FOR FILE
    if (!(Test-Path "$pwd/courses_tracked.dat")) {
        Write-Host "[+] Courses_tracked.txt Not Found"
        return
    }

    $course_file = Get-Content "$pwd/courses_tracked.dat"
    Write-Host "[+] Which Course would you like to remove"
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
    Write-Host "`n`t[$c] Exit`n"
    $menu_sel = [int](Read-Host "[-] Input Selection")

    if ($menu_sel -gt $course_file.count) {Write-Host "[*] Returning to Menu";return}

    $removed = 0
    $line_to_remove = $course_array[[int]$menu_sel-1]
    $new_array = @()
    $course_file | %{if (($_ -like "*$line_to_remove*") -and ($removed -eq 0)){Write-Host "[+] Removing $line_to_remove";$removed = 1}else {$new_array += $_}}
    $new_array | Out-File -Force -FilePath "$pwd/courses_tracked.dat"

    return
}

function view_courses() {
    # TEST FOR FILE
    if (!(Test-Path "$pwd/courses_tracked.dat")) {
        Write-Host "[+] Courses_tracked.dat Not Found"
        return
    }

    $course_file = Get-Content "$pwd/courses_tracked.dat"
    Write-Host "[+] Courses Being Tracked Are"
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
    Write-Host "`n`n"

    $links = get_course_urls
}

function get_course_urls() {

    # locate file mapping course names to ID
    if (!(Test-Path "$pwd/courses_tracked.dat")) { Write-Host "[+] Courses_tracked.dat Not Found"; return }
    $courses = Get-Content -Path "$pwd/courses_tracked.dat"

    # generate URLs for each tracked course
    $course_urls=@{}
    foreach ($course in $courses) {
        $course_name,$course_num = $course -split "="
        $course_urls.Add("$course_name","https://golearn.adls.af.mil/kc/doza.tools/functions/UTMUDM/OrgGroupReport/OrgGroupReportByCourse_Excel.asp?ReportName=SearchUsersDetails&strAction=SearchUsersDetails&customerorg_ident=$org_id&adls_org_id=&CourseID=$course_num&blnAllSubOrgs=true&strRankType=" + "'AD','CV','CT','DP','OT'")
    }

    return $course_urls
}

$smartcard = @"

using System;
using System.Net;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading;
using System.Security.Cryptography.X509Certificates;

namespace SmartCard
{
    public class mainclass
    {
        static CookieContainer cook_cont = new CookieContainer();
        static X509Certificate2 cert = null;

        public static string main_func()
        {
            Console.WriteLine("[+] ADLS Training Tracker\n");
            
            set_cookies();
            cert = pick_cert();
            if (cert == null)
            {
                return "Error";
            }

            return "Success";
        }

        public static string web_req(webpage mypage)
        {
            HttpWebRequest req = (HttpWebRequest) WebRequest.Create(mypage.url);
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:51.0) Gecko/20100101 Firefox/51.0";
            req.ContentType = mypage.contenttype;
            if (req.Accept != null)
            {
                req.Accept = mypage.accept;
            }
            req.Proxy.Credentials = CredentialCache.DefaultCredentials;
            req.CookieContainer = cook_cont;
            req.Headers.Add("X-Requested-With", "XMLHttpRequest");
            req.ClientCertificates.Add(cert);
            req.Method = mypage.method;
            req.Referer = mypage.referer;
            req.ServicePoint.Expect100Continue = false;

            if (mypage.method == "POST")
            {
                byte[] byteArray = Encoding.UTF8.GetBytes(mypage.data);
                req.ContentLength = byteArray.Length;
                Stream dataStream = req.GetRequestStream();  
                dataStream.Write (byteArray, 0, byteArray.Length); 
                dataStream.Close(); 
            }


            HttpWebResponse resp = (HttpWebResponse) req.GetResponse();
            Stream datastream = resp.GetResponseStream();
            StreamReader reader = new StreamReader(datastream);
            
            string full_resp = reader.ReadToEnd();

            datastream.Dispose();
            reader.Dispose();

            return full_resp;

        }

        private static void set_cookies()
        {
            Console.WriteLine("[+] Adding Cookies");

            Cookie cook = new Cookie();
            cook.Name = "AFPORTAL_LOGIN_AGREEMENT";
            cook.Value = "accepted";
            cook.Domain = "af.mil";
            cook_cont.Add(cook);

            cook = new Cookie();
            cook.Name = "BUILDING";
            cook.Value = "Administration";
            cook.Domain = "golearn.adls.af.mil";
            cook_cont.Add(cook);
       
        }

        private static X509Certificate2 pick_cert()
        {
            Assembly syssec_asm = Assembly.LoadFrom(@"C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Security.dll");
            Type syssec_type = syssec_asm.GetType("System.Security.Cryptography.X509Certificates.X509Certificate2UI");
            Type selection_flag = syssec_asm.GetType("System.Security.Cryptography.X509Certificates.X509SelectionFlag");
            var selection_value = selection_flag.GetField("SingleSelection").GetValue(null);

            var methodinfo = syssec_type.GetMethod("SelectFromCollection", new Type[] { typeof(X509Certificate2Collection), typeof(string), typeof(string), selection_value.GetType() });

            X509Store store = new X509Store("MY", StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
            X509Certificate2 cert = null;

            try
            {
                object[] invparams = new object[4];
                invparams[0] = store.Certificates;
                invparams[1] = "Digital Certificates";
                invparams[2] = "Select a certificate from the following list:";
                invparams[3] = selection_value;

                X509Certificate2Collection cert_col = (X509Certificate2Collection) methodinfo.Invoke(null, invparams);
                cert = cert_col[0];
            }
            catch
            {
                return null;
            }

            return cert;
        }
    }

    public class webpage
    {
        public string url;
        public string referer;
        public string contenttype;
        public string accept;
        public string method;
        public string data;
    }

}
"@

######################
# MAIN PROGRAM START #
######################

# set cookies and select certificate
#Add-Type -TypeDefinition $smartcard -IgnoreWarnings -Language CSharpVersion3
#$status = [SmartCard.MainClass]::main_func()
#if ($status -like "Error") { Write-Host "[+] CAC Certificate Error"; pause; break }

$links = login_func

:menu while (1) {
    Write-Host "`n[+] ADLS Menu Options`n"
    Write-Host "`t[1] Query ADLS for CSV"
    Write-Host "`t[2] Check CBT Dates (using local CSV file)"
    Write-Host "`t[3] Get Certificate For User"
    Write-Host "`t[4] Get Certificates For Course"
    Write-Host "`t[5] Import CSV"
    Write-Host "`t[6] Create CSV Template to Import"
    Write-Host "`t[7] Restore Backup"
    Write-Host "`t[8] Edit Courses Tracked"
    Write-Host "`n`t[9] Exit`n`n"
    $menu_input = Read-Host "[-] Input Selection"

    switch ($menu_input) {
        1 {Clear-Host;query_cbt}
        2 {Clear-Host;check_cbt}
        3 {Clear-Host;user_certificate}
        4 {Clear-Host;cert_coursenum $links}
        5 {Clear-Host;import_csv}
        6 {Clear-Host;create_sample}
        7 {Clear-Host;restore_back}
        8 {Clear-Host;edit_courses}
        9 {exit}
        default {Write-Host "[*] Error with Selection"; exit}
    }
}
