
function get_assertion() {


    try {
        $assertion_response = Invoke-RestMethod -Method 'Post'  -Uri $Global:config.assertion_url -ContentType $Global:config.content_type -Body @{'client_id' = $Global:config.client_id; 'user_id' = $Global:config.user_id; 'token_url' = $Global:config.token_url; 'private_key' = $Global:config.private_key }
        return $assertion_response

    }

    catch {

        [int]$StatusCode = $_.Exception.Response.StatusCode

        $errorMessage = $_.Exception.Message
        
        $StatusCode
        $errorMessage

    }
 
}
 
 
function get_token {
    param($assertion)   

   
    try {
 
        $token = Invoke-RestMethod -Method 'Post' -Uri $Global:config.token_url -ContentType $Global:config.content_type -Body @{'company_id' = $Global:config.company_id; 'client_id' = $Global:config.client_id; 'grant_type' = $Global:config.grant_type; 'user_id' = $Global:config.user_id; 'assertion' = $assertion }

        return $token

    }

    catch {

        [int]$StatusCode = $_.Exception.Response.StatusCode

        $errorMessage = $_.Exception.Message

        $StatusCode
        $errorMessage

    }

}
 
function remove_and_create_csv_file {

    param($source, $env)
 
    $filename_csv = $Global:config.file_path + $source + "_" + $env + ".csv"
 
    if (test-path $filename_csv) {

        Remove-Item $filename_csv

    }
 
    $result = New-Item -Path $filename_csv -ItemType File

    return $filename_csv
}
 
function get_data_from_url {
    [CmdletBinding()]
    param($token, $source)
    
    $headers = @{

        Authorization = "Bearer " + $token.access_token

    }

    $reahce = 'reached in get data initial : ' + $source

    $reahce
    #$token.access_token

    try {

        $v_url = $Global:config.($source)["url"]

        $val23 = ' MAIN URL : ' + $v_url 
        $val23

        $headers

        $iter = 0

        do {

            $iter

            $data = Invoke-RestMethod -Method 'Get' -Uri $v_url -Headers $headers

            $val24 = 'Data URL : ' + $data
            $val24

            $iter = $iter + 1

            $iter
 
            #$csv_data = convertto-json $data

            #$csv_Data

            parse_and_extract_csv -p_url_data $data -source $source

            $v_url = $data.d.__next

            $v_url

        } while ($data.d.__next)

        return $data

    }

    catch {

        [int]$StatusCode = $_.Exception.Response.StatusCode

        $errorMessage = $_.Exception.Message

        $StatusCode
        $errorMessage

    }

}
 

function get_url_data {
    
    param($url, $token)

    $headers = @{

        Authorization = "Bearer " + $token.access_token

    }

    #$reahce = "reached in get url data function"

    #$reahce

    try {

        $v_url = $url

        $v_url

        $data = Invoke-RestMethod -Method 'Get' -Uri $v_url -Headers $headers

        return $data

    }

    catch {

        [int]$StatusCode = $_.Exception.Response.StatusCode

        $errorMessage = $_.Exception.Message

        $StatusCode
        $errorMessage

    }

}
function insert_object() {


    #$req_arr_new= @()

    $Flag = if ($null -eq $dateofJoining -or $dateofJoining -eq '') { 'R' }
    $MonthN = if ($Flag -ne 'R') { (Get-Date -Format 'MM' $dateofJoining) }
    
    $YearN = if ($Flag -ne 'R') { Get-Date -Format 'yyyy' $dateofJoining }
    
		
    $FlagTJD = if ($null -eq $TentativeDateOfJoining -or $TentativeDateOfJoining -eq '') { 'S' }
    $MonthNTJD = if ($FlagTJD -ne 'S') { (Get-Date -Format 'MM' $TentativeDateOfJoining) }
                
    $YearNTJD = if ($FlagTJD -ne 'S') { Get-Date -Format 'yyyy' $TentativeDateOfJoining }
 
    foreach ($just in $result.Justification.results) {

        $justification = $just.externalCode

        #$status = $just.status

    }

    $Object = New-Object PSObject -Property @{

        HRAR                   = $HRAR

        Req_id                 = $Req_id

        candidate_id           = $candidate_id

        candidate_name         = $first_name + ' ' + $last_name

        Department             = $Department

        Recruiter              = if ($recuriter -like '*itreport*' -or $recuriter.tolower() -like '*purged*') { 'Recruiter Exit' } else { $recuriter }

        Date_of_Offer          = $Offerdate

        updated_Grade          = if ($Grade -like '*T*') { 'Freshers' }elseif ($Grade -eq 'UC') { 'Contractor' } else { 'Lateral' }
				
        Type_Of_Hire           = if ($typeofHire -eq 'vtypeofHire_FTE') { 'Full Time' } elseif ($typeofHire -eq 'vtypeofHire_Contract') { 'Contract' }else { '' }
				
        Date_Of_Joining        = $dateofJoining
				
        Source                 = if ($Source_e -like '*Referral*') { 'Referral' } elseif ($Source_e -like '*Campus*') { 'Campus' } elseif ( $Source_e -like '*Naukri*' -or $Source_e -like '*LinkedIn*' -or $Source_e -like '*Email*' -or $Source_e -like '*Indeed*' -or $Source_e -like '*Monster*' -or $Source_e -like '*Walkin*' -or $Source_e -like '*Website*' -or $Source_e -like '*Portal*' -or $Source_e -like '*Facebook*') { 'Direct' }else { 'Vendor' }

        Designation            = $Designation

        #Skill
        Year                   = if ($Flag -ne 'R') { $YearN } elseif ($FlagTJD -ne 'S') { $YearNTJD }else { '' }
                
        Month                  = if ($Flag -ne 'R') { (Get-Date -Format 'MM' $dateofJoining) } elseif ($FlagTJD -ne 'S') { (Get-Date -Format 'MM' $TentativeDateOfJoining) }else { '' }
				
        Financial_Year         = if ($Flag -ne 'R') { if ($MonthN -lt 4) { [string]($YearN - 1) + '-' + [string]$YearN } else { [string]$YearN + '-' + [string]($YearN + 1) } }elseif ($FlagTJD -ne 'S') { if ($MonthNTJD -lt 4) { [string]($YearNTJD - 1) + '-' + [string]$YearNTJD } else { [string]$YearNTJD + '-' + [string]($YearNTJD + 1) } }else { '' }
				
        Gender                 = if ($gender -eq '' -or $null -eq $gender ) { $gender } else { $gender.Substring(7) }

        Grade                  = $Grade

        Justification          = $justification
				
        Billing_Category       = if ($null -ne $BCategory1 -and $justification -eq 'Fresher') { 'FRESHER' } else { $BCategory1 }
                
        Work_Location          = if ($jobCity.tolower() -like '*ben*' -or $jobCity.tolower() -like '*ban*' -or $jobCity.tolower() -eq 'karnataka') { 'Bengaluru' } else { $jobCity }

        HRAR_Creation_Date     = $HRAR_Creation_Date
				
        Tower                  = if ($JobCode -like '*DEL*' -or $JobCode -like '*IT*' -or $Department -like '*IT*' -or $Department -like '*CEO*') { 'Technology' }elseif ($JobCode -like '*OPS*' -or $Department -like '*RSS*' -or $Department -like '*BPM*') { 'Ops' }elseif ($Department -like '*EF*' -or $Department -like '*LND*' -or $JobCode -like '*ENB*') { 'Enablement' } else { $JobCode }
        #if ($null -ne $JobCode ) { if ($JobCode -like '*DEL*') { 'Technology' }elseif ($JobCode -like '*OPS*') { 'Ops' }elseif ($JobCode -like '*ENB*') { 'Enablement' } }elseif ($null -ne $Department) { if ($Department -like 'IT') { 'Technology' } }else { '' }
        Age                    = $Age

        # J_Status = $status

        
        Recruiter_userName     = $Recruiter_userName
        Tentative_Joining_Date = $TentativeDateOfJoining
        Status                 = if ($dateofJoining -gt (Get-Date).AddDays(-1) -or $TentativeDateOfJoining -gt (Get-Date).AddDays(-1)) { 'Yet to join' } else { $o_Status }
    }

    $req_arr_new += $Object
    #return $Object
            

    $itr = $itr + 1
		
    $output = $req_arr_new | Export-Csv $filename_csv -Append -NoTypeInformation -Encoding UTF8 
}

            
 
function parse_ta_dashboard {

    param($p_url_data, $source, $token)
 
    $val3 = 'inside parse_ta_dashboard function'
    $val3
    $req_arr_new = @()
    
    $val22 = 'p url data : ' + $p_url_data
    $val22
    

    foreach ($result in $p_url_data) {
        $val7 = 'inside foreach where p_url_data is parsed to result'
        $val7
        $val9 = 'Value stored in result var : ' + $result
        $val9

        $HRAR = ""

        $Req_id = ""

        $candidate_id = ""

        $first_name = ""

        $last_name = ""

        $Department = ""

        $recuriter = ""

        $Designation = ""
 
        $gender = ""

        $candidate_id = ""

        $justification = ""

        $jobCity = ""

        $HRAR_Creation_Date = ""

        $status = ""

        $BCategory = ""

        $Source_e = ""

        $dateofJoining = ""
                 
        $typeofHire = ""

        $Age = ""

        $JobCode = ""

        $Recruiter_userName = "" 

        $Offerdate = ""

        $TentativeDateOfJoining = ""


        $filename_csv = $Global:config.file_path + $source + "_" + $env + ".csv"
        
        #if ($result.Department -like '*Voya*'){ 
        #if ($result.jobReqId -eq '18112'){ 
        
        $HRAR = $result.hrar

        $Req_id = $result.jobReqId

        $Department = $result.Department

        $Grade = $result.customGrade
        $Grade


        $Age = $result.age
			
        $BCategory1 = $result.SPriority.results.Externalcode
        $HRAR_Creation_Date = $result.HRARCDate

        $Designation = $result.designation

        $jobCity = $result.jobCity
            
        $JobCode = $result.jobCode

        $justification = $result.Justification.results.externalCode
 
        $jobReqFwdUrl = $result.jobApplications.__deferred.uri + $Global:config.($source)["job_application_filter"]

        $val21 = 'Job req fwd url : ' + $jobReqFwdUrl
        $val21
            
			
			
			
        foreach ($recuriterVal in $result.recruiter.results) {

            $recuriter = $recuriterVal.firstName + ' ' + $recuriterVal.lastName
            $Recruiter_userName = $recuriterVal.userName
        }
            
           

            
 
        if ($null -ne $result.jobApplications.__deferred.uri -and $result.jobApplications.__deferred.uri -ne '') {
            $jobReqFwd = get_url_data -url $jobReqFwdUrl -token $token
        }
        $val12 = 'Response from get_url_data func : ' + $jobReqFwd
        $val12
         
         
        foreach ($candidate in $jobReqFwd.d.results) { 
            $val13 = 'Inside i' + $candidate
            $val13


            $numofofferletters = $candidate.offerLetter.results.status.Count
            if ($null -ne $candidate.Offerextended -or $candidate.offerLetter.results[$numofofferletters - 1].status -eq 'accepted' -or $null -ne $candidate.offerLetter.results[$numofofferletters - 1].offerSentDate) {
                
                #'$Offer sent date  : ' + $candidate.offerLetter.results[$numofofferletters - 1].offerSentDate

                $candidate_id = $candidate.candidateId

                $first_name = $candidate.firstName

                $last_name = $candidate.lastName -replace "'", ""
 
                $dateofJoining = $candidate.hiredOn

                $o_Status = if ($candidate.status -eq 'open' ) { 'Joined' }elseif ($candidate.status -eq 'Disqualified') { 'Declined' }elseif ($candidate.offerLetter.results.status -eq 'accepted') { 'Joined' } 

                $gender = $candidate.custGender.results.externalCode

                $typeofHire = $candidate.vtypeofHire.results.externalCode

                $Source_e = $candidate.candidateSource.results.externalCode

                
                $Offerdate = $candidate.offerLetter.results[$numofofferletters - 1].offerSentDate

                if ("" -eq $Offerdate -or $null -eq $Offerdate) {
                    $Offerdate = $candidate.Offerextended
                }
                $Flag = if ($null -eq $dateofJoining -or $dateofJoining -eq '') { 'R' }
                $MonthN = if ($Flag -ne 'R') { (Get-Date -Format 'MM' $dateofJoining) }
                
                $YearN = if ($Flag -ne 'R') { Get-Date -Format 'yyyy' $dateofJoining }
                

                
                
                
                $count = $candidate.jobOffer.results.Count
                #'Count of offer letters : '+$count
                $TentativeDateOfJoining = $candidate.jobOffer.results[$count - 1].customDate3
                #$TentativeDateOfJoining
                $FlagTJD = if ($null -eq $TentativeDateOfJoining -or $TentativeDateOfJoining -eq '') { 'S' }
                $MonthNTJD = if ($FlagTJD -ne 'S') { (Get-Date -Format 'MM' $TentativeDateOfJoining) }
                
                $YearNTJD = if ($FlagTJD -ne 'S') { Get-Date -Format 'yyyy' $TentativeDateOfJoining }

                if ($candidate.candidateId -ne "" -and $candidate.firstName -ne "" -and $candidate.lastName -ne "") {
                    insert_object 
                }         
            } 

        }

        #}
      

    }
    #$output = $req_arr_new | Export-Csv $filename_csv -Append -NoTypeInformation -Encoding UTF8 
    #remove-DuplicatesFromCSV -FilePath $filename_csv -UniqueFields @("HRAR", "Req_id", "candidate_id")
}
 
 
function parse_and_extract_csv {

    param($p_url_data, $source )
    #$req_arr_new= @()
 
    if ($source -eq "ta_dashboard") {

        $val = 'inside ta_dashboard flow  ' + $p_url_data

        $val

        parse_ta_dashboard -p_url_data $p_url_data.d.results -source $source -token $token

    }
 
    if ($source -eq "people_count") {

        $val = 'inside people count flow : ' + $p_url_data

        $val

        parse_people_count -p_url_data $p_url_data.d.results -source $source -env $env

        
 
        $val = 'outside people count flow'

        $val

    }


    if ($source -eq "exit_details") {

        $val = 'inside exit-details flow'

        $val

        parse_exit_details -p_url_data $p_url_data.d.results -source $source -env $env
 
        $val = 'outside exit-details flow'

        $val

    }

 
    

    #$filename_csv = "C:/MIS_Platform/"+$source+".csv"
 
    #$output = $req_arr_new | Export-Csv $filename_csv -NoTypeInformation -Encoding UTF8

}
 
function find_and_replace_empty_string_for_null ($str) {

    $str

    while ($str -like ",,") {

        $str = $str -replace ",,", ",null,"
 
        $str

    }
 
    return $str

}
 
function remove-DuplicatesFromCSV {
    param (
        [string]$FilePath,
        [string[]]$UniqueFields
    )

    if (-Not (Test-Path -Path $FilePath)) {
        Write-Host "Error : File not found in the given path - $Filepath"
        return
    }

    $csvData = Import-Csv -Path $FilePath

    if (-Not $csvData) {
        Write-Host "Error : CSV file $Filepath is empty or invalid"
        return
    }

    Write-Host "Removing duplicates based on fields : $($UniqueFields -join ', ')"

    $uniqueData = $csvData | Group-Object -Property $UniqueFields | ForEach-Object { $_.Group[0] }

    $uniqueData | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8

    Write-Host "Duplicates removed, updated file and saved the file to $FilePath "

    
}
 
function generate_insert_script_from_csv($source, $env) {

    $insert_script = "insert into <table> (<names>) values (<values>)"
 
    $fileName = $Global:config.file_path + $source.toString() + ".txt"

    $filename_csv = $Global:config.file_path + $source + "_" + $env + ".csv"
    $filename_csv
 
    if (test-path $fileName) {

        Remove-Item $fileName

    }
 
    $insert_script_new = $insert_script.replace("<table>", $Global:config.($source.toString())["target_table"] )
 
    

    $csv_data = Get-Content $filename_csv 

    #$csv_data

    $iterator = 0
 
    foreach ($line in $csv_data) {

        #$iterator 

        #$line

        if ($iterator -eq 0) {

            $names = $line -replace '"', ''

            $insert_script_new = $insert_script_new -replace "<names>", $names

        }

        if ($iterator -gt 0) {

            $values = $line -replace '"', ''''

            $values = $values -replace "''", "null"
 
            $charToFind = ',,'

            $index = $values.IndexOf($charToFind)

            while ($index -ge 0) {

                $values = $values -replace ",,", ",null,"                

                $index = $values.IndexOf($charToFind)

            }
 
            add-content $fileName $insert_script_new.replace("<values>", $values )

        }

        $iterator++

    }

}
 
$startTime = Get-Date
$assertion = get_assertion
#$assertion
$token = get_token -assertion $assertion
#$token


if ($source -eq 'hr_dashboard') {
    remove_and_create_csv_file -source $source -env $env
    'People - Count'
    remove_and_create_csv_file -source 'people_count' -env $env
 
    get_data_from_url -token $token -source 'people_count'
    
    parse_and_extract_csv -p_url_data $url_data.d.results -source 'people_count' 

    merge_manualentries -source 'people_count' -env $env
    
    generate_insert_script_from_csv -source 'people_count' -env $env

    'exit - details'
    remove_and_create_csv_file -source 'exit_details' -env $env
 
    get_data_from_url -token $token -source 'exit_details'
    
    parse_and_extract_csv -p_url_data $url_data.d.results -source 'exit_details' 
    
    generate_insert_script_from_csv -source 'exit_details' -env $env

    merge_reports -env $env -source $source

    generate_insert_script_from_csv -source $source -env $env
}
elseif ($source -eq 'ta_dashboard') {
    $filepath = remove_and_create_csv_file -source $source -env $env
 
    get_data_from_url -token $token -source $source

    parse_and_extract_csv -p_url_data $url_data.d.results -source $source 

    remove-DuplicatesFromCSV -FilePath $filepath -UniqueFields @("HRAR", "Req_id", "candidate_id", "candidate_name")

    generate_insert_script_from_csv -source $source -env $env
}
else {
    remove_and_create_csv_file -source $source -env $env
 
    get_data_from_url -token $token -source $source

    parse_and_extract_csv -p_url_data $url_data.d.results -source $source 

    generate_insert_script_from_csv -source $source -env $env
}
$endTime = Get-Date
$duration = $endTime - $startTime
Write-Host "Total duration of the process : "+$duration  -ForegroundColor Green
