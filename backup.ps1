
$SchoolTypes = [ordered] @{ 
    "gymnasieskola"    = "UpperSecondarySchool"
    "gymnasiesärskola" = "UpperSecondarySchoolForLearningDisabilities"
    #"grundskola"       = "CompulsorySchool"
    #"grundsärskola"    = "CompulsorySchoolForLearningDisabilities"

}
#endregion

#region Classes
class ElevExport {   
    [string]$EnhetID
    [string]$Enhetsnamn
    [string]$SkolenhetNamn
    [string]$SkolenhetID
    [string]$Skolenhetskod
    [string]$Förnamn
    [string]$Efternamn
    [string]$Personnummer
    [string]$Kön
    [string]$Klass
    [string]$AnsvarigRektor
    [string]$Utbildning
    [string]$UtbildningNamn
    [string]$Årskurs
    [string]$Skoltyp
    [string]$ElevGUID
    [string]$KlassGUID
    [string]$SkolenhetGUID

}

class SchoolExport {
    [string]$ID
    [string]$Name
    [string]$Type
    [string]$SchooGUID
    [string]$Principal
    [string]$principalName
}

class ProgramMandatoryExport {
   
    [string]$ProgramID
    [string]$ProgramName
    [string]$CourseID
    [string]$CourseName
    [string]$CourseCode
    [string]$CourseType
    [string]$CourseLevel

}

class EducationPlan {
    [string]$Name
    [string]$Type
    [string]$Program
    
    [System.Collections.Hashtable]$Courses = [System.Collections.Hashtable]::new()
}

class EducationPlanCourse {
    [string]$CourseId
    [string]$CourseType
    [string]$CourseTypeCode
    [string]$CourseTypePoints
    [string]$CourseName 
    [string]$CoursePoints
    [string]$CourseCode
    [string]$CourseSubjectCode
    [string]$CourseSubjectName
    [string]$CourseLevel
}

class EducationplanTypes {
    [string]$Name
    [string]$Type
    [string]$Program
    [string] $CourseType
}


class OutcomeMandatoryExport {
    [string]$PersonnummerElev
    [string]$GruppNamn
    [string]$ÄmneKurs
    [string]$Kursnamn
    [string]$Poäng
    [string]$Startdatum
    [string]$Slutdatum
    [string]$Betyg
    [string]$Akttyp
    [string]$Hur
    [string]$Period
    [string]$BetygsättandeLärare
    [string]$AllaUndervisandeLärare
    [string]$LåstBetyg
    [string]$AktivitetensÅrskurs
    [string]$AktivitetensGUID
    [string]$GruppGUID
}

class OfferingCourse {
    [string]$CourseId
    [string]$CourseName
    [string]$CoursePoints
    [string]$CourseCode
    [string]$CourseSubjectCode
    [string]$CourseSubjectName
    [string]$CourseLevel
}
 
class OfferingSubject {
    [string]$SubjectId
    [string]$SubjectName
    [string]$SubjectCode
 
}

class StudentGrades {
    [string]$Id
    [string]$Name
    [array]$CourseData
    [array]$CourseTypeCode
    [array]$GradeOutcome
    [string]$StudentId
    [string]$StudentName
}

class ElementaryStudentGrades {
    [string]$StudentID
    [string]$StudentName
    [string]$SchoolName
    [string]$UnitID
    [string]$SubjectCode
    [string]$SubjectName
    [string]$Date
    [string]$SemesterType
    [string]$SemesterYear
    [string]$Grade
    [bool]$FinalGrade
    [bool]$TrialPerformed
    [PSCustomObject]$CourseData
    [PSCustomObject]$Courses
}

function Read-HistoricEnterpriseXML($historicEnterprise = $ContentOrganizations.enterprise) {
    foreach ($SchoolType in $SchoolTypes.Keys) {
        Write-Host "Processing data for school type: $SchoolType"
        $xmlSchooltype = $historicEnterprise.properties.$SchoolType
        Write-Host "XML data for $SchoolType $xmlSchooltype"
        for ($i = 0; $i -le 9; $i++) {
            $year = (Get-Date).AddYears(-$i).ToString("yyyy")
            Write-Host "Processing data for year: $year"
            $OutFile = "$WorkingFolder\Organization_Historic_$SchoolType.xml"
            $date = (Get-Date).AddYears(-$i)
            Write-Host "Retrieving data from API for year: $year"
            try {
                $APIresult = Invoke-RestMethod "$($APIBaseURI)/$OrganizationVersion/Get$($SchoolTypes[$SchoolType])Organization?LicenseKey=$($APILicenceKey)&searchDate=$date" -OutFile $OutFile -TimeoutSec 6000 -Certificate $Cert -ErrorAction Stop
                Write-Host "Data retrieval for year $year completed."
                $ContentOrganizations = [xml]::new()
                $ContentOrganizations.Load($OutFile)
                Write-Host "XML data loaded for year: $year"
                Process-MembershipsAndGroups $ContentOrganizations
            }
            catch {
                Write-Host "Error retrieving data for year $year $_"
                continue
            }
        }
    }
}


function Process-MembershipsAndGroups($ContentOrganizations) {
    foreach ($membership in $ContentOrganizations.enterprise.membership) {
        foreach ($member in $membership.member | Where-Object { $_.role.roletype -eq "Student" -and $_.sourcedid.id -notlike "CG{*" -and $_.role.extension.placement }) {
            $memberId = $member.sourcedid.id
            $membershipId = $membership.sourcedid.id  # Get the membership ID
            
            Add-Member -InputObject $member -NotePropertyName 'MembershipID' -NotePropertyValue $membershipId -Force
    
            # Write to host if MembershipID is added
            if ($member | Get-Member -Name 'MembershipID' -ErrorAction SilentlyContinue) {
                # idg Write-Host "MembershipID $( $membershipId) added to $($member.sourcedid.id)"
            }
          
            try {
                $global:HistoricEntmembers.Add("$($memberId.ToString())", $member)
            }
            catch {
                continue
            }
        }
    }
    if ($ContentOrganizations.enterprise) {
        # Grabs PID (personnummer)
        foreach ($p in $ContentOrganizations.enterprise.person) {
            # Check if <userid> element exists and has the correct useridtype attribute
            $pdValue = $p.userid | Where-Object { $_.useridtype -eq 'PID' } | ForEach-Object { $_.'#text' }
    
            if ($pdValue) {
                # Add the PD value to the dictionary
                $global:HistoricEntPersonsPiD.Add($pdValue, $p)
            }
        }
    }
    

   
    if ($ContentOrganizations.enterprise.group) {
        foreach ($groups in $ContentOrganizations.enterprise.group -notlike "CG{*") {
            $groupId = $groups.sourcedid.id
            $groupType = $groups.grouptype.typevalue.'#text'
            $schoolYear = $groups.extension.schoolyear.'#text'
    
            if ($groupId -and $groupType -eq "Class" -and $schoolYear) {
                if (-not $global:HistoricGroups.ContainsKey($groupId)) {
                    $global:HistoricGroups.Add($groupId, $groups)
                }
            }
        }
    }
    
}

function GetSchoolData {
   
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $WorkingFolder = "C:\Users\AndreasOlofsson\Documents\TeisProjekt\Kholm_Revisioned"
    $CertPath = ".\cert\2022-wildcard-katrineholm.se-PFX.pfx"
    $CertPassword = "LtRlvuLliyUbzUmskw23TQ"
    
    Set-Location $WorkingFolder
    
    $ExcludePrivacy = $true
    
    $cert = Get-PfxCertificate $CertPath -Password $( ConvertTo-SecureString $CertPassword –AsPlainText -Force )
    
    
    $identityServerUrl = "https://education.service.tieto.com/WE.Education.IdentityServer.Web"
    $clientId = "e46b407b-42b5-4d12-960c-3b17eb27a815"
    $clientSecret = "8dtmzNTJF_fhU4i3L4K-NPyUxIvKn77TrixSWtX9O6S8CSTXNIyjL4qynuaQ92wpqA"
    $scope = ""
    
    $tokenEndpoint = "$identityServerUrl/connect/token"
    
    $tokenRequestBody = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
    }
    
    $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method POST -Body $tokenRequestBody
    
    $jwtToken = $tokenResponse.access_token
    
    # Use the $jwtToken for further authentication or authorization
    $baseRequestUrl = "https://prodintegration-education.service.tieto.com/WE.Education.IntegrationPortal.ExternalApi/api/v1"
    
    
    # https://api.ist.com/ss12000v2-api/
    
    
    $headers = @{
        "Authorization" = "Bearer $jwtToken"
    }
    
    
    <#  $organisations = Invoke-RestMethod -Uri "$baseRequestUrl/organisations" -Method Get -Certificate $cert -Headers $headers | Select-Object -ExpandProperty data 
    $organsationDictionary = [System.Collections.Hashtable]::new()
    $organisations | ForEach-Object { $organsationDictionary.Add($_.id, $_) }
    # $organisations | where {$_.municipalitycode -eq "0483"}
     #>
    $persons = Invoke-RestMethod -Uri "$baseRequestUrl/persons" -Method Get -Certificate $cert -Headers $headers | Select-Object -ExpandProperty data
    $personDictionary = [System.Collections.Hashtable]::new()
    $persons | ForEach-Object { $personDictionary.Add($_.id, $_) }
    
    
    <#   $duties = Invoke-RestMethod -Uri "$baseRequestUrl/duties" -Method Get -Certificate $cert -Headers $headers | Select-Object -ExpandProperty data
    $dutyDictionary = [System.Collections.Hashtable]::new()
    # $duties | ForEach-Object { $dutyDictionary.Add($_.id, $_) }
    
    $activities = Invoke-RestMethod -Uri "$baseRequestUrl/activities" -Method Get -Certificate $cert -Headers $headers | Select-Object -ExpandProperty data
    $activityDictionary = [System.Collections.Hashtable]::new()
    # $activities | ForEach-Object { $activityDictionary.Add($_.id, $_) }
     #>
    
    $groups = Invoke-RestMethod -Uri "$baseRequestUrl/groups" -Method Get -Certificate $cert -Headers $headers | Select-Object -ExpandProperty data
    $groupDictionary = [System.Collections.Hashtable]::new()
    $groups | ForEach-Object { $groupDictionary.Add($_.id, $_) }
    $groups | Select-Object -Property groupType -Unique




    $civicNoDisplaynameDictionary = [System.Collections.Hashtable]::new()

    foreach ($group in $groups) {
        
        $groupMemberships = $group.groupMemberships

        foreach ($membership in $groupMemberships) {
            $personId = $membership.person.id
            $displayName = $group.displayName

            if ($personDictionary.ContainsKey($personId)) {
                $matchedPerson = $personDictionary[$personId]
                $civicNo = $matchedPerson.civicNo.value

                $civicNoDisplaynameDictionary[$civicNo] = $displayName
            }
            else {
                continue
            }
        }
    }


}

function Match-GradeAuthority {
    param (
        [string[]]$gradeAuthorityPid,
        [string[]]$gradeAuthorityName
    )

    # Create an array 
    $formattedPairs = @()

    for ($i = 0; $i -lt $gradeAuthorityPid.Count; $i++) {
        $pids = $gradeAuthorityPid[$i] -replace '^..', ''   # Remove the first two from personnummer
        $nameParts = $gradeAuthorityName[$i] -split ',\s'   # Split the name into parts with a comma and space

        # Add hyphen 
        $pidsWithHyphen = $pids -replace '(\d{6})(\d{4})', '$1-$2'

        # Concatenate first and last name parts without a space if there are two last names
        if ($nameParts.Count -eq 2) {
            $formattedFirstName = $nameParts[1] -replace '\s', ''  # Remove any spaces in the first part of the name
            $formattedLastName = $nameParts[0] -replace '\s', ''  # Remove any spaces in the second part of the name
            $formattedPairs += "$pidsWithHyphen, $($formattedFirstName)$($formattedLastName)"
        }
        else {
            $formattedPairs += "$pidsWithHyphen, $($nameParts[0])"
        }
    }
    # Join the formatted pairs
    $output = $formattedPairs -join '|'

    return $output
}

function Match-Instructors {
    param (
        [string[]]$instructorPid,
        [string[]]$instructorName
    )

    # Create an array 
    $formattedPairs = @()

    for ($i = 0; $i -lt $instructorPid.Count; $i++) {
        $pids = $instructorPid[$i] -replace '^..', ''   # Remove the first two from personnummer
        $nameParts = $instructorName[$i] -split ',\s'   # Split the name into parts with a comma and space

        # Add hyphen 
        $pidsWithHyphen = $pids -replace '(\d{6})(\d{4})', '$1-$2'

        # Concatenate first and last name parts without a space if there are two last names
        if ($nameParts.Count -eq 2) {
            $formattedFirstName = $nameParts[1] -replace '\s', ''  # Remove any spaces in the first part of the name
            $formattedLastName = $nameParts[0] -replace '\s', ''  # Remove any spaces in the second part of the name
            $formattedPairs += "$pidsWithHyphen, $($formattedFirstName)$($formattedLastName)"
        }
        else {
            $formattedPairs += "$pidsWithHyphen, $($nameParts[0])"
        }
    }
    # Join the formatted pairs
    $output = $formattedPairs -join '|'

    return $output
}

function Read-EnterpriseXML($enterprise = $ContentOrganization.enterprise) {
    $global:EntUnits = [System.Collections.Hashtable]::new()
    $global:EntGroups = [System.Collections.Hashtable]::new()
    $global:EntPersons = [System.Collections.Hashtable]::new()
    $global:EntPersonsPiD = [System.Collections.Hashtable]::new()
    $global:EntMemberships = [System.Collections.Hashtable]::new()
    $global:EntActivities = [System.Collections.Hashtable]::new()
    $global:Entmembers = [System.Collections.Hashtable]::new()
     $global:EntPlacement = [System.Collections.Hashtable]::new()
 
    $global:EntSchoolType = $enterprise.properties.schooltype
 
    if ($enterprise.person) {
        # grabs id (guid)
        foreach ($p in $enterprise.person) {
            #if (!$global:EntPersons.ContainsKey($p.sourcedid.id))
            #{
            $global:EntPersons.Add($p.sourcedid.id, $p)
            #}
        }
    }
    if ($enterprise.person) {
        #grabs pid (personnummer)
        foreach ($p in $enterprise.person) {
            # Check if <userid> element exists and has the correct useridtype attribute
            $pdValue = $p.userid | Where-Object { $_.useridtype -eq 'PID' } | ForEach-Object { $_.'#text' }
    
            if ($pdValue) {
                # Add the PD value to the dictionary
                $global:EntPersonsPiD.Add($pdValue, $p)
            }
        }
    }
    

    if ($enterprise.group) {
        foreach ($g in $enterprise.group) {
            #if (! $global:EntGroups.ContainsKey($g.sourcedid.id))
            #{
            if ($g.grouptype.typevalue.'#text' -eq "Unit") {
                $global:EntUnits.Add($g.sourcedid.id, $g)
            }
 
            else {
                $global:EntGroups.Add($g.sourcedid.id, $g)
            }
            #}
        }
    }
    if ($enterprise.membership) {
        foreach ($member in $enterprise.membership.member) {
            
            if ($member.role.extension.placement) {
                
                $memberId = $member.sourcedid.id
                $placementInfo = $member.role.extension.placement
    
                
                $global:EntPlacement[$memberId] = $placementInfo
            }
        }
    }
    
    if ($enterprise.group) {
        foreach ($c in $enterprise.group) {
            
            if ($c.grouptype.typevalue.'#text' -eq "Class") {
                $global:EntClasses.Add($c.sourcedid.id, $c)
            }
 
            else {
                continue
            }
            
        }
    }
   
    if ($enterprise.membership) {
        foreach ($m in $enterprise.membership) {
            #if (!$global:EntMemberships.ContainsKey($m.sourcedid.id))
            #{
            if ($m.sourcedid.id -like "CG{*") {
                continue
            }
            $global:EntMemberships.Add($m.sourcedid.id, $m);
            #}
        }
    }
    
    if ($enterprise.membership.member) {
        foreach ($memberships in ($enterprise.membership)) {

            $membershipid = $memberships.sourcedid.id
            foreach ($member in  ( $memberships.member | where { $_.role.roletype -eq "Student" -and $_.sourcedid.id -notlike "CG{*" })) {

                $activity = $member.role.extension.activity

                if ($activity) {
                    
                    
                    
                    Add-Member -InputObject $activity -NotePropertyName 'MembershipID' -NotePropertyValue $membershipid -Force
                    Add-Member -InputObject $activity -NotePropertyName 'MemberID' -NotePropertyValue $member.sourcedid.id -Force
    
                    $memberBeginDate = $activity.begin
                    $memberEndDate = $activity.end
    
                    if ($memberBeginDate -and $memberEndDate) {
                        # Save if the activity is active
                        if ($memberBeginDate -le $global:RunDateStart -and $global:RunDateEnd -le $memberEndDate) {
                            # $activityId = $member.sourcedid.id
                            $activityId = $member.sourcedid.id.Trim('{}')
    
                            if (!$global:EntActivities.ContainsKey("$activityId")) {
                                $global:EntActivities["$activityId"] = [System.Collections.ArrayList]::new()
                                $global:EntActivities["$activityId"].Add($activity) | Out-Null
                            }
                            else {
                                $global:EntActivities["$activityId"].Add($activity) | Out-Null
                            }
                        }
                    }
                }
            }
        
        
      
        }
        
    }
    foreach ($membership in $enterprise.membership) {
        foreach ($member in $membership.member | where { $_.role.roletype -eq "Student" -and $_.sourcedid.id -notlike "CG{*" -and $_.role.extension.placement }) {
            $memberId = $member.sourcedid.id
            $global:Entmembers[$memberId] = $member
        }
    }
    

}


function Read-GradeOutcomeXml($XML, $type) {

    if ($type -eq 'Outcome') {
        if (!$XML) { $XML = $OutcomeGradesDataFixed }
        $Outcomes = $XML.outcome
                
        # $global:StudentOutcome = [System.Collections.Hashtable]::new() # Flytta ut den
        
        if ($OutcomeGradesDataFixed.outcome.properties.schooltype -eq 'GY' -or $OutcomeGradesDataFixed.outcome.properties.schooltype -eq 'GS') {
            if ($Outcomes.gradeoutcome) {
                foreach ($o in $Outcomes.gradeoutcome) {
                    foreach ($courseGrade in $o.coursegrade) {
                        $newStudentOutcome = [StudentGrades]::new()

                        $newStudentOutcome.Id = $o.student.id
                        $newStudentOutcome.Name = $o.student.Name

                        $newCourseData = [PSCustomObject]@{
                            SchoolName = $courseGrade.schoolname
                            UnitId     = $courseGrade.unitid
                            GroupId    = $courseGrade.groupid 
                        }

                        $newCourseTypeCode = [PSCustomObject]@{
                            CourseCode   = $courseGrade.course.code
                            CourseName   = $courseGrade.course.name
                            CoursePoints = $courseGrade.course.points
                        }

                        $newAssessor = [PSCustomObject]@{
                            AssessorId   = $courseGrade.assessor.id
                            AssessorName = $courseGrade.assessor.name
                        }

                        $newGradeOutcome = [PSCustomObject]@{
                            Date           = $courseGrade.date
                            Grade          = $courseGrade.grade
                            TrialPerformed = $courseGrade.trialperformed
                            Assessor       = $newAssessor
                        }

                        $newStudentOutcome.CourseData = $newCourseData
                        $newStudentOutcome.CourseTypeCode = $newCourseTypeCode
                        $newStudentOutcome.GradeOutcome = $newGradeOutcome

                        try {
                            if ($global:StudentOutcome[$newStudentOutcome.Id]) {
                                $global:StudentOutcome[$newStudentOutcome.id].Add($newStudentOutcome) | Out-Null
                            }
                            else {
                                $global:StudentOutcome[$newStudentOutcome.id] = [System.Collections.ArrayList]::new()
                                $global:StudentOutcome[$newStudentOutcome.id].Add($newStudentOutcome) | Out-Null
                            }
                        }
                        catch {
                            continue
                            Write-Host 'Exit'
                        }
                    }
                }
            }
        }
        else {
            if ($Outcomes.gradeoutcome) {
                foreach ($x in $Outcomes.gradeoutcome) {
                    foreach ($courseGrade in $x.subjectgrade) {
                        $NewElementaryStudentOutcome = [StudentGrades]::new()
                        $NewElementaryStudentOutcome.StudentID = $x.student.id
                        $NewElementaryStudentOutcome.StudentName = $x.student.name

                        $newCourseData = [PSCustomObject]@{
                            SchoolName = $courseGrade.schoolname
                            UnitID     = $courseGrade.unitid
                        }

                        $newCourseTypeCode = [pscustomobject]@{
                            SubjectCode = $courseGrade.subject.code
                            SubjectName = $courseGrade.subject.name
                        }

                        $newGradeOutcome = [pscustomobject]@{
                            Date           = $courseGrade.date
                            SemesterType   = $courseGrade.semester.type
                            SemesterYear   = $courseGrade.semester.year
                            Grade          = $courseGrade.grade
                            FinalGrade     = $courseGrade.finalgrade
                            Trailpreformed = $courseGrade.trialperformed
                        }

                        $NewElementaryStudentOutcome.CourseData = $newCourseData
                        $NewElementaryStudentOutcome.CourseTypeCode = $newCourseTypeCode
                        $NewElementaryStudentOutcome.GradeOutcome = $newGradeOutcome 

                        try {
                            if ($global:StudentOutcome[$NewElementaryStudentOutcome.StudentID]) {
                                $global:StudentOutcome[$NewElementaryStudentOutcome.StudentID].Add($NewElementaryStudentOutcome) | Out-Null
                            }
                            else {
                                $global:StudentOutcome[$NewElementaryStudentOutcome.StudentID] = [System.Collections.ArrayList]::new()
                                $global:StudentOutcome[$NewElementaryStudentOutcome.StudentID].Add($NewElementaryStudentOutcome) | Out-Null
                            }
                        }
                        catch {
                            continue
                            Write-Host 'Exit'
                        }
                    }
                }
            }
        }
    }
}

function Read-OfferingXML($type, $XML) {
    if ($type -eq "EducationPlans") {
        if (!$XML) { $XML = $ContentOfferingEducationPlans }
        $offering = $XML.offering
        $global:EntEducationPlans = [System.Collections.Hashtable]::new()
        $global:EntEducationPlansByCourse = [System.Collections.Hashtable]::new()
 
        if ($offering.educationplan) {
            foreach ($e in $offering.educationplan) {
                $newEducationPlan = [EducationPlan]::new()
                $newEducationPlan.Name = $e.name
                $newEducationPlan.Type = $e.type
                $newEducationPlan.Program = $e.program
        
                foreach ($courseType in $e.coursetype) {
                    
                    foreach ($course in $courseType.course) {
                        $newCourse = [EducationPlanCourse]::new()
                        $newCourse.CourseId = $course.id
                        $newCourse.CourseType = $courseType.type
                        $newCourse.CourseTypeCode = $courseType.code
                        $newCourse.CourseTypePoints = $courseType.points
                        $newCourse.CourseName = $course.name
                        $newCourse.CoursePoints = $course.points
                        $newCourse.CourseCode = $course.code
                        $newCourse.CourseSubjectCode = $course.subjectcode
                        $newCourse.CourseSubjectName = $course.subjectname
                        $newCourse.CourseLevel = $course.courselevel
        
                        $newEducationPlan.Courses.Add($newCourse.CourseCode, $newCourse)
        
                        $newEducationPlanTypes = [EducationplanTypes]::new()
                        $newEducationPlanTypes.Name = $e.name
                        $newEducationPlanTypes.Type = $e.type
                        $newEducationPlanTypes.Program = $e.program
                        $newEducationPlanTypes.CourseType = $courseType.type
        
                        $global:EntEducationPlansByCourse.Add("$($e.name)|$(($course.code))", $newEducationPlanTypes)
                    }
                }
        
                $global:EntEducationPlans.Add($e.name, $newEducationPlan)
            }
        }
        
    }
    
    if ($type -eq "Courses") {
        if (!$XML) { $XML = $ContentOfferingCourses }
        $offering = $XML.offering
        $global:EntCourses = [System.Collections.Hashtable]::new()
        $global:EntCoursesWithId = [System.Collections.Hashtable]::new()
 
 
 
        if ($offering.unitoffering) {
            foreach ($u in $offering.unitoffering) {
               
                foreach ($course in $u.course) {
                    $newCourse = [OfferingCourse]::new()
                    $newCourse.CourseId = $course.id
                    $newCourse.CourseName = $course.name
                    $newCourse.CoursePoints = $course.points
                    $newCourse.CourseCode = $course.code
                    $newCourse.CourseSubjectCode = $course.subjectcode
                    $newCourse.CourseSubjectName = $course.subjectname
                    $newCourse.CourseLevel = $course.courselevel
 
                    $global:EntCourses.Add("$($u.unit.id)|$(($course.code))", $newCourse)

                    $keyEntCoursesWithId = "$($course.id)|$(($course.code))"
                    if (-not $global:EntCoursesWithId.ContainsKey($keyEntCoursesWithId)) {
                        $global:EntCoursesWithId.Add($keyEntCoursesWithId, $newCourse)
                    }
                    else {
                        continue
                    }
                }
               
                #}
            }
        }
 
    }
 
 
    if ($type -eq "Subjects") {
        if (!$XML) { $XML = $ContentOfferingSubjects }
        $offering = $XML.offering
        $global:EntSubjects = [System.Collections.Hashtable]::new()
 
 
        if ($offering.unitoffering) {
            foreach ($s in $offering.unitoffering) {
               
                foreach ($subject in $s.subject) {
                    $newSubject = [OfferingSubject]::new()
                    $newSubject.SubjectId = $subject.id
                    $newSubject.SubjectCode = $subject.code
                    $newSubject.SubjectName = $subject.name
 
 
                    $global:EntSubjects.Add("$($s.unit.id)|$(($subject.code))", $newSubject)
                }
               
                #}
            }
        }
 
    }
}
#$member = $L2_SchoolMember
function Check-IsMemberActive($member) {
    $foundActivityOrPlacement = $false
    $TimeframeFound = $false

    # loop over all roles
    foreach ($role in $member.role) {
        # loop over all placements

        foreach ($placement in $role.extension.placement) {
            $memberBeginDate = $placement.begin
            $memberEndDate = $placement.end

            # assure that the placement have non empty start and end date
            if ($memberBeginDate -and $memberEndDate) {
                $foundActivityOrPlacement = $true

                # return true if the placement is active 
                if ($memberBeginDate -le $global:RunDateStart -and $global:RunDateEnd -le $memberEndDate) {
                    return [PSCustomObject]@{
                        begin = $memberBeginDate
                        end   = $memberEndDate
                    }
                }
            }
        }


        # loop over all activities
        foreach ($activity in $role.extension.activity) {
            $memberBeginDate = $activity.begin
            $memberEndDate = $activity.end

            if ($memberBeginDate -and $memberEndDate) {
                $foundActivityOrPlacement = $true

                # return true if the activity is active 
                if ($memberBeginDate -le $global:RunDateStart -and $global:RunDateEnd -le $memberEndDate) {
                    return [PSCustomObject]@{
                        begin = $memberBeginDate
                        end   = $memberEndDate
                    }
                    
                }
            }
        }
    }

    # if we found an activity or placement but none of them are active, return false
    if ($foundActivityOrPlacement) {
        return $false
    }
    else {
        # if no activity or placement found from any of the roles, use timeframe from role
        # loop over all roles
        
        foreach ($role in $member.role) {
            $memberBeginDate = $role.timeframe.begin
            $memberEndDate = $role.timeframe.end
            
            # assure that the timeframe have non empty start and end date
            if ($memberBeginDate -and $memberEndDate) {
                $TimeframeFound = $true

                # return true if the placement is active 
                if ($memberBeginDate -le $global:RunDateStart -and $global:RunDateEnd -le $memberEndDate) {
                    return [PSCustomObject]@{
                        begin = $memberBeginDate
                        end   = $memberEndDate
                    }
                    
                }
            }
        }
    }

    # if we havent found any placement, activity or timeframe, assume the member is active, else return false
    return $false
    #if (!$TimeframeFound)
    #{
    #    return $true
    #else
    #    return $false
    #}
}

function Fix-PersId($person = $L3_GroupPerson) {
    $persid = $person.userid | Where-Object { $_.useridtype -eq "PID" } | Select-Object -First 1 -ExpandProperty '#text'
    $persid = $persid.Insert($persid.Length - 4, "-")
    $persid = $persid.Substring(2)

    return $persid;
}

function Format-PersonnummerElev ($Id) {
    if ($Id -ne $null) {
        $Personnummer = $Id.Substring(2, 6)
        $FormattedId = $Id.Substring(8)

        # Formatting 
        return "{0}-{1}" -f $Personnummer, $FormattedId
    }
    else {
        continue
    }
}
function Format-Personnummer ($Personnummer) {
    if ($Personnummer -ne $null) {
        $PersonnummerPart = $Personnummer.Substring(2, 6)
        $FormattedIdPart = $Personnummer.Substring(8)

        # Formatting 
        return "{0}-{1}" -f $PersonnummerPart, $FormattedIdPart
    }
    else {
        return $null
    }
}

function Format-TeacherPersonnummer {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$TeacherIds,
        [string]$Separator = "|"
    )

    process {
        if ($TeacherIds -ne $null) {
            $formattedTeacherIds = $TeacherIds | ForEach-Object {
                $year = $_.Substring(2, 2)
                $month = $_.Substring(4, 2)
                $day = $_.Substring(6, 2)
                $number = $_.Substring(8)

                "{0}{1}{2}-{3}" -f $year, $month, $day, $number
            }

            $formattedString = $formattedTeacherIds -join $Separator

            Write-Output $formattedString
        }
    }
}

function FormatTeacherName {
    param (
        [string]$fullName
    )

    # Remove commas from the full name
    $fullName = $fullName -replace ','

    # Split the full name into individual components
    $names = $fullName -split ' '

    # Check if there are more than two names
    if ($names.Count -gt 2) {
        # Check if the last name consists of two parts (double last name)
        if ($names[-2] -match '\w+-\w+') {
            # Concatenate the last name parts along with the given names
            $formattedName = "$($names[-1])$($names[-2])$($names[1])$($names[0])"
        }
        else {
            # Concatenate the last name with the given names
            $formattedName = "$($names[-1])$($names[0])$($names[1])"
        }
    }
    else {
        # Reverse the order of names
        $formattedName = "$($names[1])$($names[0])"
    }

    return $formattedName
}


function Format-Principal {
    param (
        [string]$datePart,
        [string]$namePart
    )

    # Extract the date part
    $formattedDate = $datePart.Substring(2, 6) + "-" + $datePart.Substring(8, 4)

    # Format the name part
    $nameParts = $namePart -split ',\s*'  # Split by comma with optional space

    if ($nameParts.Count -eq 1) {
        
        $nameParts = $namePart -split ','
    }

    $formattedName = "$($nameParts[1])$($nameParts[0])"

    # Combine the formatted date and name
    $formattedString = "$formattedDate, $formattedName"

    

    return $formattedString
}

function Read-GradeHistory { 
    foreach ($SchoolType in $SchoolTypes.keys) {
        $CurrentDate = Get-Date
        $NumberOfHistoricYears = 9 # Data is empty after 2001.
        Write-Host "skoltypen är $SchoolType"
        # Loop over the last 10 years
        for ($i = 1; $i -le $NumberOfHistoricYears; $i++) {
            
            # Calculate start and end 
            $StartDate = $CurrentDate.AddYears(-$i).ToString("yyyy-01-01")
            $EndDate = $CurrentDate.AddYears(-$i).ToString("yyyy-12-31")
            Write-Verbose "Processing year $($CurrentDate.Year - $i)" -Verbose
            # Make the API request based on school type
            if ($SchoolType -like "Gymnasie*") {
                Write-Host "skoltypen är gymnasieskola"
                $APIresult = Invoke-WebRequest -Certificate $cert -Method get -Uri "$APIBaseURI/Outcome/$OutcomeVersion/Outcome/Get$($SchoolTypes[$SchoolType])Grades?LicenseKey=$($APILicenceKey)&StartDate=$($StartDate)&EndDate=$($EndDate)"
                
                $OutFile = "$WorkingFolder\Outcome____$SchoolType.xml"
                $OutcomeGradesDataFixeds = [System.Text.Encoding]::UTF8.GetString($APIresult.RawContentStream.ToArray())
                
                # Save to the file
                $OutcomeGradesDataFixeds | Out-File -Encoding UTF8 $OutFile
                
                # Load XML from the file
                $OutcomeGradesDataFixed = [xml]::new()
                $OutcomeGradesDataFixed.Load($OutFile)
                Read-GradeHistoryOutcomeXml -XML $OutcomeGradesDataFixed -type 'Outcome'  # Process data and add to Global dict for each iteration.
            }
            elseif ($SchoolType -like "grund*") {
                Write-Host "skoltypen är grundskola"
                $APIresult = Invoke-WebRequest -Certificate $cert -Method get -Uri "$APIBaseURI/Outcome/$OutcomeVersion/Outcome/Get$($SchoolTypes[$SchoolType])Grades?LicenseKey=$($APILicenceKey)&StartDate=$($StartDate)&EndDate=$($EndDate)"
                
                $OutFile = "$WorkingFolder\Outcome____$SchoolType.xml"
                $OutcomeGradesDataFixeds = [System.Text.Encoding]::UTF8.GetString($APIresult.RawContentStream.ToArray())
                
                # Save to the file
                $OutcomeGradesDataFixeds | Out-File -Encoding UTF8 $OutFile
                
                # Load XML from the file
                $OutcomeGradesDataFixed = [xml]::new()
                $OutcomeGradesDataFixed.Load($OutFile)
                Read-GradeHistoryOutcomeXml -XML $OutcomeGradesDataFixed -type 'Outcome'  # Process data and add to Global dict for each iteration.
            }
             
            # Add-HistoricGrades # Add iteration and export

            if ($i -eq $NumberOfHistoricYears) {
                Write-Host "Reached the end of loop for $SchoolType"
            }
        }
    }
}

function Read-GradeHistoryOutcomeXml($XML, $type) {

    if ($type -eq 'Outcome') {
        if (!$XML) { $XML = $OutcomeGradesDataFixed }
        $Outcomes = $XML.outcome
        # $global:StudentHistoryOutcome = [System.Collections.Hashtable]::new()
         
        
        
        if ($OutcomeGradesDataFixed.outcome.properties.schooltype -eq 'GY' -or $OutcomeGradesDataFixed.outcome.properties.schooltype -eq 'GS') {
           
            if ($Outcomes.gradeoutcome) {
                foreach ($o in $Outcomes.gradeoutcome) {
                    foreach ($courseGrade in $o.coursegrade) {
                        $newStudentOutcome = [StudentGrades]::new()

                        $newStudentOutcome.Id = $o.student.id
                        $newStudentOutcome.Name = $o.student.Name
                        #  Write-Host " GYMID$($newStudentOutcome.Id)"
                        #  Write-Host " GymdName $($newStudentOutcome.Name)"

                        $newCourseData = [PSCustomObject]@{
                            SchoolName = $courseGrade.schoolname
                            UnitId     = $courseGrade.unitid
                            GroupId    = $courseGrade.groupid
                        }

                        $newCourseTypeCode = [PSCustomObject]@{
                            CourseCode   = $courseGrade.course.code
                            CourseName   = $courseGrade.course.name
                            CoursePoints = $courseGrade.course.points
                        }

                        $newAssessor = [PSCustomObject]@{
                            AssessorId   = $courseGrade.assessor.id
                            AssessorName = $courseGrade.assessor.name
                        }

                        # Write-Host "Adding AssessoID $($courseGrade.assessor.id)"
                        # Write-Host "Adding Name $($courseGrade.assessor.name)"
                        $newGradeOutcome = [PSCustomObject]@{
                            Date           = $courseGrade.date
                            Grade          = $courseGrade.grade
                            TrialPerformed = $courseGrade.trialperformed
                            Assessor       = $newAssessor
                        }

                        $newStudentOutcome.CourseData = $newCourseData
                        $newStudentOutcome.CourseTypeCode = $newCourseTypeCode
                        $newStudentOutcome.GradeOutcome = $newGradeOutcome

                        try {
                            if ($global:StudentHistoryOutcome[$newStudentOutcome.Id]) {
                                $global:StudentHistoryOutcome[$newStudentOutcome.Id].Add($newStudentOutcome) | Out-Null
                            }
                            else {
                                $global:StudentHistoryOutcome[$newStudentOutcome.Id] = [System.Collections.ArrayList]::new()
                                $global:StudentHistoryOutcome[$newStudentOutcome.Id].Add($newStudentOutcome) | Out-Null
                            }
                        }
                        catch {
                            continue
                            Write-Host 'Exit'
                        }
                    }
                }
            }
        }

        if ($OutcomeGradesDataFixed.outcome.properties.schooltype -eq 'GR') {
            

            if ($Outcomes.gradeoutcome) {
                foreach ($x in $Outcomes.gradeoutcome) {
                    foreach ($courseGrade in $x.subjectgrade) {
                        $NewElementaryStudentOutcome = [StudentGrades]::new()
                        $NewElementaryStudentOutcome.StudentID = $x.student.id
                        $NewElementaryStudentOutcome.StudentName = $x.student.name
                         
                        
                        $newCourseData = [PSCustomObject]@{
                            SchoolName = $courseGrade.schoolname
                            UnitID     = $courseGrade.unitid
                        }

                        $newCourseTypeCode = [pscustomobject]@{
                            SubjectCode = $courseGrade.subject.code
                            SubjectName = $courseGrade.subject.name
                        }

                        $newGradeOutcome = [pscustomobject]@{
                            Date           = $courseGrade.date
                            SemesterType   = $courseGrade.semester.type
                            SemesterYear   = $courseGrade.semester.year
                            Grade          = $courseGrade.grade
                            FinalGrade     = $courseGrade.finalgrade
                            Trailpreformed = $courseGrade.trialperformed
                        }

                        $NewElementaryStudentOutcome.CourseData = $newCourseData
                        $NewElementaryStudentOutcome.CourseTypeCode = $newCourseTypeCode
                        $NewElementaryStudentOutcome.GradeOutcome = $newGradeOutcome

                        try {
                            if ($global:StudentHistoryOutcome[$NewElementaryStudentOutcome.StudentID]) {
                                $global:StudentHistoryOutcome[$NewElementaryStudentOutcome.StudentID].Add($NewElementaryStudentOutcome) | Out-Null
                            }
                            else {
                                $global:StudentHistoryOutcome[$NewElementaryStudentOutcome.StudentID] = [System.Collections.ArrayList]::new()
                                $global:StudentHistoryOutcome[$NewElementaryStudentOutcome.StudentID].Add($NewElementaryStudentOutcome) | Out-Null
                            }
                        }
                        catch {
                            continue
                            Write-Host 'Exit'
                        }
                    }
                }
            }
        }
        # Add-HistoricGrades # Function processes data and adds to $BetygExportDictonary and exports to csv.
    }
}

function FormatPeriod ($inputTimestamp) { 
   
    if (-not $inputTimestamp) {
        continue
    }

    try {
        $dateTime = [datetime]::ParseExact($inputTimestamp, "yyyy-MM-dd", $null)

        if ($dateTime) {
            $month = $dateTime.Month
            $result = "$($dateTime.Year % 100)$((1, 2)[$month -ge 8])"
            return $result
        }
    }
    catch {
        return ""
    }
}

function TranslateCourseTypes ($inputCourseType) {
    # Translate coursetypes to match old data.
  
    switch ($inputCourseType) {
        "ProgrammeSpecificSubjects" {
            return "PGÄ"
        }
        "UpperSecondarySubjects" {
            return "GGÄ"
        }
        "ProgrammeSpecialisation" {
            return "PFS"
        }
        "IndividualOptions" {
            return "IND"
        }
        "UpperSecondarySubjectsForStudentsWithLearningDisabilities" {
            return "GGF"
        }
        "AssessedCourseWork" {
            return "ACW"
        }
        "DiplomaProject" {
            return "GYA"
        }
        default {
            return $inputCourseType 
        }
    }
    
}

function Load-ProgramCodes {
    param(
        [string]$CsvPath = "ProgramkoderSkolverket.csv"
    )

   
    if (-not (Test-Path $CsvPath)) {
        Write-Host "Error: CSV file not found $CsvPath"
        return
    }

    $csvData = Import-Csv -Path $CsvPath -Delimiter ";"

    foreach ($row in $csvData) {
        $programName = $row.ProgramName.Trim()
        $programCode = $row.ProgramCode.Trim()

        
        $global:ProgramCodsSkolverket[$programCode] = $programName
    }

    Write-Host "Program codes loaded in dict."
}



#region Script
$global:ProgramCodsSkolverket = [System.Collections.Hashtable]::new()
Load-ProgramCodes -CsvPath "ProgramkoderSkolverket.csv"
GetSchoolData
Write-Verbose "Reading Tieto API - School data" -Verbose

$global:HistoricEntmembers = [System.Collections.Hashtable]::new()
$global:HistoricGroups = [System.Collections.Hashtable]::new()
$global:HistoricEntPersonsPiD = [System.Collections.Hashtable]::new()
Write-Verbose "Reading historic EnterPrise data" -Verbose
Read-HistoricEnterpriseXML

$ElevExportDictionary = [System.Collections.Hashtable]::new()
$SchoolExportDictionary = [System.Collections.Hashtable]::new()
$ProgramExportDictionary = [System.Collections.Hashtable]::new()
$BetygExportDictonary = [System.Collections.Hashtable]::new()
$global:StudentOutcome = [System.Collections.Hashtable]::new()
$global:StudentHistoryOutcome = [System.Collections.Hashtable]::new()
$global:EntClasses = [System.Collections.Hashtable]::new()
Read-GradeHistory # Add historic grades for students.

 
 




#   $SchoolType = "gymnasieskola"

#region Script
foreach ($SchoolType in $SchoolTypes.Keys) {

    
   

    #region GetOrganization

    Write-Verbose "Getting data from $("Get$($SchoolTypes[$SchoolType])Organization")..." -Verbose

    $OutFile = $("$WorkingFolder\Organization_$SchoolType.xml")
    $APIresult = Invoke-RestMethod "$($APIBaseURI)/$OrganizationVersion/Get$($SchoolTypes[$SchoolType])Organization?LicenseKey=$($APILicenceKey)" -OutFile $OutFile -TimeoutSec 6000 -Certificate $Cert -ErrorAction Stop

    $ContentOrganization = [xml]::new()
    $ContentOrganization.Load($OutFile)

    
    #endregion


    #region GetOffering    

    if ($SchoolType -like "Gymnasie*") {
        # Get Offering Gymnasieskola
        Write-Verbose "Getting data from $("Get$($SchoolTypes[$SchoolType])EducationPlans")..." -Verbose

        $OutFile = $("$WorkingFolder\Offering_$($SchoolType)_EducationPlans.xml")
        $APIresult = Invoke-RestMethod "$($APIBaseURI)/$OfferingVersion/Get$($SchoolTypes[$SchoolType])EducationPlans?LicenseKey=$($APILicenceKey)" -OutFile $OutFile -TimeoutSec 6000 -Certificate $Cert -ErrorAction Stop
                        
        $ContentOfferingEducationPlans = [xml]::new()
        $ContentOfferingEducationPlans.Load($OutFile)

        $OutFile = $("$WorkingFolder\Offering_$($SchoolType)_Courses.xml")
        $APIresult = Invoke-RestMethod "$($APIBaseURI)/$OfferingVersion/Get$($SchoolTypes[$SchoolType])Courses?LicenseKey=$($APILicenceKey)" -OutFile $OutFile -TimeoutSec 6000 -Certificate $Cert -ErrorAction Stop
                        
        $ContentOfferingCourses = [xml]::new()
        $ContentOfferingCourses.Load($OutFile)



    }
    else {
        # Get Offering Grundskola
        Write-Verbose "Getting data from $("Get$($SchoolTypes[$SchoolType])Offering")..." -Verbose

        $OutFile = $("$WorkingFolder\Offering_$($SchoolType)_Subjects.xml")
        $APIresult = Invoke-RestMethod "$($APIBaseURI)/$OfferingVersion/Get$($SchoolTypes[$SchoolType])Subjects?LicenseKey=$($APILicenceKey)" -OutFile $OutFile -TimeoutSec 6000 -Certificate $Cert -ErrorAction Stop
                        
        $ContentOfferingSubjects = [xml]::new()
        $ContentOfferingSubjects.Load($OutFile)
    }

        

    #endregion

    
    if ($SchoolType -like "gymnasie*") {
          
        $StartDate = "2023-01-01"
        $EndDate = "2023-12-01"
        
        Write-Verbose "Getting data from $("Get$($SchoolTypes[$SchoolType])Outcome")..." -Verbose
        $OutFile = "$WorkingFolder\Outcome_$SchoolType.xml"
   
        # Make the API request
        $APIresult = Invoke-WebRequest -Certificate $cert -Method get -Uri "$APIBaseURI/Outcome/$OutcomeVersion/Outcome/Get$($SchoolTypes[$SchoolType])Grades?LicenseKey=$($APILicenceKey)&StartDate=$($StartDate)&EndDate=$($EndDate)"
   
        # Convert the raw content to UTF-8 string and save it to a file
        $OutcomeGradesDataFixeds = [System.Text.Encoding]::UTF8.GetString($APIresult.RawContentStream.ToArray())
        $OutcomeGradesDataFixeds | Out-File -Encoding UTF8 $OutFile 
   
        $OutcomeGradesDataFixed = [xml]::new()
        $OutcomeGradesDataFixed.Load($OutFile)


    }
    elseif ($SchoolType -like "grundskola") {
        $StartDate = "2023-01-01"
        $EndDate = "2023-12-01"
        
        Write-Verbose "Getting data from $("Get$($SchoolTypes[$SchoolType])Outcome")..." -Verbose
          
        $OutFile = "$WorkingFolder\Outcome_$SchoolType.xml"

        # Make the API request
        $APIresult = Invoke-WebRequest -Certificate $cert -Method get -Uri "$APIBaseURI/Outcome/$OutcomeVersion/Outcome/Get$($SchoolTypes[$SchoolType])Grades?LicenseKey=$($APILicenceKey)&StartDate=$($StartDate)&EndDate=$($EndDate)"

        # Convert the raw content to UTF-8 string and save it to a file
        $OutcomeGradesDataFixeds = [System.Text.Encoding]::UTF8.GetString($APIresult.RawContentStream.ToArray())
        $OutcomeGradesDataFixeds | Out-File -Encoding UTF8 $OutFile 

        $OutcomeGradesDataFixed = [xml]::new()
        $OutcomeGradesDataFixed.Load($OutFile)
    }

     

      

     
      

    #region Export

    Write-Verbose "Exporting data" -Verbose

    Write-Verbose "Reading Enterprise" -Verbose
    Read-EnterpriseXML
    Write-Verbose "Reading Outcome" -Verbose
    Read-GradeOutcomeXml -XML $OutcomeGradesDataFixed -type 'Outcome'
    
   
    
    
    
    if ($SchoolType -like "gymnasie*") {
        Read-OfferingXML "EducationPlans" $ContentOfferingEducationPlans
        Read-OfferingXML "Courses" $ContentOfferingCourses
    }
    else {
        $global:EntEducationPlans = ""
        Read-OfferingXML "Subjects" $ContentOfferingSubjects
    }

    if (!($global:EntGroups -and $global:EntMemberships -and $global:EntUnits)) {
        Write-Verbose "Enterprise data missing" -Verbose
        exit 0
    }

   



    $global:RunDateStart = (Get-Date).AddDays(7).ToString("yyyy-MM-dd")
    $global:RunDateEnd = (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")



    # Fetch all groups of type "Unit" and where governedby = "MUNICIPAL", these are the schools in the municipality
    $EntGroupsInScope = $global:EntUnits.Values | where { $_.extension.governedby -eq "MUNICIPAL" -and $_.extension.municipalityname -eq "Katrineholm" }

    # $GroupMembers
    
    # Gå igenom alla kommunala skolor
    :L1_SchoolUnit foreach ( $L1_SchoolUnit in $EntGroupsInScope ) {

       

        $schoolExternalId = $L1_SchoolUnit.sourcedid.Id

        #$newElev = [ElevExport]::new()

        $Enhetsnamn = $L1_SchoolUnit.description.short
        $SkolenhetNamn = $L1_SchoolUnit.description.short
        $SkolenhetGUID = $L1_SchoolUnit.sourcedid.Id
        $Skoltyp = $global:EntSchoolType

        



        # slå upp ansvarig rektor och enhetskod
        $currPrincipal = $global:EntMemberships[ $schoolExternalId ].member | Where-Object { $_.role.roletype -eq "Principal" -and $_.role.status -eq "Active" } 
        if ($currPrincipal) { 
            $currPrincipal | ForEach-Object {
                $EnhetID = $_.role.extension.responsibility.schoolunitname
                $SkolenhetID = $_.role.extension.responsibility.schoolunitname
                $Skolenhetskod = $_.role.extension.responsibility.schoolunitCode
                $principal = $EntPersons[$_.sourcedid.id]
                $AnsvarigRektor = $principal.userid | where { $_.useridtype -eq "PID" } | Select-Object -ExpandProperty '#text'
                $principalName = $EntPersons[$_.sourcedid.id].name
                

            }
        }
        else {
            $EnhetID = ""
            $SkolenhetID = ""
            $Skolenhetskod = ""
            $AnsvarigRektor = ""
        }

        # Export data for school unit
        if ($SchoolExportDictionary.ContainsKey( "$schoolExternalId|$($Skoltyp)" )) {
            # "Dubblett $("$schoolExternalId|$($Skoltyp)")"
        }
        else {
            $newSchool = [SchoolExport]::new()
            $newSchool.ID = $SkolenhetID
            $newSchool.Name = $SkolenhetNamn
            $newSchool.Type = $Skoltyp
            $newSchool.SchooGUID = $SkolenhetGUID
            $newSchool.Principal = "$AnsvarigRektor - $principalName" 
            
            

            $SchoolExportDictionary.add("$schoolExternalId|$($Skoltyp)", $newSchool)
        }


        # hämta alla medlemmar i skolan
        $SchoolMembers = $global:EntMemberships[$schoolExternalId].member | Where-Object { $_.role.roletype -in ("Class") }

       
        
        



        #  gå igenom alla klasser
        :L2_SchoolMember foreach ($L2_SchoolMember in $SchoolMembers ) {
            $ClassGroup = $global:EntGroups[$L2_SchoolMember.sourcedid.id]
            $Klass = $ClassGroup.description.short
            $Årskurs = $ClassGroup.extension.schoolyear
            $KlassGUID = $ClassGroup.sourcedid.Id

            if ( $Årskurs -eq "0" -and $schooltype -like "grundskola" ) {
                continue L2_SchoolMember
            }
        
            $ClassMembers = $global:EntMemberships[$ClassGroup.sourcedid.Id].member | Where-Object { $_.role.roletype -in ("Student") }
            
            
           
            



            if (!$ClassMembers) {
                continue L2_SchoolMember 

                
            }
            
            #$ClassMemmbers
            :L3_ClassMember foreach ($L3_ClassMember in $ClassMembers) {
                
                
                
                if (! (Check-IsMemberActive($L3_ClassMember))) {
                    continue L2_SchoolMember
                }
            
                if ($SchoolType -eq "grundskola") {
                    $Utbildning = "GR"
                    $UtbildningNamn = "Grundskola" 
                }
                elseif ($SchoolType -eq "grundsärskola") {
                    $Utbildning = "GRS"
                    $UtbildningNamn = "Grundskolasärskola" 
                }
                else {
                    $Utbildning = $L3_ClassMember.role.extension.placement.programcode
                     

                    if ($global:ProgramCodsSkolverket.ContainsKey($Utbildning)) {
                        ## Check Skolkod againt dict from CSV.
                       
                       
                        $UtbildningNamn = $global:ProgramCodsSkolverket[$Utbildning]
                         
                        #  $global:ProgramCodsSkolverket["IMA"]
                        
                    }
                    else {
                        $UtbildningNamn = $global:EntEducationPlans[$Utbildning].Program 
                    }

                    # Write-Host "Utbildningsnamn: $UtbildningNamn"
                    #  $Utbildning = $L3_ClassMember.role.extension.placement.programcode 
                    #  $UtbildningNamn = $global:EntEducationPlans[$Utbildning].Program 
                }
                
               

                $Person = $global:EntPersons[$L3_ClassMember.sourcedid.id]

                if ($ExcludePrivacy -and $Person.extension.privacy.'#text' -eq "true") {
                    ## Remove student with <privacy> level.
                    Write-Host 'removed student'
                    continue
                }
                
                $newElev = [ElevExport]::new()

                $newElev.Enhetsnamn = $Enhetsnamn
                $newElev.SkolenhetNamn = $SkolenhetNamn
                $newElev.SkolenhetGUID = $SkolenhetGUID
                $newElev.Skoltyp = $Skoltyp
        
                $newElev.EnhetID = $EnhetID
                $newElev.SkolenhetID = $SkolenhetID
                $newElev.Skolenhetskod = $Skolenhetskod
                # $newElev.AnsvarigRektor = "$AnsvarigRektor - $($principalName.fn)"
                $newElev.AnsvarigRektor = "$(Format-Principal -datePart $AnsvarigRektor -namePart $principalName.fn)"

                $newElev.Klass = $Klass
                $newElev.Årskurs = $Årskurs
                $newElev.KlassGUID = $KlassGUID

                $newElev.Förnamn = $Person.name.n.given
                $newElev.Efternamn = $Person.name.n.family
                $newElev.Personnummer = Fix-PersId($Person)
                $newElev.Kön = $(if ($person.demographics.gender -eq "Female") { "K" }else { "M" })
                $newElev.Utbildning = $Utbildning
                

                $newElev.UtbildningNamn = $UtbildningNamn 
                $newElev.ElevGUID = $Person.sourcedid.Id
      
                
                if ($ElevExportDictionary.ContainsKey("$( $newElev.Personnummer )|$( $newElev.Klass )")) {
                    # "Dubblett $($newElev.Personnummer)"
                }
                else {
                    $ElevExportDictionary.add("$($newElev.Personnummer)|$($newElev.Klass)", $newElev)
 
                     
                    
                    
                  



                    # // Läser nuvarande skolår från Organisationfilen.

                    # $persid = "200409239332"
                    $persid = $person.userid | Where-Object { $_.useridtype -eq "PID" } | Select-Object -First 1 -ExpandProperty '#text'
                   
                    
                    $guid = $newElev.ElevGUID.Trim('{}') 


                    $ForeCastmatches = $global:EntActivities[$guid]

                    # Check if no result is found with the original GUID
                    if (-not $ForeCastmatches) {
                        # Try with curly braces around the GUID
                        $ForeCastmatches = $global:EntActivities["{$guid}"]
                    }
                    



                    
                    $GetMemberSchoolyear = $global:Entmembers[$newElev.ElevGUID]    
                    
                    if ($GetMemberSchoolyear) {
                        $schoolYear = $GetMemberSchoolyear.role.extension.placement.schoolyear
                    }
                    else {
                        $schoolYear = 'Schoolyear empty'
                    }
                    
                    foreach ($activity in $ForeCastmatches) {
                        


                        $GetMembershipID = $global:EntMemberships[$activity.MembershipID]
                        #  Write-Host "Membership ID: $($GetMembershipID.sourcedid.id)"

                         

                             $test = $global:EntClasses[$KlassGUID]
                            if($test.timeframe.begin -eq $activity.begin ){

                                $TestSchoolYeat = $test.extension.schoolyear
                                
                            }
                            else {
                                $TestSchoolYeat = "No value in TestSchoolYeat"
                            }
                        

                        $instructorMember = $GetMembershipID.member | Where-Object { $_.role.roletype -eq "Instructor" -or $_.role.roletype -eq "GradeAuthority" }

                        if ($instructorMember) {
                            $instructorId = $instructorMember.sourcedid.id
                            $GetInstructor = $global:EntPersons[$instructorId]

                            if ($GetInstructor) {
                                $instructorRole = $instructorMember.role.roletype

                                if ($instructorRole -eq "Instructor") {
                                    $instructorUserId = $GetInstructor.userid | Where-Object { $_.useridtype -eq "PID" }
                                    if ($instructorUserId) {
                                        $instructorPid = $instructorUserId.'#text'
                                        $GetinstructorName = $GetInstructor.name.fn


                                        $instructor = Match-Instructors -instructorPid $instructorPid -instructorName $GetinstructorName

                                    }
                                    else {
                                        # Write-Host "Instructor PID not found."
                                    }
                                }
                                elseif ($instructorRole -eq "GradeAuthority") {
                                    
                                    
                                    $gradeAuthorityUserId = $GetInstructor.userid | Where-Object { $_.useridtype -eq "PID" }
                                    if ($gradeAuthorityUserId) {
                                        $gradeActivityDate = $instructorMember.role.extension.activity.begin
                                        $gradeAuthorityPid = $gradeAuthorityUserId.'#text'
                                        $gradeAuthorityName = $GetInstructor.name.fn
                                        

                                        $gradeAutority = Match-GradeAuthority -gradeAuthorityPid $gradeAuthorityPid -gradeAuthorityName $gradeAuthorityName
                                      

                                        

                                    }
                                    else {
                                        # Write-Host "GradeAuthority PID not found."
                                    }
                                }
                            }
                            else {
                                # Write-Host "Instructor not found in the dictionary."
                            }
                        }
                        
                        $EducationPlanName = $global:EntEducationPlans[$Utbildning].Name 
                        $GetCourseGroup = $global:EntEducationPlansByCourse["$($EducationPlanName)|$($activity.coursecode)"] # Get´s kurstyp(HUR)
                        $GetCourseName = $global:EntCoursesWithId["$($activity.courseid)|$($activity.coursecode)"] # Get´s Kursnamn

                        if (-not $EducationPlanName) {
                            $EducationPlanName = "No value found EducationPlanName"
                        }
                        
                     
                        
                        if (-not $GetCourseName) {
                            $GetCourseName = "No value found GetCourseName"
                        }
                       
                
                            
                        if (!$GetCourseGroup) {
                                    
                            try {
                                $GetCourseGroup = $global:EntEducationPlansByCourse["$($EducationPlanName.Substring(0,2))|$($activity.coursecode)"]
                            }
                            catch {
                                # $EducationPlanName is null
                                continue
                            }
                                   
                        }
                           
                       
     
                        $NewSubjectForecast = [OutcomeMandatoryExport]::new()

                        if ($global:EntGroups[$activity.MembershipID]) {
                            $GetMembership = $global:EntGroups[$activity.MembershipID]
                        }
                        else {
                            if ($civicNoDisplaynameDictionary[$persid]) {
                                $GetMembership = $civicNoDisplaynameDictionary[$persid]
                            }
                            else {
                                $GetMembership = "No value found GetMembership"
                            }
                        }
                        

                        $GetMembership.description
                        
                        $NewSubjectForecast.PersonnummerElev = Format-PersonnummerElev -Id $persid
                        $NewSubjectForecast.GruppNamn = $GetMembership.description.short
                        $NewSubjectForecast.ÄmneKurs = $( if ($SchoolType -like 'gymnasie*') { $activity.coursecode } else { $activity.subjectcode } )
                        $NewSubjectForecast.Kursnamn = $GetCourseName.CourseName
                        $NewSubjectForecast.Poäng = $activity.hours
                        $NewSubjectForecast.Startdatum = $activity.begin
                        $NewSubjectForecast.Slutdatum = $activity.end
                        $NewSubjectForecast.Betyg = ""
                        $NewSubjectForecast.Akttyp = ""
                        $NewSubjectForecast.Hur = TranslateCourseTypes -inputCourseType $GetCourseGroup.CourseType
                        $NewSubjectForecast.Period = FormatPeriod -inputTimestamp $activity.begin
                        $NewSubjectForecast.BetygsättandeLärare = if ($activity.begin -eq $gradeActivityDate) { $gradeAutority }else { "" } # Match the date of the gradeauth to the date of the activity to get current data.
                        $NewSubjectForecast.AllaUndervisandeLärare = $instructor
                        $NewSubjectForecast.LåstBetyg = ""
                        $NewSubjectForecast.AktivitetensÅrskurs = $TestSchoolYeat
                        $NewSubjectForecast.AktivitetensGUID = $( if ($SchoolType -like 'gymnasie*') { $activity.courseid } else { $activity.subjectid } )
                        $NewSubjectForecast.GruppGUID = ""
               
                        if ($BetygExportDictonary.ContainsKey("$($NewSubjectForecast.ÄmneKurs)|$($NewSubjectForecast.PersonnummerElev)")) {
                            continue
                        }
                        else {
                            $BetygExportDictonary.Add("$($NewSubjectForecast.ÄmneKurs)|$($NewSubjectForecast.PersonnummerElev)", $NewSubjectForecast)
                        }
                    }  
            





                

                       





                    $PlannedHistoryOutcome = $global:StudentHistoryOutcome[$persid]
                    foreach ($HistCourses in $PlannedHistoryOutcome) {

                        
                        $GetmemberId = $global:EntPersonsPiD[$HistCourses.Id] # Get the MemberId

                        $GetMembershipID = $global:HistoricEntmembers[$GetmemberId.sourcedid.id] # Get the membershipID

                        $GetGroup = $global:HistoricGroups[$GetMembershipID.MembershipID]



                        $historicOutcome = [OutcomeMandatoryExport]::new()
                        if ($HistCourses.Id) {

                            $historicOutcome.PersonnummerElev = Format-PersonnummerElev -Id $HistCourses.Id
                            $historicOutcome.GruppNamn = "" # $HistCourses.CourseData.SchoolName | tomt på historsiak betyg.
                            $historicOutcome.ÄmneKurs = $HistCourses.CourseTypeCode.CourseCode
                            $historicOutcome.Kursnamn = $HistCourses.CourseTypeCode.CourseName
                            $historicOutcome.Poäng = $HistCourses.CourseTypeCode.CoursePoints
                            $historicOutcome.Startdatum = $HistCourses.GradeOutcome.Date
                            $historicOutcome.Slutdatum = $HistCourses.GradeOutcome.Date
                            $historicOutcome.Betyg = $HistCourses.GradeOutcome.Grade
                            $historicOutcome.Akttyp = ""
                            $historicOutcome.Hur = ""
                            $historicOutcome.Period = FormatPeriod -inputTimestamp $HistCourses.GradeOutcome.Date
                            $historicOutcome.BetygsättandeLärare = "$(Format-Personnummer -Personnummer $HistCourses.GradeOutcome.Assessor.AssessorId), $(FormatTeacherName -fullName $HistCourses.GradeOutcome.Assessor.AssessorName -replace ',',' ')"
                            $historicOutcome.AllaUndervisandeLärare = ""
                            $historicOutcome.LåstBetyg = $HistCourses.GradeOutcome.trialperformed
                            $historicOutcome.AktivitetensGUID = $HistCourses.CourseData.UnitId
                            $historicOutcome.AktivitetensÅrskurs = $SetSchoolYear
                            $historicOutcome.GruppGUID = ""
                        }
                        elseif ($HistCourses.StudentId) {
                            $historicOutcome.PersonnummerElev = Format-PersonnummerElev -Id $HistCourses.StudentId
                            $historicOutcome.GruppNamn = "" # $HistCourses.CourseData.SchoolName | tomt på historsiak betyg.
                            $historicOutcome.ÄmneKurs = $HistCourses.CourseTypeCode.SubjectCode
                            $historicOutcome.Kursnamn = $HistCourses.CourseTypeCode.SubjectName
                            $historicOutcome.Poäng = $HistCourses.CourseTypeCode.CoursePoints
                            $historicOutcome.Startdatum = FormatPeriod -inputTimestamp $HistCourses.GradeOutcome.Date
                            $historicOutcome.Slutdatum = $HistCourses.GradeOutcome.Date
                            $historicOutcome.Betyg = $HistCourses.GradeOutcome.Grade
                            $historicOutcome.Akttyp = ""
                            $historicOutcome.Hur = ""
                            $historicOutcome.Period = FormatPeriod -inputTimestamp $HistCourses.GradeOutcome.Date
                            $historicOutcome.BetygsättandeLärare = "$(Format-Personnummer -Personnummer $HistCourses.GradeOutcome.Assessor.AssessorId), $( FormatTeacherName -fullName $HistCourses.GradeOutcome.Assessor.AssessorName -replace ',',' '))" -replace ', $', ''
                            $historicOutcome.AllaUndervisandeLärare = ""
                            $historicOutcome.LåstBetyg = $HistCourses.GradeOutcome.Trailpreformed
                            $historicOutcome.AktivitetensÅrskurs = $HistCourses.GradeOutcome.SemesterType
                            $historicOutcome.AktivitetensGUID = $HistCourses.CourseData.UnitID.UnitId
                            $historicOutcome.GruppGUID = ""
                        }
                   
                        
                        

                        if ($BetygExportDictonary.ContainsKey("$($historicOutcome.PersonnummerElev)|$($historicOutcome.ÄmneKurs)|$($historicOutcome.Startdatum)")) { 
                         
                            continue
                        }
                        else {
                           
                            $BetygExportDictonary.Add("$($historicOutcome.PersonnummerElev)|$($historicOutcome.ÄmneKurs)|$($historicOutcome.Startdatum)", $historicOutcome)
                        }
                    }   
 

                   
            

                


                }       
            }
 

        }

        foreach ($EducationPlan in $global:EntEducationPlans.Values) {

            foreach ($Course in $EducationPlan.Courses.Values) {
                $newProgram = [ProgramMandatoryExport]::new()
                $newProgram.ProgramID = $EducationPlan.Name
                $newProgram.ProgramName = $EducationPlan.Program
                $newProgram.CourseID = $Course.CourseCode
                $newProgram.CourseName = $Course.CourseName
                $newProgram.CourseCode = $Course.CourseCode
                $newProgram.CourseType = $Course.CourseType
                $newProgram.CourseLevel = $Course.CourseLevel

                if ($ProgramExportDictionary.ContainsKey("$( $EducationPlan.Name )|$( $newProgram.CourseID )|$( $newProgram.CourseType )")) {
                    # "Dubblett $("$( $EducationPlan.Name )|$( $newProgram.CourseID )|$( $Course.CourseType )")"
                    continue
                }
                else {
                    $ProgramExportDictionary.add("$( $EducationPlan.Name )|$( $newProgram.CourseID )|$( $newProgram.CourseType )", $newProgram)
                }
            }
       
        }

    }
  


}




$ElevExportDictionary.Count
$SchoolExportDictionary.Count
$ProgramExportDictionary.Count
$BetygExportDictonary.Count





$ElevExportDictionary.values | Export-Csv -Encoding utf8BOM -Delimiter ";" -Path "$WorkingFolder\EDLEVOELEV.csv" -NoTypeInformation -Force
$SchoolExportDictionary.values | Export-Csv -Encoding utf8BOM -Delimiter ";" -Path "$WorkingFolder\EDLEVOSKOLOR.csv" -NoTypeInformation -Force
$ProgramExportDictionary.values | Export-Csv -Encoding utf8BOM -Delimiter ";" -Path "$WorkingFolder\EDLEVOPROGRAM.csv" -NoTypeInformation -Force
# $BetygExportDictonary.values | Export-Csv -Encoding utf8BOM -Delimiter ";" -Path "$WorkingFolder\EDLEVOBETYG.csv" -NoTypeInformation -Force
$sortedValues = $BetygExportDictonary.Values | Sort-Object -Property PersonnummerElev
$sortedValues | Export-Csv -Encoding utf8BOM -Delimiter ";" -Path "$($WorkingFolder)\EDLEVOBETYG.csv" -NoTypeInformation -Force

#endregion






