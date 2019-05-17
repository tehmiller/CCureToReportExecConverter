<#
	CCure Export to Report Exec CSV Generator
#>

#Load CCure XML export
[xml]$ccex = Get-Content -Path staff_students.xml
$nodes = $ccex.CrossFire.'SoftwareHouse.NextGen.Common.SecurityObjects.Personnel'

#Array of users
$users = New-Object System.Collections.ArrayList

#Process each user in the export
foreach($user in $nodes)
{
    #If the user doesn't have a Student ID then things can break, skip them
    if($user.UDF__Student_ID__.Length -eq 0) { continue }

    $imgName = $user.UDF__Student_ID__
    $imgName = $pwd.Path + "\Pictures\" + $imgName + ".jpg"
    
    #User does not have a picture
    if($user.'SoftwareHouse.NextGen.Common.SecurityObjects.Images' -eq $null) {
        $imgStr = ""
    }
    else {
        #Handling the users that have multiple pictures
        if($user.'SoftwareHouse.NextGen.Common.SecurityObjects.Images'.GetType().IsArray) {
            $imgStr = $user.'SoftwareHouse.NextGen.Common.SecurityObjects.Images'[0].Image    
        }
        else {
            $imgStr = $user.'SoftwareHouse.NextGen.Common.SecurityObjects.Images'.Image
        }

        $bytes = [Convert]::FromBase64String($imgStr)
        [IO.File]::WriteAllBytes($imgName, $bytes)
    }
    $DOB = if($user.UDF__Date_of_Birth_ -ne $null) { $(Get-Date $user.UDF__Date_of_Birth_ -Format d) } else { "" }

    $reuser = [PSCustomObject]@{
    	'First Name*' = "$($user.FirstName)"
		'Middle Name' = "$($user.MiddleName)"
		'Last Name*' = "$($user.LastName)"
		'Date Of Birth' = "$DOB"
		Gender = ''
		Race = ''
		'Hair Color' = ''
		'Eye Color' = ''
		Height = ''
		Weight = ''
		'Drivers License Number' = ''
		'Drivers License State' = ''
		Department = ''
		Title = ''
		SSNumber = "$($user.UDF__Student_ID__)"
		'Street Number' = ''
		'Street Name' = ''
		'Street Type' = ''
		Direction = ''
		Apartment = ''
		City = ''
		State = ''
		Zip = ''
		HomeNumber = ''
		CellNumber = ''
		WorkNumber = ''
		Notes = ''
		Email = ''
		Photo = "$imgName"
    }
 
    $users.Add($reuser) | out-null
}

$users | Export-Csv -NoTypeInformation ContactTemplate.csv