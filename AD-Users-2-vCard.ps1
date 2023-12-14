<#
.SYNOPSIS

Export AD Users as single vcf file

#>

Clear-Host

$filename = "C:\vCards\abook.vcf"

remove-item $filename

Get-ADUser -filter {Enabled -eq $True} -Properties * | where {$_.enabled -eq $true} | Foreach-Object {

$user = $_.sAMAccountName
$email = $_.EmailAddress
$phone = $_.OfficePhone
$name = $_.Name
$short = $_.HomePhone
$mobile = $_.MobilePhone
$sede = $_.physicalDeliveryOfficeName
$ufficio = $_.Department
$qualifica = $_.title
#$responsabile = ($_.manager -split ',*..=')[1]

if ($phone -or $email -or $mobile){
    write-host $user

    # Generate vCard
    Add-Content -Path $filename "BEGIN:VCARD"
    Add-Content -Path $filename "VERSION:3.0"
    Add-Content -Path $filename ("FN:" + $_.givenName + " " + $_.sn)
    Add-Content -Path $filename ("ORG:Ing.Polin E C. S.p.A.")
    Add-Content -Path $filename ("TITLE:" + $qualifica)
    Add-Content -Path $filename ("TEL;TYPE=WORK:" + $phone)
    Add-Content -Path $filename ("TEL;TYPE=HOME:" + $short)
    Add-Content -Path $filename ("TEL;TYPE=CELL:" + $mobile)
    #Add-Content -Path $filename ("ADR;WORK;PREF:" + ";;" + $_.streetAddress + ";" + $_.l + ";" + $_.st + ";" + $_.postalCode)
    Add-Content -Path $filename ("EMAIL;TYPE=WORK:" + $email)
    Add-Content -Path $filename ("CATEGORIES:" + $ufficio)
    $notedate = (get-date).ToString('d/MM/yyyy')
    Add-Content -Path $filename ("NOTE:[*** DO NOT EDIT ***] Il contatto viene sovrascritto. Ultimo aggiornamento: $notedate")
    Add-Content -Path $filename ("UID:" + $_.objectGUID )
    Add-Content -Path $filename "END:VCARD" #`r`n"
    }

}
