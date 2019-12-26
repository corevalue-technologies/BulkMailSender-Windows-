<#
    Developed By : Deepak Singh
    this code was intended to run into windows server where it was very likely to send emails to thousands of users 
    and using java code to do that was very costly to server, In that case we were using powershell which is fast then java mail api.
    you don't need to give your username or password you just have to login to your outlook
    The biggest chanllege was to check if the user exists in Active directory, Otherwise we will have to bear the cost of wrong email.
    here I was using AD for checking the user and if he is active or not
#>

#function to send mails
function sendmail($sendingarray){
   $userstatus = Get-ADUser $sendingarray[2]|Select-Object -ExpandProperty "Enabled" #Check whether the user is enable or not
   if($userstatus){
   $string1 = "<tr>" #HTML string
   $counter=1
  foreach($i in $sendingarray){

        $string1+="<td>"+$i+"</td>"

        if($counter-eq 3){

            $string1+="</tr><tr>"

            $counter=0
        }

        $counter+=1

   }

   #create COM object named Outlook

    $Outlook = New-Object -ComObject Outlook.Application

    $user = Get-ADUser $sendingarray[2] | Select-Object -ExpandProperty "GivenName" #will use the get-aduser command for getting the username
    Write-Host "sending mail to $user"
    #create Outlook MailItem named Mail using CreateItem() method

    $Mail = $Outlook.CreateItem(0)

    #add properties as desired
    $user1 =Get-ADUser $sendingarray[2] -Properties EmailAddress|Select-Object -ExpandProperty "EmailAddress"
    $Mail.To = "$user1"
    $Mail.CC = "GCC@example.com"

    $Mail.Subject = "your subject"
    #Body of the mail Change your singnature.
    $Mail.HTMLBody = "

    <html>
    <head>
        <style>
        table,td,th {
         border : 1px solid black;
        }

       td {
            text-align: left;
            padding: 2px;
 
        }

        td:nth-child(even){background-color: yellow;}

        th {
            background-color: #0099ff;
            color: black;
            padding : 2px;
        }
        </style>
</head>

    <body>
    Hello $user,<br>

    <p>

    We are contacting you to update the secondary owner in the Network folder lookup in service catalog.<br><br>



    We could see that you are marked as primary owner of the below mentioned folder and secondary owner details

    are not available.<br>
    Kindly let us know the secondary owner of the folder.<br><br>

    </p>

    <table>

    <tr>

    <th>Server Name</th>

    <th>Share Name</th>

    <th>Primary Owner</th>

    </tr>

                                               

       "+

    $string1+

    "        

    </table></br>
    </br>
    <p>
Thanks & Regards<br><br>

Deepak Singh<br>

Global Hosting & Storage | deepak.singh@example.com<br>

Baker Hughes, a GE Company<br>

Please copy GHS_Wintel_Support@example.com in all your email communications, in order, for us to service your request better.<br>
    </p>
    </body>
    </html>

    "

    #send message

    $Mail.Send()

    #quit and cleanup

    #$Outlook.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

  }else{
  $user = $sendingarray[2]
    Write-Host "Mail not sent to $user "
  }

}




#start row please check if there is a header or not. if no header make it 1
$row = 2

#[int]$endrow = Read-Host "Enter the end row"

#Check the path this is important Change the 431 id to your 431 id

$file = "C:\Users\singdee01\Desktop\sheet3.xlsx"

#creating object of excel to read the excel sheet
$xl=New-Object -ComObject "Excel.Application"

#opening excel sheet
$wb=$xl.Workbooks.Open($file)

$ws=$wb.ActiveSheet
 
$servernam = @() #array of servere names

$sharenam = @() #array of sharenames

$primaryown = @() #array of primary owners

do{

$data = $ws.Range("A$row").text

if($ws.Range("E$row").text -eq 'mail sent'){

    $row+=1

    continue

}


#logic to remove duplicacy.
$data2 = $ws.Range("B$row").text

$data3 = $ws.Range("C$row").text

$servernam +=$data

$sharenam += $data2

$primaryown += $data3

$row+=1

}while($data)

$sendingarray = @()

$anotherarray = @()

$pri=$primaryown[0]

$i = 0

do{

      if($primaryown[$i] -ne $pri){

        $pri = $primaryown[$i]

        sendmail $sendingarray

        $sendingarray = @()       

      }

        $anotherarray = @()

        $anotherarray += $servernam[$i]

        $anotherarray+= $sharenam[$i]

        $anotherarray+= $primaryown[$i]

        $sendingarray += $anotherarray

         $i+=1

}while($i -ne $servernam.Length)

$xl.Workbooks.Close()

$xl.Quit()
#relesing the excel object which prints 0 at the end. this is important
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
