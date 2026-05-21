$dir = 'D:\<END FOLDER>'

(Get-Acl -Path $dir).Access | Format-Table -AutoSize 

