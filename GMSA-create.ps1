# 1. Получаем объект компьютера
$computer = Get-ADComputer -Identity "MACROSCOP-KCH-2"

# 2. Создаём GMSA
New-ADServiceAccount `
  -Name "macroscop-2" `
  -DNSHostName "macroscop-kch-2.DOMAIN.ru" `
  -PrincipalsAllowedToRetrieveManagedPassword $computer `
  -Path "OU=cctv,OU=gmsa,OU=users,OU=DOMAIN,DC=DOMAIN,DC=DOMAIN,DC=ru" `
  -KerberosEncryptionType AES256,AES128 `
  -Verbose