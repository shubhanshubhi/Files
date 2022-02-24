$server1 = "C:\Logfile\"
$server2 = "\\DWSERVER\c$\MovedLogFile\"
foreach ($server1 in gci $server1 -include *.csv -recurse)
 { 
 Move-Item -path $server1.FullName -destination  $server2
 }