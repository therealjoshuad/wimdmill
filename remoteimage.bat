diskpart /s remotediskcreate.txt
set servername=SERVERNAME
net use Q: "\\%servername%\System Images" /user:USERNAME PASSWORD
imagex /apply Q:\install.wim 1 W:\
w:\windows\system32\bcdboot W:\windows /l en-us
exit
