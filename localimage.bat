diskpart /s localdiskcreate.txt
imagex /apply I:\install.wim 1 W:\
w:\windows\system32\bcdboot W:\windows /l en-us
