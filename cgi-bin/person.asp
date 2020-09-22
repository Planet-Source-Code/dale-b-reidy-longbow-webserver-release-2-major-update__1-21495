
if (instr(1,referer$,'/cgi-bin/lovepage.asp') > 0)

print '<html>'
print '<title>'
print person1$
print ' LOVES '
print person2$
print '</title>'
print '<body bgcolor="#ffffff">'

print '<p align="center">&nbsp;</p>'

print '<p align="center">&nbsp;</p>'

print '<p align="center"><font size="4" face="comic sans ms" color=#0000FF><strong>'
print person1$
print '<img '
print 'src="hrt.jpg" width="160" height="186">'
print person2$
print '</strong></font></p>'
print '</body>'
print '</html>'

else

 includefile 'LOVEPAGENOACCESS.html'

endif