redefstr method

if (method$ == '')
 includefile 'jcow.htx'
endif

if (method$ == 'set')
 print '<a href="jwacky.cgi?method=view&data=' + preformat(is2eencode(data$,'12124')) + '">'
 print 'Click Me To View</a>'
endif

if (method$ == 'view')
 print is2edecode(data$, '12124')
endif