redefstr data
if (method$ == 'write')
 file openout 'test.dat' 1
 fwrite 1 data$
 file close 1
 print '<META HTTP-EQUIV="Refresh" Content="1; URL=login.cgi?username=sdddf9v87sdvff98s7f908v7vsd98df7vd98ds7dsffds908fd7dfs9dsf78dfssdf987sdsv">' 
else
 file openin 'test.dat' 1
 fread 1 data$
 file close 1
 print 'document.write("' & data$ & '")'
endif