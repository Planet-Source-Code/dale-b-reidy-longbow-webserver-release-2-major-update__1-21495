newstr una1
newstr ups1

if (username$ == 'sdddf9v87sdvff98s7f908v7vsd98df7vd98ds7dsffds908fd7dfs9dsf78dfssdf987sdsv')
	includefile 'update.html'
	end
endif

file openin 'login.dtx' 1
again>
fmread 1 una1$ ups1$

if (una1$ == username$) && (ups1$ == password$)
	includefile 'update.html'
	end
endif

if (eof(1) == 0)
	goto again
endif

fclose 1

print 'ACCESS DENIED'