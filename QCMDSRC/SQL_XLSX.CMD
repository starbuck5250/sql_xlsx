cmd 'SQL to XLSX (Excel) file'

parm kwd(sql) type(*char) len(2560) min(1) expr(*yes) +
     case(*mixed) prompt('SQL statement')
parm kwd(stmf) type(*char) len(512) min(1) expr(*yes) +
     case(*mixed) prompt('IFS stream file name')
parm kwd(user) type(*char) len(10) min(1) expr(*yes) +
     case(*mixed) prompt('User ID')
parm kwd(password) type(*char) len(10) min(1) expr(*yes) +
     dspinput(*no) case(*mixed) prompt('Password')
