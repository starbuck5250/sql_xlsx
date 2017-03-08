/* SQL to XLSX (Excel) */
/* Uses Python in PASE */

pgm (&sqlstm &stmf &user &password)

dcl &script *char 64 value('/Opensource/sql_xlsx.py')
dcl &sqlstm *char 2560
dcl &stmf *char 512
dcl &user *char 10
dcl &user11 *char 11
dcl &password *char 10
dcl &password11 *char 11
dcl &x00 *char 1 value(x'00')

/* add C-like string terminator */
chgvar &script (&script *tcat &x00)
chgvar &sqlstm (&sqlstm *tcat &x00)
chgvar &stmf (&stmf *tcat &x00)
chgvar &user11 (&user *tcat &x00)
chgvar &password11 (&password *tcat &x00)

call qp2shell ('/QOpenSys/usr/bin/python' &script &sqlstm &stmf +
 &user11 &password11)

endpgm
