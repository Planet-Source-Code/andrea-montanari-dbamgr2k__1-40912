--------------------------------------------------------------------------------
DbaMgr2k  -       An alternative, little Console for MSDE 2.0
--------------------------------------------------------------------------------
Foreword.
i come from the Jet engine world, from the one part of VB3, but i never had
a copy of Access, what i needed to do was done by vb code and Visual Data
Manager.
Msde2000, like his brother got a major restyling, in configuration and admin
tasks to meet, or reach, TCO=0 needs (Total Cost of Ownership =0).
The bad thing is that MS does not include any visual tool to manage
Msde 2000 installation, only the powerfull characterbased oSql.exe, a tool
not to play with and not for everyone, at least not for my usual customers...
this brings me to a nice situation.... if my customer do have a Dba among
his employees, no problem, perhaps he owns the major Sql engine, but what
if no one is a Dba or a Sa?? can i afford and trust him to type obscure
T-Sql commands to perform simply tasks like Backup/restore or similar?
This project is brother of the preceding DbaMgr project, born in 1999, to
manage only Microsoft SqlServer 7.0 and Msde1.0 solution, project available
in the my original web site too.
This idea comes to me by reading Peter Swaniker article (vbpj july 1999
"Managing Sql Server 7.0"), that unfortunately was implemented on the NOT
redistributable Com Namespace object, so only part of frmMain is him due,
while i'm responsible of the rest... =;-(
it does not claim to be bug-free, nor to be the 'panacea', but could be of
interest for someone..

Like it's brohter, DbaMgr, this project was entirely written in VB6 sp5,
fully based on SqlServer2000 SQL-DMO, the com object that implements the
programming and maintenance of the engine and dependent/created objects.
It was still written in VB6 instead of .Net languages, becouse of lack of a
native .Net Sql_DMO component or similar to it, developing time woul'd have
exceded to much and .Net framework's Interop call to Sql-DMO woul'd have been
a very big overhead.
Needs win common controls, RichEdit, MSFlexGrid, Ms DataGrid and Mdac 2.6.
The project has been vouluntary split from DbaMgr, becouse DbaMgr2k
implements some version "2" of Sql-DMO interfaces available only on Sql
Server 2000, and it woul'd be not possible to be re-compiled from people
with only SqlServer 7.0 available.

It will not be an alternative to Enterprise Manager, it's simple and
primitive, but can be usefull for simple installation of MSDE 2.0 for the
common housekeeping tasks..., i guess not so simple installation would
requirea DBA and use of Enterprise Manager anyway.

Two other little utilities are integrated in the project:
the little query interface, even if very poor, let you split commands with
GO statements and retrieve multiples QueryResults (Resultsets).

BCP; a visual interface to support Bulk Copy operations from/to Tables/Views,
with choices for import/export data to native, character,comma-delim. and
custom format, with manipulation of the referenced format file and support
for Sql 6.x format (not tested, don't have Sql6.x); supports for destination
table definition and/or pre-written formatfile field definition;
for sure the batch possibility of BCP.exe can accomplish more difficult
tasks, but simple ops. can be managed visually very well.

In addition (and differently from SqlServer 2000), DbaMgr2k can manage more
extended properties for objects (database object included) that support this
features; extended properties values will always be saved as text.

It's not very well documented, you know, i'll do it later...

I miss a deep debug test, Mantainance Plan management, View modification
(possibly visual), DTS, etc.

English speeking people can immediately translate the tool from italian
(default) language to English (my poor english <g>) simply switching to the
General Setting ("Settaggi Generali") Tab  of the login form and setting
the Default Language ("Linguaggio Predefinito") comboBox to English;
DbaMgr2k can be localized with some work on Languages files ".Ln2" that must
reside in the application folder; DbaMgr2k autogenerate, if not present,
Italian and English language files, and, copying and renaming one of it to
ie.: Deutsch.Ln2 for German language, translating each row's content to the
new language, the new localization will than appears in the settings combo
box of the conection form.


I'm not an Interface Fan, i'm relative spartan..


Enclosed you'll find one installation script, a .ISS InnoSetup 2.0.18/19
version file, please look at BOL for dependencies for SQL-DMO anyway.


Someone told me they had interface reference troubles with some dmo
propertyes, i personally tested it against
msde 2000 sp2
sqldmo.dll      ver file 2000.11.19     ver 2000.08.00.534

and retail version of Msde2000
NB..
at
http://www.imranweb.com/freesoft/
you can download a free version of WinSQL and Enhaced ISQL/w utilityes,
written for Sql6.5, but quite stable with Sql7/Msde 1.0 and Sql2000/Msde2000,
implementing powerfull (and free) tools.

It's suggested too to download a free utility "SqlUtil" (SqlUtil2k.exe) from
www.innovmgtsys.com.

===========================================================================
If you want to be notified of new releases of DbaMgr2k, send me an e-mail and
you can subscribe to this service.
The service is free of charge and coul'd be interrupted by me at any time
with no previuos information.
===========================================================================

----- DISCLAIMER ---------------------------------------------------------------
All the material included in this compressed archive is FREEWARE;
you can use and distribute it without restriction, on condition
that the copyright and author informations are maintained unchanged.

You are entitled to bundle (with no extra cost) this software with software/
application of yours; shoul'd you do this way you have to inform the author
about this bundle and to provide information about your application.

This software is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; the author (me) has no responsibility for any damage or loss
this software could produce.
		USE IT AT YOUR OWN RISK
--------------------------------------------------------------------------------
Shoul'd you find bugs (and you will), suggestions improvements and so on,
please feel free contact me;
feed back, comments advices are welcome, not so insults... ;-)

enjoy.. and excuse my poor english..

Andrea Montanari
montanari_andrea@virgilio.it
http://utenti.lycos.it/asql/index.html
*************************
	HISTORY
*************************
0.2.0   xx/xx/2002
fPwd            added + and - chars;
frmLogin

bObjDMO         ErrExecSqlDirect(); modified parameter EXEC_TYPE from
                SQLDMOExec_ContinueOnError to SQLDMOExec_Default;
                (thank's to Allanon for the feedback);
--------------------------
0.1.0   19/11/2002
Initial Release alligned with DabMgr 0.47.0 version
END
