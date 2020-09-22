Attribute VB_Name = "ConstRes"
' ultima modifica
' last modified     : ver 0.1.0

Option Explicit

'Table Design Toolbar
Public Const keyEnd = "End"
Public Const keySave = "Save"
Public Const keyIndex = "Index"
Public Const keyTrigger = "Trigger"
Public Const keyPermission = "Permission"
Public Const keyDependencies = "Dependencies"
Public Const keyFitGrid = "FitGrid"
Public Const keyParam = "Param"
Public Const keyNew = "New"
Public Const keyDelete = "Delete"
Public Const keyProperty = "Property"
Public Const keyRefresh = "Refresh"
Public Const keyRun = "Run"
Public Const keyArrange = "Arrange"
Public Const keyExplorer = "Explorer"
Public Const keyStop = "Stop"
Public Const keyDefault = "Default"
Public Const keyWWrap = "WWrap"
Public Const keyHint = "Hint"
Public Const keyEProp = "eProp"

'immagini Res File
Public Const k_ResSelector = 101        'Flex Row selector
Public Const k_ResCombo = 102           'Flex Combo
Public Const k_ResChkOFF = 103          'Flex ChkOff
Public Const k_ResChkON = 104           'Flex ChkOn
Public Const k_ResPush = 105            'Flex Push
Public Const k_ResChkDel = 106          'Flex Deleted
Public Const k_ResUser = 107            'Flex User/Role
Public Const k_ResEllipsis = 108        'Flex push/edit

Public Const k_ResTB = 110              'Flex Table
Public Const k_ResView = 111            'Flex View
Public Const k_ResSP = 112              'Flex Sp
Public Const k_ResDef = 113             'Flex Default
Public Const k_ResRule = 114            'Flex Rule
Public Const k_ResUDT = 115             'Flex UDT
Public Const k_ResTRIGGER = 116         'Flex Trigger
Public Const k_ResPrimaryFile = 117     'Flex Primari DB File
Public Const k_ResFUNCTION = 118        'Funzioni Utente

Public Const k_ResColGrant = 123        'Column Granted
Public Const k_ResColDeny = 124         'Column Deny
Public Const k_ResColGrantDeny = 125    'Column Grant-Deny

'icone Res File
Public Const k_ResJoin = 101            'Join
Public Const k_ResLJoin = 102           'Left Join
Public Const k_ResRJoin = 103           'Right Join
Public Const k_ResFJoin = 104           'Full Join


'Stringhe Res File
Public Const k_Trusted_NT = 0                               ' Autenticazione Nt
Public Const k_Cancel = 1                                   ' Annulla
Public Const k_Connect = 2                                  ' Connetti
Public Const k_Select_Server = 3                            ' Seleziona Server
Public Const k_Time_Out_Not_Defined = 4                     ' Time Out Non Definito
Public Const k_mnuInclude_System_Objects = 5                ' Includi Oggetti di Sistema
Public Const k_Refresh = 6                                  ' Aggiorna
Public Const k_Confirm = 7                                  ' Confermate Operazione
Public Const k_Remove = 8                                   ' Elimina
Public Const k_End = 9                                      ' Esci
Public Const k_Full_Data_Path = 10                          ' Posizione Dati relativa all'Host per DB Scollegati
Public Const k_Select_Full_Data_Path = 11                   ' Seleziona Percorso Dati
Public Const k_Unknown = 12                                 ' Sconosciuto ..
Public Const k_Server_Unavailable = 13                      ' Server Non Disponibile
Public Const k_New_User = 15                                ' Nuovo Utente
Public Const k_Property = 16                                ' Proprietà
Public Const k_Confirm_Remove = 17                          ' Confermate Cancellazione
Public Const k_New_Table = 18                               ' Nuova Tabella
Public Const k_Show_all_Rows = 19                           ' Mostra Tutti i Record
Public Const k_Show_Max = 20                                ' Mostra Max ...
Public Const k_System_Object_not_removable = 21             ' Impossibile Eliminare Oggetto di Sistema
Public Const k_Executed = 22                                ' Eseguito
Public Const k_Text_Property = 23                           ' Proprietà Testo
Public Const k_Restore = 24                                 ' Ripristino
Public Const k_New_Stored_Procedure = 25                    ' Nuova Stored Procedure
Public Const k_Full_Data_PathToolTip = 26                   ' Posizione Dati per verifica presenza DB Scollegati
Public Const k_Only_DB_Users_has_access_to_this_function = 27   ' Dati disponibili solo per gli Utenti del DB
Public Const k_Only_DB_Owner_has_access_to_this_function = 28   ' Solo il Proprietario del DB ha accesso a questa Funzione
Public Const k_Save = 29                                    ' Salva
Public Const k_New_DB = 30                                  ' Nuovo DB
Public Const k_Detach_DB = 31                               ' Scollega DB
Public Const k_Shrink_DB = 32                               ' Compatta DB >>
Public Const k_Attach_DB = 35                               ' Collega DB
Public Const k_Extended_Property = 36                       ' Proprietà Estese
Public Const k_Extended_PropertyIdxTrig = 37                ' Proprietà Estese Indici..Trigger
Public Const k_Properties_for = 39                          ' Proprietà per
Public Const k_New_Login = 40                               ' Nuovo Login
Public Const k_Login_properties = 41                        ' Proprietà Login
Public Const k_New_Login_Properties = 42                    ' Proprietà Nuovo Login
Public Const k_New_BackUp_Device = 45                       ' Nuovo BackUp Device
Public Const k_Test_For = 46                                ' eseguito tentativo
Public Const k_Space_Gain_5 = 47                            ' recupero spazio 5%
Public Const k_With_NO_Space_Gain = 48                      ' liberazione spazio senza recupero
Public Const k_DBerror_wait_4_retry = 49                    ' Se viene ritornato errore di 'Database in uso', attendere 30/45 secondi dopo aver selezionato il DB da Collegare e riprovare
Public Const k_Not_restorable_Operation = 50                ' operazione non ripristinabile
Public Const k_Extended_PropertyIdxRel = 51                 ' Proprietà Estese Indici..Relazioni
Public Const k_Extended_PropertyTriggers = 52               ' Proprietà Estese Trigger

Public Const k_Obsolete_Sql_Version = 53               '
'ATTENZIONE....|É stata rilevata la versione # 1% # di Microsoft SqlServer.|La versione attualmente supportata da 2% è: # 3% #|Per versioni precedenti Utilizzare 4%.||2% Non può accedere a questo server.
'WARNING....|Version # 1% # of Microsoft SqlServer was detected.|SqlServer version currently supported by 2% is: # 3% #|For earlier version plase install 4%.||2% can not access this server.
Public Const k_GenWarning = 54                              ' Attenzione....
Public Const k_Not_Supported_Sql_Version = 55               '
' Attenzione....|É stata rilevata la versione # 1% # di Microsoft SqlServer.|La versione attualmente supportata da 2% è: # 3% #|non è stato testato l'uso con questa versione.||Non si consiglia l'uso di 2% con questo server.
' WARNING....|Version # 1% # of Microsoft SqlServer was detected.|SqlServer currently supported by 2% is: # 3% #|behaviour under this version was not tested.||It is strongly suggested not to use  2% with this server.

Public Const k_Rows = 56                                    ' Righe
Public Const k_Option_Setting = 57                          ' Settaggio Opzioni Database: 1% = 2%
Public Const k_Value = 58                                   ' Valore
Public Const k_Create_Date = 59                             ' Data Creazione
Public Const k_Last_BackUp = 60                             ' Ultimo BackUp
Public Const k_Users = 61                                   ' Utenti
Public Const k_Available_Space_Mb = 62                      ' Spazio Disponibile (Mb)
Public Const k_Num_of_Rows_to_return = 65                   ' Numero di Righe da Recuperare
Public Const k_Confirm_New_Password = 66                    ' Conferma Nuova Password
Public Const k_Password_confirmation_aborted = 67           ' Conferma Password non corrisponde
Public Const k_CompatLevel = 68                             ' Livello Compatibilità
Public Const k_New_DB_Role = 70                             ' Nuovo DB Role
Public Const k_Role_Name = 71                               ' Nome Ruolo
Public Const k_Drop = 72                                    ' Rimuovi
Public Const k_Database_Role_Properties = 73                ' Proprietà Ruolo DB:
Public Const k_Add_Role_Members = 75                        ' Aggiungi Membri Ruolo
Public Const k_Select_Users_to_Add = 76                     ' Selezionare Utenti da aggiungere
Public Const k_No_more_Users_4_Role = 77                    ' Non ci sono altri Utenti da aggiungere per questo Ruolo
Public Const k_Frm_Ok = 80                                  ' &Ok
Public Const k_Showing_Results_via_DMO = 88                 ' Presentazione Risultato via oggetto DMO
Public Const k_Abnormaly_Broken_Transaction = 90            ' Transazione Interrotta Abnormalmente: Verificare che gli Oggetti dipendenti (Trigger, Indici, Chiavi Primarie, Chiavi Esterne di Integrità Referenziale, Vincoli Check) risultino ancora esistenti
Public Const k_Committ_Structural_Changes = 91              ' Confermate Variazione Struttura?
Public Const k_Computed_Text = 94                           ' Testo Campo Calcolato
Public Const k_Computed_Fields = 95                         ' Campi Calcolati Contenuti:
Public Const k_Message_about_Tables = 96                    ' Messaggio di Avviso sulle Tabelle
Public Const k_Table_with_Computed_Fields_warning = 97      ' La Tabella in oggetto contiene Campi Calcolati; modifiche strutturali potrebbero invalidare i campi calcolati.
Public Const k_Confirm_Loose_of_Changes = 98                ' Confermate Annullamento Modifiche
Public Const k_Error_Executing_Table_Definition_for = 99    ' Errore in Esecuzione Definizione Tabella:
Public Const k_Select_File_Sql = 100                        ' Selezione File Sql
Public Const k_All_Files = 101                              ' Tutti i Files (*.*)
Public Const k_Save_Query = 102                             ' Salvataggio Query
Public Const k_New = 103                                    ' Nuovo
Public Const k_Execute_Query_F5 = 104                       ' Esegui Comando (F5)
Public Const k_ShowGrid_Text = 108                          ' Mostra Testo / Griglia
Public Const k_Format_File = 109                            ' File di Formato
Public Const k_Previous = 110                               ' < Indietro
Public Const k_Next = 111                                   ' Avanti >
Public Const k_Done = 112                                   ' Fine
Public Const k_Import = 113                                 ' Importazione
Public Const k_Export = 114                                 ' Esportazione
Public Const k_Selected_Object = 115                        ' Oggetto Selezionato:
Public Const k_Collected_All_Information = 116              ' Raccolta Informazioni Terminata
Public Const k_Select_Database = 117                        ' Selezionare Database
Public Const k_Select_Table_or_View = 118                   ' Selezionare Tabella o Vista
Public Const k_Select_File_for = 119                        ' Selezionare File di
Public Const k_Error_File = 120                             ' File di Errore
Public Const k_Error_FormatFile = 121                       ' Creazione File Formato
Public Const k_Qry_OpenQry = 122                            ' Apri Query da File
Public Const k_Qry_Delete = 123                             ' Cancella Testo Query
Public Const k_Qry_Opt_Reset = 124                          ' Resetta Opzioni Default

Public Const k_Lock_Type = 140                              ' Tipo di Lock
Public Const k_Pessimistic_Lock_default = 141               ' Lock Pessimistico (default)
Public Const k_Optimistic_Lock = 142                        ' Lock Ottimistico
Public Const k_Help_File = 143                              ' File Guida
Public Const k_Help_File_Browse = 144                       ' Seleziona File Guida
Public Const k_Beginning_Data_Transfer = 150                ' Inizio Trasferimento
Public Const k_Transferred_Rows = 151                       ' Righe Trasferite:
Public Const k_Error_reading_System_Handle = 155            ' Errore nella lettura Handle di Sistema:
Public Const k_Column_Name = 160                            ' Nome Colonna
Public Const k_Column_Type_modif_if_compatible = 161        ' Tipo Colonna, modificabile dalla selezione purchè compatibile
Public Const k_Actual_Column_Size = 162                     ' Dimensione Attuale Colonna
Public Const k_Column_Size_to_get_only_CHAR = 163           ' Dimensione Obiettivo Colonna (solo per tipi *CHAR compatibili)
Public Const k_Field_Terminator_4_BCP = 164                 ' Terminatore Campo (\t =TAB, \n =ACAPO, \r =RIT.CARRELLO, \\ =BACKSLASH, \0 =Term.NULLO, VUOTO per Nessuno)
Public Const k_Column_Position_in_current_T_V = 165         ' Posizione Colonna nella Tabella/Vista attuale
Public Const k_EXTERNAL_Format_File_Definition = 166        ' (Definizione Formato ESTERNO)"
Public Const k_Column_Type = 167                            ' Tipo Colonna

Public Const k_Type = 170                                   ' Tipo
Public Const k_Dimension = 171                              ' Dimensione
Public Const k_Dimension_to_reach = 172                     ' Dimensione Obiettivo
Public Const k_Connection = 180                             ' Connessione
Public Const k_Use_this_connection = 181                    ' Usa questa Connessione
Public Const k_Direction = 182                              ' Direzione
Public Const k_Table_View_selection = 183                   ' Selezione della Tabella/Vista
Public Const k_Data_File_Name = 184                         ' Nome File Dati
Public Const k_Max_Error_Number = 185                       ' Max Numero Errori
Public Const k_First_Row = 186                              ' Prima Riga
Public Const k_Last_Row = 187                               ' Ultima Riga
Public Const k_Batch_Size = 188                             ' Larghezza Batch
Public Const k_Preserve_Identity = 189                      ' Mantieni valori Identity
Public Const k_Compatible_SQL_6x = 190                      ' Compatibile SQL 6.x
Public Const k_IO_Format = 191                              ' Formato Input/Output
Public Const k_Load_Data_Definition_Format_from_Table = 192 ' Carica Definizione Formato da Tabella
Public Const k_Error_File_Empty_for_no_file = 193           ' File Errori (Vuoto per nessun File)
Public Const k_Hint = 194                                   ' Indicazioni:
Public Const k_Transfer_Status = 195                        ' Stato Trasferimento
Public Const k_Name = 200                                   ' Nome
Public Const k_Contents = 201                               ' Contenuto
Public Const k_Custom = 202                                 ' Personalizzato
Public Const k_Native = 203                                 ' Nativo
Public Const k_Character = 204                              ' Carattere
Public Const k_CommaDel = 205                               ' Separato "," CSV
Public Const k_Separator = 206                              ' Separatore
Public Const k_LoginName = 209                              ' Nome Login
Public Const k_Default_Connection_Information = 210         ' Impostazioni di Accesso
Public Const k_Default_Database = 211                       ' Database predefinito
Public Const k_Language = 212                               ' Lingua
Public Const k_Server_Role_granted = 213                    ' Ruoli assegnati al Login
Public Const k_Access_granted_to_Database_for_this_Login = 214  ' Database consentiti a questo Login
Public Const k_Database_Roles_for = 215                     ' Ruoli Database per
Public Const k_Permit = 216                                 ' Accesso
Public Const k_Permit_in_Database_role = 217                ' Consenti come
Public Const k_Login_Server_Roles = 218                     ' Ruoli Server
Public Const k_Login_Database_Access = 219                  ' Accesso ai Database
Public Const k_Indexes_caption = 220                        ' &Indici
Public Const k_Table_Design = 221                           ' Struttura Tabella
Public Const k_Indexes_Management = 222                     ' Gestione Indici
Public Const k_In_Primary_Key = 225                         ' In Chiave Primaria
Public Const k_Field_Name = 226                             ' Nome Campo
Public Const k_Data_Type = 227                              ' Tipo Dati
Public Const k_Size = 228                                   ' Dimensione
Public Const k_Allow_Null = 229                             ' Ammetti Null
Public Const k_Default_Value = 230                          ' Valore Predefinito
Public Const k_Precision = 231                              ' Precisione
Public Const k_Scale = 232                                  ' Scala
Public Const k_Is_RowGuid = 233                             ' Campo RowGuid
Public Const k_Is_Identity = 234                            ' Campo Identità
Public Const k_Initial_Value = 235                          ' Inizio
Public Const k_Increment = 236                              ' Incremento
Public Const k_Computed_Field = 237                         ' Campo Calcolato
Public Const k_Uncommitted_Structure_Changes = 239          ' Modifiche Struttura Non Salvate
Public Const k_Trigger_for_Table = 240                      ' Trigger per Tabella
Public Const k_Trigger_for_View = 241                       ' Trigger per Vista

Public Const k_Limit_of = 242                               ' Raggiunto limite di
Public Const k_Columns_reached_for_Table = 243              ' Colonne per Tabella:
Public Const k_Column_can_not_be_dropped = 244              ' Impossibile rimuovere Colonna:
Public Const k_Precision_value_must_be_between_1_and_28 = 245   ' L'impostazione di Precisione deve essere compresa tra 1 e 28
Public Const k_Scale_value_must_be_between_0_and_28 = 246   ' L'impostazione di Scala deve essere compresa tra 0 e 28
Public Const k_for_Table = 247                              ' per Tabella
Public Const k_Adding_Column = 248                          ' Aggiunta Campo
Public Const k_Object_already_exist_with_same_name = 249    ' Oggetto già esistente con stesso nome
Public Const k_ADMIN_Login_requested = 250                  ' Necessario Login come ADMIN
Public Const k_OBJ_Create_As_DBO = 251                      ' Crea Oggetto come DBO
Public Const k_OBJ_Create_As_DBO_tolTip = 252               ' Oggetto da creare con Proprietario DBO
Public Const k_for_View = 253                               ' per Vista
Public Const k_Insert_Field = 255                           ' Inserisci Campo
Public Const k_Table_Creation_Not_Allowed = 260             ' Impossibile Creare Tabella
Public Const k_Table_Modification_Not_Allowed = 261         ' Impossibile Modificare Tabella
Public Const k_Primary_Key = 265                            ' Chiave Primaria
Public Const k_Err_Importing_Extended_Properties_4_Tb = 270  ' Errori in RE-Importazione Proprieta' Estese per Tabella [1%]
Public Const k_Err_Importing_Extended_Properties_4_Tb_Desc = 271  ' La Tabella [1%] è stata correttamente rigenerata e salvata,|ma la RE-Importazione delle Proprietà Estese ha generato i seguenti errori;|le relative Proprietà Estese non sono state reimportate.||
' Table [1%] was succesfully saved and regenerated,|but the foolowing errors occured while RE-Importing Extended Properties;|error's related Extended Properties have been not reimported.||

Public Const k_Rules = 278                                  ' Regole
Public Const k_Rule = 279                                   ' Regola
Public Const k_Indexes = 280                                ' Indici
Public Const k_Check_Constraints = 281                      ' Vincoli Check
Public Const k_Relations = 282                              ' Relazioni
Public Const k_Indexes_Keys_constraints = 283               ' Indici/Chiavi
Public Const k_Constraint_Expression = 284                  ' Espressione Vincolo
Public Const k_Primary_Key_Table = 285                      ' Tab. Chiave Primaria
Public Const k_Indexed_Columns = 286                        ' Colonne Indicizzate
Public Const k_Constraint = 287                             ' Vincolo
Public Const k_Not_refresh_statistics = 288                 ' Non Aggiorna Statistiche
Public Const k_DRI_Key = 289                                ' Chiave DRI
Public Const k_New_Check_Constraint = 290                   ' Nuovo vincolo Check
Public Const k_New_Relation = 291                           ' Nuova Relazione
Public Const k_New_Index = 292                              ' Nuovo Indice
Public Const k_Descending = 293                             ' Discendente
Public Const k_Check_existing_data = 295                    ' Verifica Dati Esistenti durante creazione
Public Const k_Activate_Relation_for_Replica = 296          ' Attiva Relazione per Replica
Public Const k_Activate_Constraint_for_INSERT_and_UPDATE = 297  ' Attiva Vincolo per INSERT e UPDATE
Public Const k_Activate_Constraint_for_Replica = 298        ' Attiva Vincolo per Replica
Public Const k_Activate_Relation_for_INSERT_and_UPDATE = 299    ' Attiva Relazione per Insert e Update
Public Const k_Columns_not_defined = 300                    ' Campi Non Definiti
Public Const k_Update_Cascade = 301                         ' Update Cascade
Public Const k_Delete_Cascade = 302                         ' Delete Cascade
Public Const k_ModifySQL = 309                              ' Mod. SQL
Public Const k_Frm_ModifySQL = 310                          ' Modifica Sintassi SQL di creazione oggetto
Public Const k_Tables_not_defined = 311                     ' Tabelle non definite
Public Const k_Check_not_defined = 312                      ' Vincolo Check non definito



Public Const k_Connection_Propertyes = 329                  ' Proprietà di Connessione
Public Const k_Settings = 330                               ' Settaggi Generali
Public Const k_DbaMGR_Language = 331                        ' Linguaggio 1%
Public Const k_Hide_Non_Granted_DB = 332                    ' Nascondi DB non autorizati
Public Const k_Seconds_to_wait_for_Detach = 333             ' Secondi di Attesa x Scollegamento
Public Const k_Security_Mode = 334                          ' Tipo Sicurezza
Public Const k_Audit_Level = 335                            ' Tipo Audit
Public Const k_Detached_Databases = 339                     ' Database Scollegati
Public Const k_Database_Property = 340                      ' Proprietà Database
Public Const k_File_Property = 341                          ' Proprietà File
Public Const k_Automatic_Growth = 342                       ' Crescita Automatica
Public Const k_File_Growth = 343                            ' Crescita File
Public Const k_Max_File_Growth = 344                        ' Max Dimensione File
Public Const k_Unlimited = 345                              ' Illimitata
Public Const k_DB_is_OffLine = 348                          ' Database 1% è OffLine
Public Const k_DB_Name_for_reattach = 349                   ' Nome DB da Collegare
Public Const k_File_Name = 350                              ' Nome File
Public Const k_Location = 351                               ' Posizione
Public Const k_Space_Allocated = 352                        ' Spazio Allocato
Public Const k_Initial_Allocation = 353                     ' Allocazione Iniziale
Public Const k_LocateDB_file = 354                          ' Individuare Percorso File Database
Public Const k_LocateLog_file = 355                         ' Individuare Percorso File Log

Public Const k_RefreshFromSource = 357                      ' Aggiorna Dati da Origine
Public Const k_RemoveAddedFile = 358                        ' Rimuovi Nuovo File
Public Const k_AddNewFile = 359                             ' Aggiungi Nuovo File
Public Const k_Description = 360                            ' Descrizione
Public Const k_Add = 361                                    ' Aggiungi
Public Const k_Destination_Disk = 362                       ' Destinazione: Disk
Public Const k_Add_to_Media = 363                           ' Aggiungi al Media
Public Const k_Insert_BackUp_Name = 364                     ' Inserire Nome del BackUp
Public Const k_Insert_BackUp_Destination = 365              ' Inserire Destinazione BackUp
Public Const k_BackUp_File_Location = 366                   ' Posizione File BackUp
Public Const k_Options = 367                                ' Opzioni
Public Const k_Verify_BackUp = 368                          ' Verifica BackUp
Public Const k_Remove_old_data_from_Transaction_Log = 369       ' Rimuovi Voci inattive dal Log Transazioni
Public Const k_Verify_MediaSet_Name_and_Expiration_Date = 370   ' Verifica Nome MediaSet e Data Scadenza
Public Const k_Media_Set_Name = 371                         ' Nome Media Set
Public Const k_BackUp_Set_Expires = 372                     ' BackUp Set Scade:
Public Const k_In_Days = 373                                ' tra Giorni:
Public Const k_On = 374                                     ' Il:
Public Const k_BackUp_Destination = 375                     ' Destinazione BackUp
Public Const k_Server_File_System_PathToolTip = 377         ' File BackUp selezionato sul File System del Server
Public Const k_This_Device_does_not_contain_any_BackUp_sets = 379       ' Il Device selezionato non contiene alcun BackUp Set
Public Const k_Restore_as_DB = 380                          ' Ripristina come DB
Public Const k_Parameter_Restore_from_Device = 381          ' Parameter - Restore da Device
Public Const k_BackUp_Number = 382                          ' Numero BackUp
Public Const k_Force_Restore_to_overwrite_existing_DB = 390     ' Imponi ripristino sul Database esistente
Public Const k_Let_DB_operational_no_other_Transaction_log_Restore = 391    ' Imposta DB operativo - nessun altro Transaction log Restore successivo
Public Const k_Let_DB_NOT_operational_other_Transaction_log_follow = 392    ' Imposta DB NON operativo - altri Transaction log Restore seguono
Public Const k_Clear_BackUp_History = 395                   ' Pulisci BackUp History
Public Const k_Till_Date = 396                              ' Fino alla Data
Public Const k_Generate_SQL_Script = 400                    ' Genera SQL Script..
Public Const k_General = 401                                ' Generale
Public Const k_Formatting = 402                             ' Formattazione
Public Const k_Add_All = 404                                ' Aggiungi >>
Public Const k_Remove_All = 405                             ' << Rimuovi
Public Const k_All_Objects = 406                            ' Tutti gli Oggetti
Public Const k_All_Tables = 407                             ' Tutte le Tabelle
Public Const k_All_Views = 408                              ' Tutte le Viste
Public Const k_All_Stored_Procedures = 409                  ' Tutte le Stored Proc
Public Const k_All_Defaluts = 410                           ' Tutti i Defaults
Public Const k_All_Rules = 411                              ' Tutti i Rules
Public Const k_All_User_Defined_DataTypes = 412             ' Tutti i Tipi Definiti Utente
Public Const k_All_User_Defined_Functions = 413             ' Tutte le Funzioni Utente
Public Const k_Scripting_Opt_how_2_scrip = 414              ' Opzioni di Script per specificare come un oggetto sarà trattato
Public Const k_Generate_CREATE_DATABASE_command = 415       ' Genera comando CREATE <DATABASE>
Public Const k_Generate_DROP_Object_command = 416           ' Genera comando DROP <oggetto>
Public Const k_Generate_Script_for_all_dependent_objects = 417          ' Genera Script per oggetti dipendenti
Public Const k_Include_descriptive_headers_in_the_script_file = 418     ' Includi Testata descrittiva nello script
Public Const k_Include_Extended_Properties = 419            ' Includi Proprietà Estese
Public Const k_Include_Only_Ver_7_compliants = 420          ' Solo caratteristiche compatibili Versione 7.0
Public Const k_Script_DB_Users_Roles = 423                  ' Script Utenti e Ruoli Database
Public Const k_Script_SQL_Server_Logins_Windows_NT_and_SQL_Server_Logins = 424  ' Script SQL Server Logins (Windows NT e SQL Server Logins)
Public Const k_Script_Object_Level_Permission = 425         ' Script Permessi a livello Oggetto
Public Const k_Script_Indexes = 426                         ' Script Indici
Public Const k_Script_Full_Text_Indexes = 427               ' Script Indici Full Text
Public Const k_Script_Triggers = 428                        ' Script Trigger
Public Const k_Script_PKs_Foreign_Keys_Defaults_and_Check_Constraints = 429   ' Script PK, Foreign Key, Defaults e vincoli Check
Public Const k_Script_OutOfMemory = 430                     ' Caricamento Griglia incompleto terminato all'elemento # 1% #;|mancano ancora # 2% # elementi non visualizzabili.|Non sarà possibile gestire individualmente questi oggetti.
Public Const k_Ownership = 431                              ' Proprietà
Public Const k_File_Format = 435                            ' Formato File
Public Const k_MS_DOS_Text_OEM = 436                        ' Testo MS -DOS (OEM)
Public Const k_Windows_Text_Ansi = 437                      ' Testo Windows (Ansi)
Public Const k_International_Text_Unicode = 438             ' Testo Internazionale(Unicode)
Public Const k_Files_to_Generate = 440                      ' Files da Generare
Public Const k_Create_One_File = 441                        ' Crea Un File
Public Const k_Create_One_File_per_Object = 442             ' Crea Un File per Oggetto
Public Const k_Save_As = 450                                ' Salva con Nome
Public Const k_Save_Scripts_in_Directory = 451              ' Salva Script nella Directory
Public Const k_Unavailable_for_System_Objects = 452         ' Non concesso per Oggetti di Sistema
Public Const k_File_Size = 480                              ' Dimensione File
Public Const k_Created = 481                                ' Data Creazione
Public Const k_Last_Modified = 482                          ' Ultima Modifica
Public Const k_Last_Access = 483                            ' Ultimo Accesso
Public Const k_Read_Only = 484                              ' Solo Lettura
Public Const k_Archive = 485                                ' Archivio
Public Const k_Permissions = 500                            ' Autorizzazioni
Public Const k_Manage_Permissions = 501                     ' Gestione Autorizzazioni
Public Const k_List_all_Users_user_defined_DB_Roles_public = 502    ' Mostra Tutti gli Utenti / Ruoli DB / Public
Public Const k_List_only_Users_user_defined_DB_Roles_public_with_permission_on_this_object = 503    ' Mostra solo gli Utenti / Ruoli DB / Public con Permessi su questo oggetto
Public Const k_List_all_objects = 504                       ' Mostra tutti gli Oggetti
Public Const k_List_only_Objects_with_Permissions_for_this_User = 505   ' Mostra solo gli Oggetti con privilegi per questo Utente
Public Const k_List_only_Objects_with_Permissions_for_this_Role = 506   ' Mostra solo gli Oggetti con privilegi per questo Ruolo
Public Const k_PrivColumns = 507                            ' Colonne..
Public Const k_Apply = 508                                  ' Applica
Public Const k_Users_DB_Roles_public = 509                  ' Utenti / Ruoli DB / public
Public Const k_Table = 510                                  ' Tabella
Public Const k_View = 511                                   ' Vista
Public Const k_Stored_Procedure = 512                       ' Stored Procedure
Public Const k_User = 513                                   ' Utente
Public Const k_Object = 514                                 ' Oggetto
Public Const k_Owner = 515                                  ' Proprietario
Public Const k_Non_Valid_Option_for_DbOwner = 516           ' Opzione Non Valida per il Proprietario
Public Const k_ObjDefault = 517                             ' Default
Public Const k_Database_Role = 518                          ' Ruolo Database
Public Const k_User_Defined_Data_Type = 519                 ' Tipo Definito Utente

Public Const k_Priv_Col_Autorization = 530                  ' Autorizzazioni Colonne..
Public Const k_Priv_Col_OnlyAut_Col = 531                   ' Elenca solo Colonne autorizzate per questo Utente
Public Const k_Priv_User_Name = 532                         ' Nome Utente
Public Const k_Priv_Object_Name = 533                       ' Nome Oggetto
Public Const k_Priv_Error_Executing = 534                   ' Esecuzione ...



Public Const k_New_User_Defined_Data_Type = 550             ' Nuovo Tipo Definito dall'Utente
Public Const k_User_Defined_Data_Type_Properties = 551      ' Tipo Definito dall'Utente - Properietà:
Public Const k_Where_Used = 552                             ' Usato da...
Public Const k_the_following_columns_use_this_UDT = 553     ' le seguenti colonne usano questo tipo definito dall'utente
Public Const k_UDT_Collation_Not_Applicable = 554           ' Proprietà non applicabile

Public Const k_Dependencies_for = 560                       ' Dipendenze per
Public Const k_Objects_obj_THAT_Depends_on = 561            ' Oggetti che dipendono DA
Public Const k_Objects_THAT_obj_Depends_on = 562            ' Oggetti dai quali % Dipende
Public Const k_Object_Owner = 563                           ' Oggetto (Proprietario)
Public Const k_Object_Sequence = 564                        ' Sequenza
Public Const k_Privil_OutOfMemory = 565                     'Caricamento Griglia incompleto terminato all'elemento # 1% #;|mancano ancora # 2% # elementi non visualizzabili.|L'appropriata gestione è effettuabile solo tramite l'interfaccia di Query|con comandi T-Sql adeguati quali GRANT, DENY and REVOKE|oppure tramite gestione privilegi diretti dell'oggetto.
                                                            'Grid loading interrupted at element # 1% #;|other # 2% # elemnts are missing.|You can only manages these elements via Query Interface|with adeguate T-Sql commands like GRANT, DENY and REVOKE|or via direct privileges management of each object.
                                                            
Public Const k_Dropping_Objects = 579                       ' Cancellazione Oggetti....
Public Const k_Drop_Objects = 580                           ' Cancella Oggetti
Public Const k_Detach_Databases = 581                       ' Scollega Database
Public Const k_DeleteSPID = 582                             ' Termine spid (processo) SqlServer
Public Const k_Show_Dependencies = 599                      ' Mostra Dipendenze
Public Const k_New_View = 600                               ' Nuova Vista...
Public Const k_Save_View = 601                              ' Salva Vista..
Public Const k_Show_Hide = 603                              ' Mostra/Nascondi
Public Const k_Diagram_Pane = 604                           ' Pannello Diagrammi
Public Const k_Grid_Pane = 605                              ' Pannello Griglia
Public Const k_Sql_Pane = 606                               ' Pannello Sql
Public Const k_Result_Pane = 607                            ' Pannello Risultati
Public Const k_Run = 608                                    ' Esegui
Public Const k_Verify_Sql = 609                             ' Verifica Sql
Public Const k_Use_Group_By = 610                           ' Usa "Group By"
Public Const k_Clear_Result = 611                           ' Pulisci Risultati
Public Const k_Show_Table_View_List = 615                   ' Mostra Lista Tabelle/Viste
Public Const k_Close = 617                                  ' Chiudi
Public Const k_Tables = 618                                 ' Tabelle
Public Const k_Views = 619                                  ' Viste
Public Const k_Remove_Object_from_current_View = 620        ' Elimina Oggetto da Vista corrente
Public Const k_Column = 621                                 ' Colonna
Public Const k_Alias = 622                                  ' Alias
Public Const k_Table_ = 623                                 ' Tabella
Public Const k_Show = 624                                   ' Mostra
Public Const k_Grouping = 625                               ' Raggruppamento
Public Const k_Criteria = 626                               ' Criteri
Public Const k_Or = 627                                     ' Oppure
Public Const k_Data_in = 630                                ' Dati in:
Public Const k_Include_Row = 670                            ' Includi Righe:
Public Const k_All_Rows_from = 671                          ' Tutte le Righe da:
Public Const k_Join_Line = 672                              ' Linea Join
Public Const k_View_Name = 675                              ' Nome Vista
Public Const k_All_Columns = 676                            ' Tutte le Colonne
Public Const k_DISTINCT_Value = 677                         ' Valori DISTINCT
Public Const k_View_Encryption = 678                        ' Cifratura Vista
Public Const k_GROUP_BY_Extensions = 679                    ' Estensioni GROUP BY
Public Const k_Explicit_AnsiNull = 680                      ' ANSI Null on Esplicito
Public Const k_Explicit_QuotedIdentifier = 681              ' Quoted Identifier on Esplicito
Public Const k_Sql_Syntax_verified_successfully = 685       ' Sintassi Sql Verificata
Public Const k_Automatic_Arrange_Tables = 686               ' Disponi Tabelle Automaticamente
Public Const k_Fit_Grid = 687                               ' Organizza Griglia
Public Const k_Connect_Edit_Connection_Properties = 700     ' Connetti
Public Const k_Disconnect = 701                             ' Disconnetti
Public Const k_SqlServer_Connection_Properties = 705        ' Proprietà Connessione SqlServer
Public Const k_Reconnect = 706                              ' Riconnetti
Public Const k_RES_e_mail = 710                             ' e-mail contact
Public Const k_RES_www = 711                                ' www address
Public Const k_RES_FitGrid = 712                            ' Fit Grid
Public Const k_RES_Management = 720                         ' Amministrazione
Public Const k_RES_CurrentActivity = 721                    ' Attività Corrente
Public Const k_RES_Process_Info = 722                       ' Processi Attivi
Public Const k_PROC_KillSpid = 725                          ' Termina Processo

Public Const k_RES_Spid = 750                               ' spid
Public Const k_RES_User = 751                               ' User
Public Const k_RES_Database = 752                           ' Database
Public Const k_RES_Status = 753                             ' Status
Public Const k_RES_Command = 754                            ' Command
Public Const k_RES_Application = 755                        ' Application
Public Const k_RES_Cpu = 756                                ' Cpu
Public Const k_RES_MemUsage = 757                           ' Memory Usage
Public Const k_RES_Blocked = 758                            ' Blocked By spid
Public Const k_RES_Blocking = 759                           ' Blocking spid
Public Const k_RES_NetAddress = 760                         ' Network Address
Public Const k_RES_NetLib = 761                             ' Network Library
Public Const k_RES_Host = 762                               ' Host
Public Const k_RES_LastBatch = 763                          ' Last Batch
Public Const k_RES_LoginTime = 764                          ' Login Time
Public Const k_RES_IO = 765                                 ' Physical IO
Public Const k_RES_WaitType = 766                           ' Wait Type
Public Const k_RES_WaitTime = 767                           ' Wait Time
Public Const k_RES_openTran = 768                           ' Open Transaction



Public Const k_DbName = 780                                 ' Nome Database
Public Const k_Mode = 781                                   ' Modo di lock
Public Const k_Status = 782                                 ' Stato di lock
Public Const k_TableName = 783                              ' Nome Tabella
Public Const k_IndexName = 784                              ' Nome Indice


Public Const k_MixedSecurity = 790                          ' Sicurezza Mista
Public Const k_SqlSecurity = 791                            ' Sicurezza SqlServer
Public Const k_WinSecurity = 792                            ' Sicurezza Windows
Public Const k_Audit_None = 795                             ' Nessun Audit
Public Const k_Audit_Success = 796                          ' Registra Successi
Public Const k_Audit_Failure = 797                          ' Registra Errori
Public Const k_Audit_All = 798                              ' Registra Tutto

Public Const k_SqlProperty = 800                            ' Proprietà SqlServer - [1%]
Public Const k_ApplyChanges = 802                           ' Applica Variazioni
Public Const k_GeneralTab = 805                             ' Generale
Public Const k_SecurityTab = 806                            ' Sicurezza
Public Const k_ConnectionTab = 807                          ' Connessioni
Public Const k_ServerSettingsTab = 808                      ' Settaggi Server
Public Const k_DMOTab = 809                                 ' Componenti Client
Public Const k_ServerLanguage = 810                         ' Linguaggio Predefinito
Public Const k_AllowChanges = 811                           ' Permetti modifiche dirette su cataloghi di sistema
Public Const k_NestedTriggers = 812                         ' nested triggers
Public Const k_2yerCutoff = 813                             ' Data fino alla quale 2 cifre per l'anno sono preimpostate

Public Const kMaxConcurrentUser = 814
Public Const kConstraintCheck = 815
Public Const kImplicitTrans = 816
Public Const kCloseCursors = 817
Public Const kANSI_warn = 818
Public Const kANSI_pad = 819
Public Const kANSI_nulls = 820
Public Const kArit_aborth = 821
Public Const kArit_ignore = 822
Public Const kQuoted_ident = 823
Public Const kNO_count = 824
Public Const kNULLS_definedON = 825
Public Const kNULLS_definedOFF = 826

Public Const kDefaultDataRoot = 829                         'Directory Predefinita Dati e Log
Public Const kExecutedOverrides = 830

Public Const kODBCVersionString = 831
Public Const kGroupRegistrationServer = 832
Public Const kBlockingTimeout = 833
Public Const kFullName = 834
Public Const kVersion = 835
Public Const kFileVersion = 836
Public Const kProductVersion = 837
Public Const kServicePack = 838
Public Const kMdacVersion = 839

Public Const k_Param = 840                                  ' Paramentri di Start Up
Public Const k_Parameter = 841                              ' Parametro
Public Const k_Existing_Parameter = 842                     ' Parametri Esistenti
Public Const k_ParameterWarning = 845                       ' Warning Parametri


Public Const k_DB_PrimaryFilePath = 880                     ' Path File Primario

Public Const k_DBStatus_Normal = 885                        ' Normale
Public Const k_DBStatus_OffLine = 886                       ' Off Line
Public Const k_DBStatus_Recovering = 887                    ' In Recupero
Public Const k_DBStatus_StandBy = 888                       ' Stand By
Public Const k_DBStatus_Suspect = 889                       ' Sospetto
Public Const k_DBStatus_Inaccessible = 890                  ' Inaccessibile
Public Const k_DBStatus_UnKnown = 891                       ' Sconosciuto
Public Const k_DBStatus_dboUseOnly = 892                    ' ad Uso Solo DBO
Public Const k_DBStatus_ReadOnly = 893                      ' Sola Lettura
Public Const k_DBStatus_SingleUser = 894                    ' Singolo Utente
Public Const k_DBStatus_Loading = 895                       ' In Caricamento
Public Const k_DBStatus_Stand_By = 896                      ' In Stand by

Public Const k_Detach_Of_DB = 900                           ' Scollegamento Database [1%]
Public Const k_ReAttach_Of_MultiFile_DBMSG = 901            ' spiegazione di sp_attach_db

Public Const k_tab_Files_Selection = 908                    ' Selezione File(s) da includere
Public Const k_tab_Post_Attach_DB_Options = 909             ' Opzioni DB Post-Collegamento
Public Const k_Skipped_Detach_Of_DB = 910                   ' Scartata Operazione di Scollegamento per Database [1%]
Public Const k_Attach_DB_to_Server = 911                    ' Collegamento Database al Server
Public Const k_Select_Destination_File = 912                ' Selezione File Destinazione
Public Const k_Select_Data_File = 914                       ' Selezione File Dati-Log da Inserire
Public Const k_Exec_ReAttach = 915                          ' Esegui Collegamento DB
Public Const k_Attaching_DB_Name = 916                      ' Nome del DB da Collegare
Public Const k_Attaching_MSG = 917                          ' spiegazioni Riattacco DB
            'Prima di procedere, verificare che il Database [2%] non sia presente ed impostato OFF-LINE,|e che NON sia un Database Attivo di un'altra istanza di SqlServer.|Al fine di Attaccare al Server [1%] il Database [2%], è necessario elencare tutti i File Dati ed i File di Log appartenenti ad esso;|il File Dati principale ha usualmente estensione '.MDF', i file Dati successivi '.NDF' ed i file di Log '.LDF'.|Vengono preselezionati i File che dovrebbero costituire il set completo del Database.|Se si omettono 1 o più files, l'operazione non avrà esito positivo.
            'Before proceding, please check that Database [2%] is not currently Off-Line,|and that it is NOT an Active Database of another SqlServer Instance.|In order to Attach to Server [1%] the [2%] Database, you have list all Data and Log Files belonging to it;|Primary Data File usually has '.MDF' extention, while other Data Files have '.NDF' and Log Files '.LDF'.|The Files that are supposed to contitue the Database set are preselected.|Should 1 or more files be omitted, attach operation will abort.
Public Const k_DbDetacheRefWarning = 918                    ' Oggetto Database [ 1% ]... Warning - 2%
Public Const k_DbDetachedScanWarning = 919                  ' Riscontrati Errori durante la scansione di eventuali Database Scollegati.
Public Const k_DbDetached_PrimaryFileMoved = 920            ' Il file [ 1% ] è stato probabilmente spostato dalla posizione originale.|Verrà tentato comunque il recupero delle informazioni tramite l'attuale nome fisico e posizione.
Public Const k_DbDetached_NoDbFile = 921                    ' Il File selezionato non appare un Database Scollegato.
Public Const k_DbDetached_FileNotFound = 922                ' File del Database Set [ 1% ] non trovati nelle posizioni indicate:

Public Const k_DbOpt_AccessGrantedTo = 930                  ' Accesso consentito a
Public Const k_DbOpt_RecoveryModel = 931                    ' Modello di Recupero
Public Const k_DbOpt_AnsiNullDefault = 932                  ' Valore Predefinito ANSI NULL
Public Const k_DbOpt_RecursiveTriggers = 933                ' Trigger Ricorsivi
Public Const k_DbOpt_AutoCreateStat = 934                   ' Creazione Automatica Statistiche
Public Const k_DbOpt_AutoUpdateStat = 935                   ' Aggiornamento Automatico Statistiche
Public Const k_DbOpt_TornPageDetection = 936                ' Rilevamento pagine incomplete
Public Const k_DbOpt_AutoClose = 937                        ' Chiusura Automatica
Public Const k_DbOpt_AutoShrink = 938                       ' Compattazione Automatica     'Auto Shrink
Public Const k_DbOpt_UseQuotedIdentifier = 939              ' Usa Identificatori tra virgolette     'Use quoted identifier

Public Const k_DbOpt_Access_ALL = 945                       ' Tutti
Public Const k_DbOpt_Access_Dbo = 946                       ' Membri di db_owner

Public Const k_DbOpt_Recovery_Simple = 948                  ' Semplice
Public Const k_DbOpt_Recovery_Bulklogged = 949              ' Con Registrazioni di Massa
Public Const k_DbOpt_Recovery_Full = 950                    ' Completo


Public Const k_mnuSrvConfg = 980                            ' Utilità di Rete Sql Server
Public Const k_mnuCliConfg = 981                            ' Utilità di Rete del Client
Public Const k_RES_Object_Not_Found_simple = 982            ' Impossibile Referenziare Oggetto
Public Const k_mnuBCP = 983                                 ' BCP/Imp.Exp. Dati
Public Const k_mnuQuery = 984                               ' Query/interrogazioni
Public Const k_mnuActivity = 985                            ' Attività
Public Const k_mnuInfo = 986                                ' Informazioni
Public Const k_mnuLicense = 987                             ' Licenza Utente
Public Const k_mnuAbout = 988                               ' About
Public Const k_mnuEngineVersion = 989                       ' Versione Sql Server
Public Const k_ClientConnection = 990                       ' {Connessione Aperta Lato Client}
Public Const k_Error_Opening_Client_RS = 991                ' Errore nell'apertura RecordSet lato Client
Public Const k_RefreshClientConnectionData = 992            ' Aggiorna Dati dalla Fonte per Cursore Lato Client
Public Const k_ClientCursorWarning = 993                    ' Cursore Lato Client, Default non saranno visibile se non dopo un refresh dei dati
Public Const k_RsStateNotOpen = 994                         ' Stato RecordSet Non Aperto

Public Const k_RegenLngFilesDONE = 995                      ' File Linguaggio Default Rigenerati
Public Const k_mnuRegenLngFiles = 996                       ' Rigenera File Linguaggio Default
Public Const k_Exit = 997                                   ' Esci
Public Const k_RES_Invalid_Value = 998                      ' Valore Non Valido per Proprietà
Public Const k_RES_Object_Not_Found = 999                   ' Impossibile Referenziare Oggetto
Public Const k_mnuDependencies = 1000                       ' Dipendenze 1%
Public Const k_DependenciesNotFound = 1001                  ' File Dipendenze 1% non trovato
Public Const k_EulaNotFound = 1002                          ' File Licenza 1% non trovato
Public Const k_EulaWelcome = 1003                           ' File Licenza 1% non trovato

Public Const k_Search4Orphaned = 1100                       ' Ricerca Utenti Orfani per DB [1%]
Public Const k_Mapped = 1101                                ' Mappato
Public Const k_AssociatedLogin = 1102                       ' Login Assogiato
Public Const k_UserMap = 1103                               ' Utente
Public Const k_Drop_Orhaned_User = 1105                     ' Elimina Utente Orfano
Public Const k_Ask4Drop = 1106                              ' Volete veramente eliminare Utente [1%]
Public Const k_OrphanedWARNING = 1107

Public Const k_ModifyDB_Owner = 1120                        'Modifica Proprietario
Public Const k_Old_Obj_Owner = 1121                         'Proprietario Attuale
Public Const k_New_Obj_Owner = 1122                         'Nuovo Proprietario
Public Const k_Change_DB_Owner_Frm = 1123                   'Modifica Proprietario Database # 1% #
Public Const k_Modified_DB_Owner = 1124                     'Modificato Proprietario Database # 1% # |Nuovo Proprietario: # 2% #
Public Const k_DB_Owner_is_Same = 1125                      'Proprietario Originale e Nuovo Proprietario sono Uguali
Public Const k_Change_Obj_Owner_Frm = 1126                  'Modifica Proprietario Oggetti per Database # 1% #
Public Const k_Change_DbOwnerSame = 1127                    'Il Vecchio Proprietario # 1% # del Database # 2% # ed il Nuovo Proprietario sono uguali.|Continuare?


Public Const k_MainRelViewer = 1149                         ' Visualizzazione Relazioni
Public Const k_RelGrapficalView = 1150                      ' Visualizzazione Grafica Relazioni per Database [1%]
Public Const k_RelSearch = 1151                             ' Ricerca Relazioni
Public Const k_RelRecurviveSearch = 1155                    ' Ricerca Ricorsiva
Public Const k_RelRecurviveShow = 1156                      ' Mostra oggetti Dipendenti {Ricerca Ricorsiva}
Public Const k_RelManageIdx = 1157                          ' Gestione Indici per Tabella selezionata [1%]
Public Const k_RelSelectTable = 1158                        ' Selezionare Tabella dal ComboBox
Public Const k_RelSelectTB = 1160                           ' Selezione Oggetto Tabella
Public Const k_RelPaneMSG = 1161                            ' Selezionare Tabella per evidenziare campi Chiave e Referenziati
Public Const k_RelSelfRefer = 1164                          ' {AUTOREFERENZIANTE RICORSIVA}
Public Const k_RelOrigTbl = 1165                            ' Tabella Principale 1%
Public Const k_RelReferencedTblCols = 1166                  ' Campi Tabella Referenziata
Public Const k_RelReferencedTbl = 1167                      ' Tabella Referenziante 1%
Public Const k_RelShowDetails = 1168                        ' Mostra/nascondi Dettagli
Public Const k_RelName = 1170                               ' Relation 1%
Public Const k_RelReferencingTblCols = 1171                 ' Campi Tabella Referenziante

Public Const k_HtmDocum = 1200                              ' Database Documentation
Public Const k_HtmDocumInfo = 1201                          ' Generazione Documentazione per DB: [1%]
Public Const k_HtmLocation = 1202                           ' Directory di Destinazione
Public Const k_HtmIncludeSYSobj = 1203                      ' Includi Oggetti di Sistema

Public Const k_HtmShow = 1205                               ' Mostra Documentazione
Public Const k_HtmKillExport = 1208                         ' Interrompi Elaborazione..
Public Const k_HtmProducedBy = 1209                         ' Prodotto da 1% il 2%
Public Const k_HtmHome = 1210                               ' Pagina Iniziale


Public Const k_HtmFileGroup = 1220                          ' Gruppo di File
Public Const k_HtmFileName = 1221                           ' Nome File
Public Const k_HtmPhysicalName = 1222                       ' Nome Fisico
Public Const k_HtmFileGrowth = 1223                         ' Crescita File
Public Const k_HtmFileMaxSize = 1224                        ' Massima Dim.
Public Const k_HtmFileSizeKb = 1225                         ' Dimens.KB
Public Const k_HtmFileGrowthType = 1226                     ' Tipo Cresc.
Public Const k_HtmLogFile = 1230                            ' Log File

Public Const k_HtmTBDef = 1235                              ' Definizione Tabella
Public Const k_HtmColDef = 1236                             ' Dettagli per Colonna: [1@%]
Public Const k_HtmColList = 1237                            ' PK - FK referenzianti
Public Const k_HtmUNKNOWN = 1238                            ' SCONOSCIUTO
Public Const k_HtmPK = 1239                                 ' Chiave Primaria
Public Const k_HtmFK = 1240                                 ' Chiave Esterna
Public Const k_HtmUK = 1241                                 ' Chiave Unica

Public Const k_HtmPriTB = 1245                              ' Tabella Principale
Public Const k_HtmForTB = 1246                              ' Tabella Referenziante
Public Const k_HtmKType = 1247                              ' Tipo Chiave

Public Const k_HtmKClustered = 1250                         ' Clustered
Public Const k_HtmKFillFactor = 1251                        ' FillFactor
Public Const k_HtmKCheck = 1252                             ' Verifica Congruità
Public Const k_HtmKCcolumns = 1253                          ' Campi Chiave
Public Const k_HtmFKCcolumns = 1254                         ' Campi Chiave Referenziati
Public Const k_HtmChkText = 1255                            ' Testo Vincolo

Public Const k_HtmText = 1259                               ' Testo
Public Const k_HtmScript = 1260                             ' Script per la rigenerazione oggetto
Public Const k_HtmClearDir = 1261                           ' Cartella Esistente, Cancellare Contenuto
Public Const k_HtmDeleting = 1262                           ' Cancellazione ...
Public Const k_HtmWorking = 1263                            ' Operazione in corso ...."
Public Const k_HtmReady = 1264                              ' Operazione in Terminata in..
Public Const k_HtmKilled = 1265                             ' Operazione Interrotta..

Public Const k_HtmSpParameter = 1270                        ' Parametri
Public Const k_HtmViewColumns = 1271                        ' Campi referenziati
Public Const k_HtmBoundColumns = 1272                       ' Campi Collegati
Public Const k_HtmBoundUDT = 1273                           ' Tipi Definiti dall'Utente Collegati

Public Const k_HtmSysObj = 1275                             ' 1% di Sistema

Public Const k_HtmScanning = 1280                           ' scansione... 1%...
Public Const k_HtmScanDB = 1281                             ' Database
Public Const k_HtmScanTB = 1282                             ' Tabelle
Public Const k_HtmScanSP = 1283                             ' Stored Procedures
Public Const k_HtmScanV = 1284                              ' Viste
Public Const k_HtmScanUDT = 1285                            ' Tipi Definiti dall'Utente
Public Const k_HtmScanR = 1286                              ' Regole
Public Const k_HtmScanDef = 1287                            ' Default
Public Const k_HtmScanFunction = 1288                       ' Funzione Utente

Public Const k_QRY_SaveQry = 1350                           ' Salva Query
Public Const k_QRY_SaveResult = 1351                        ' Salva Risultato Query
Public Const k_QRY_Save_Done = 1355                         ' Salvataggio Risultato

Public Const k_Priv_Create_Table = 1360
Public Const k_Priv_Create_View = 1361
Public Const k_Priv_Create_SP = 1362
Public Const k_Priv_Create_Default = 1363
Public Const k_Priv_Create_Rule = 1364
Public Const k_Priv_Backup_DB = 1365
Public Const k_Priv_Backup_Log = 1366
Public Const k_Priv_CreateFunction = 1367


Public Const k_B4F_Ok = 1400                                ' Ok
Public Const k_B4F_FileName = 1401                          ' Nome File
Public Const k_B4F_SelFile = 1402                           ' File Selezionato
Public Const k_B4F_SelPath = 1403                           ' Percorso Selezionato
Public Const k_B4F_SqlIsNothing = 1405                      ' Sql Server Host non reperibile

Public Const k_B4F_TviewToolTip = 1406                      ' FileSystem dell'Host SqlServer # 1% #

Public Const k_B4F_Err_NoFileSelected = 1410                ' Nessun File Selezionato
Public Const k_B4F_Err_DirMustExists = 1411                 ' La Directory deve esistere
Public Const k_B4F_Err_FileMustExists = 1412                ' Il File deve esistere
Public Const k_B4F_Err_CantChangeDir = 1413                 ' La directory Selezionata deve essere # 1% #
Public Const k_B4F_Err_NoDirSelected = 1414                 ' Nessuna Directory Selezionata

Public Const k_ObjectOwnedby = 1500                         ' Oggetti Posseduti da [1%]
Public Const k_ModifyObjectOwner = 1501                     ' Modifica Proprietario
Public Const k_ChangingObjectOwner = 1510                   ' Cambiamento Proprietario Oggetto 1% : [2%] da [3%] a [4%]
Public Const k_ChangingImpossible = 1511                    ' Impossibile effettuare operazione
Public Const k_ChangingDone = 1512                          ' Operazione Eseguita
Public Const k_ChangingResult = 1513                        ' Risultato Modifica Proprietario
Public Const k_ChangingPreWarning = 1515                    ' ATTENZIONE...||La modifica della Proprietà degli Oggetti può comportare problemi di riferimento nelle applicazioni che utilizzano il database.||Solamente gli utenti del database con privilegi di "ddl_admin" e tutti le eventuali Login con privilegi "sys admin" verranno resi disponibili come candidati per il cambio di proprietà.||Qualora venisse scelto un Login (sysadmin) NON presente tra gli utenti correnti, verrà generato un nuovo Utente con stesso nome e privilegi di "ddl_admin".
                                                            ' WARNING...||Changing Object(s)'s Ownership can break application referencing objects of the database.||Only Database User's with "ddl_admin" privileges as long as "sys admin" Logins will be listed as available candidates.||Shoul'd a (sysadmin) Login, not part of current Db users be choosed as the new owner, a new user with the same name and "ddl_admin" privileges will be added to the users collection.
                                                            
Public Const k_SqlObjInvalid = 1516                         ' L'oggetto SqlServer Non è più attendibile/utilizzabile.|É necessario riconnettersi all'applicazione.
Public Const k_ChangingPreWarning2 = 1517                   ' ||Solo gli oggetti Non di Sistema appartenenti allo stesso proprietario saranno passati alla procedura di cambio proprietario.|Tutti gli eventuali altri oggetti saranno scartati.
                                                            ' ||Only Not System object(s) belonging to the same owner will be passed to the "Changing Object(s) Ownership" procedure.|All other object(s) will not be enlisted.

Public Const k_AddingChangingUser = 1518                    ' Aggiunta Utente [1%] alla Collezione Utenti Database # 2% #
Public Const k_Adding_DDL_User = 1519                       ' Abilitato Utente [1%]
Public Const k_AddingChangingLogin = 1520                   ' Aggiunta Login Temporaneo [ 1% ]
Public Const k_ChangingLoginRoles = 1521                    ' Conferimento privilegi [ 1% ] a Login Temporaneo [ 2% ]


Public Const k_DetachMode_Check = 1550                      ' Effettua CheckDB
Public Const k_DetachMode_NoCheck = 1551                    ' Salta CheckDB
Public Const k_DetachMode_Ask = 1552                        ' Richiedi per ogni DB
Public Const k_Detach_Setting = 1555                        ' Specifiche di Scollegamento
Public Const k_Detach_Warning = 1556                        ' É possibile modificare le specifiche di verifica dei Database preventive all'operazione di Scollegamento nel pannello di connessione.||L'attuale valore impostato è : # 1% #
Public Const k_Detach_Prompt = 1557                         ' Evitare CheckPoint per database # 1% #
Public Const k_Detach_Skipped_stbar = 1558                  ' Operazione Saltata
Public Const k_Detach_Skipped_warning = 1559                ' NON eseguito CheckPoint prima dell'operazione di Scollegamento.|Statistiche di supporto per l'ottimizzazione delle query aggiornate comunque prima dell'operazione di Scollegamento.
                                                            ' Skipped CheckPoint before Detach operation.|Statistics supporting query optimization are updated prior to the detach operation.
Public Const k_Detach_CheckDB_stbar = 1560                  ' Aggiornamento per DB: # 1% #
                                                            ' CheckPoint for DB: # 1% #

Public Const k_ServerAccess = 1580                          ' Accesso al Server
Public Const k_AccessDeny = 1581                            ' Negato
Public Const k_AccessPermit = 1582                          ' Permesso
Public Const k_Log_Authentication = 1585                    ' Autenticazione
Public Const k_Log_NTAuthentication = 1586                  ' Autenticazione Windows NT
Public Const k_Log_Domain = 1587                            ' Dominio
Public Const k_Log_Deny_Grant = 1588                        ' Nega Accesso\permetti accesso

Public Const k_sqlTbarResult = 1599                         ' Mostra Risultato
Public Const k_sqlMnuGenerateInsert = 1600                  ' Genera Insert Script
Public Const k_sqlFrmGenerateInsert = 1601                  ' Generazione Script INSERT per Tabella '1%' - Database : '2%'
Public Const k_sqlIsRowGuid = 1602                          ' Campo RowGuid
Public Const k_sqlIsTimeStamp = 1603                        ' Campo timestamp/rowversion
Public Const k_sqlInclude = 1604                            ' Includi
Public Const k_sqlReplaceNull = 1605                        ' Rimpiazza Valori NULL
Public Const k_sqlPosition = 1606                           ' Posizione Campo in Inserimento
Public Const k_sqlFieldAlias = 1607                         ' Nome Alias del Campo nella Tabella Destinazione
Public Const k_sqlFieldReplace4Null = 1608                  ' Rimpiazzo Valori NULL

Public Const k_sqlTB_opt_Delete = 1610                      ' Svuotamento Dati Tabella
Public Const k_sqlTB_opt_DeleteHelp = 1611                  ' Operazioni Preliminari di Svuotamento Dati Tabella
Public Const k_sqlTB_opt_DateFormatHelp = 1612              ' Le Date vengono restituite nel formato ODBC (121) yyyy-mm-dd hh:mi:ss.mmm
Public Const k_sqlTB_opt_NoCountHelp = 1613                 ' Imposta SET NOCOUNT ON
Public Const k_sqlTB_opt_IdentityHelp = 1614                ' Mantiene i Valori Identità come da origine
Public Const k_sqlTB_opt_IsolationHelp = 1615               ' Imposta Livello Isolamento Transazione
Public Const k_sqlTB_opt_BatchHelp = 1616                   ' Numero di Righe per Batch (0 = unico Batch)
Public Const k_sqlTB_opt_File = 1617                        ' File di Esportazione
Public Const k_sqlTB_opt_FileHelp = 1618                    ' Nome File di Esportazione
Public Const k_sqlTB_opt_Where = 1619                       ' Filtro WHERE
Public Const k_sqlTB_opt_WhereHelp = 1620                   ' Valida Istruzione WHERE (200 char, solo filtri senza WHERE)
Public Const k_sqlTB_opt_OrderBy = 1621                     ' Ordine Dati
Public Const k_sqlTB_opt_OrderByHelp = 1622                 ' Valida Istruzione ORDER BY (200 char, solo campi senza ORDER BY)
Public Const k_sqlTB_opt_TbAlias = 1623                     ' Alias Tabella
Public Const k_sqlTB_opt_TbAliasHelp = 1624                 ' Nome Tabella Alias Valido per Tabella Esportata (200 char)
Public Const k_sqlTB_opt_Top = 1625                         ' Top Righe
Public Const k_sqlTB_opt_TopHelp = 1626                     ' TOP Numero di Righe da ritornare (0 = tutte)
Public Const k_sqlTB_opt_TimeOut = 1627                     ' Time Out Connessione
Public Const k_sqlTB_opt_TimeOutHelp = 1628                 ' Tempo di TImeOut Connessione

Public Const k_sqlDlgFileEport = 1630                       ' File di Esportazione

Public Const k_sqlStatementGO = 1631                        ' Termine Batch
Public Const k_sqlTimeGenerated = 1632                      ' Generato il # 1% #
Public Const k_sqlInputTbl = 1633                           ' Tabella di Origine # 1% #
Public Const k_sqlInputDB = 1634                            ' Database di Origine # 1% #
Public Const k_sqlSettings = 1635                           ' Settaggi
Public Const k_sqlProp_Value = 1636                         ' Proprietà - Valore
Public Const k_sqlProp_Preliminary = 1637                   ' Operazioni Preliminarie
Public Const k_sqlProp_DestinationTable = 1638              ' Tabella Destinazione
Public Const k_sqlProp_BatchSize = 1639                     ' Grandezza Batch
Public Const k_sqlProp_PreserveIdentity = 1640              ' Preserva Campi Identity
Public Const k_sqlProp_TranIsolationLevel = 1641            ' Livello Isolamento Transazione
Public Const k_sqlAction_Deleting = 1645                    ' Svuotamento Tabella Destinazione
Public Const k_sqlAction_Loading = 1646                     ' Caricamento Dati
Public Const k_sqlAction_Break = 16147                      ' Job Interrrotto dall'utente
Public Const k_sqlAction_End = 1648                         ' Fine Caricamento Dati
Public Const k_sqlErrSqlStatement = 1650                    ' Comando SQL non Valido o non Riconosciuto
Public Const k_sqlErrAdoConnOpening = 1651                  ' Errori Apertura Connessione ADO

Public Const k_sqlWarning0 = 1655                           ' Siete pregati di voler verificare lo script.
Public Const k_sqlWarning1 = 1656                           ' Inoltre, provvedete ad effettuare il BackUp del Database prima dell'esecuzione.
Public Const k_sqlWarning2 = 1657                           ' Campi di tipo VARCHAR e IMAGE non sono esportabili per ovvi motivi.
Public Const k_sqlGettingResult = 1660                      ' Recupero Righe in Corso....
Public Const k_sqlWritingResult = 1661                      ' Scrittura Righe in Corso...
Public Const k_sqlWritingFileNum = 1662                     ' File numero # 1% #
Public Const k_sqlWritingClosingFileNum = 1663              ' Chiusura File numero # 1% #

Public Const k_sqlValidSqlSyntax = 1665                     ' Valida Sintassi T-Sql (Testi tra ') incluse Funzioni T-Sql
Public Const k_sqlTB_opt_SlitFile = 1666                    ' Nuovo File ogni N righe
Public Const k_sqlTB_opt_SlitFileHelp = 1667                ' Crea Nuovo File di Export ogni N righe

Public Const k_sqlTB_opt_ScriptTable = 1668                 ' Crea SQL Tabella
Public Const k_sqlTB_opt_ScriptTableHelp = 1669             ' Genera Sql DDL di generazione Tabella
Public Const k_sqlTB_opt_ScriptKeys = 1670                  ' Chiavi Interne
Public Const k_sqlTB_opt_ScriptKeysHelp = 1671              ' Genera Chiavi Interne (no Foreign Keys)

Public Const k_sqlTB_opt_ScriptGenerate = 1675              ' Script DDL di generazione Tabella
Public Const k_sqlTB_opt_ScriptGenerateWarning0 = 1676      ' Lo Script genera esclusivamente la struttura della tabella       'Max 74 char
Public Const k_sqlTB_opt_ScriptGenerateWarning1 = 1677      ' e degli oggetti selezionati;
Public Const k_sqlTB_opt_ScriptGenerateWarning2 = 1678      ' Gli Oggetti dai quali Dipende sono esclusi.


Public Const k_PROP_Property_Delete = 1847                  ' Elimina Properietà Estesa
Public Const k_PROP_Property_New = 1848                     ' Nuova Properietà Estesa
Public Const k_PROP_Property_Value = 1849                   ' Valore Properietà Estesa
Public Const k_PROP_Selected_Property = 1850                ' Properietà Estesa Selezionata...
Public Const k_PROP_Form_Extended_Property = 1851           ' Proprietà Estese oggetto # 1% # ( 2% )
Public Const k_PROP_Form_Available_Property = 1852          ' Proprietà Estese Disponibili
Public Const k_PROP_Form_Property_Name = 1853               ' Nome Proprietà Estesa
Public Const k_PROP_Form_Property_Value = 1854              ' Valore Testuale Proprietà Estesa (Visualizzata ed eventualmente Salvata come Testo)


Public Const k_TRI_Type_Update = 1862                       ' Variazione
Public Const k_TRI_Type_Insert = 1863                       ' Inserimento
Public Const k_TRI_Type_Delete = 1864                       ' Cancellazione
Public Const k_TRI_Type_All = 1865                          ' Variaz./Inser./Canc.
Public Const k_Sp_Type_Extended = 1869                      ' Estesa
Public Const k_Sp_Type_Standard = 1870                      ' Standard

Public Const k_Param_QuotedIdentifier = 1871                ' Stato Quoted Identifier
Public Const k_Param_ANSI_nulls = 1872                      ' Stato ANSI Nulls
Public Const k_Func_Param_IsSchemaBound = 1873              ' Associata a Schema
Public Const k_Func_Param_IsDeterministic = 1874            ' Deterministica
Public Const k_Param_Encrypted = 1875                       ' Criptata

Public Const k_Param_ColId = 1876                           ' Ordine
Public Const k_Param_Name = 1877                            ' Parametro
Public Const k_Param_Direction = 1879                       ' Direzione

Public Const k_Func_NewFunction = 1890                      ' Nuova Funzione Utente
Public Const k_Func_Scalar = 1891                           ' Scalare
Public Const k_Func_InLine = 1892                           ' In-Line
Public Const k_Func_Table = 1893                            ' Tabella


Public Const k_DBCC_Comleted = 1900                         ' Esecuzione DBCC completata

Public Const k_TriggerInsteadOf = 1912                      ' INSTEAD OF
Public Const k_TriggerName = 1913                           ' Nome Trigger
Public Const k_ManageTriggers = 1914                        ' Gestione Trigger
Public Const k_objFunction = 1915                           ' Funzione Utente

'Bck
Public Const k_Bck_Complete = 1918                          ' Database - Completo
Public Const k_Bck_Differential = 1919                      ' Database - Differenziale
Public Const k_Bck_Log = 1920                               ' Log delle Transazioni
'newDB
Public Const k_NewDbSort = 1923                             ' Regole di Confronto
Public Const k_NewDbDataFile = 1924                         ' File di Dati
Public Const k_NewDbTrLog = 1925                            ' Log delle Transazioni

Public Const k_MnuSelectAll = 1930                          ' Seleziona Tutto
Public Const k_MnuCut = 1931                                ' Taglia
Public Const k_MnuCopy = 1932                               ' Copia
Public Const k_MnuPaste = 1933                              ' Incolla
Public Const k_WordWrap = 1934                              ' Disabilita Word Wrap
Public Const k_WordWrapHelp = 1935                          ' Disabilita/Abilita Word Wrap
Public Const k_Font_GoStm = 1936                            ' Terminatore Batch
Public Const k_Font_RES = 1937                              ' Font Risultato e Messaggi
Public Const k_Font_Grd = 1938                              ' Font Griglia Risultato
Public Const k_Font_Qry = 1939                              ' Font Testo Query

Public Const k_Log_Name = 1940                              ' Log Nome

Public Const k_Login = 1954                                 ' Login
Public Const k_Device = 1955                                ' Device
Public Const k_Role = 1956                                  ' Ruolo
Public Const k_Role_Type = 1957                             ' Tipo Ruolo
Public Const k_Physical_Location = 1958                     ' Posizione Fisica
Public Const k_Lenght = 1959                                ' Lunghezza
Public Const k_BaseType = 1960                              ' Tibo Dati Base

Public Const k_Shrink_SingleUser = 1968                     ' (Single User)
Public Const k_Shrink_Default5% = 1969                      ' Shrink_Default (prova a 5%)

Public Const kTviewDatabases = 1970                         ' Database
Public Const kTviewDatabasesUsers = 1971                    ' Utenti Database
Public Const kTviewDatabasesTables = 1972                   ' Tabelle Database
Public Const kTviewDatabasesViews = 1973                    ' Viste Database
Public Const kTviewDatabasesStoredProc = 1974               ' Stored Procedures Database
Public Const kTviewDatabasesRoles = 1975                    ' Ruoli Database
Public Const kTviewDatabasesUDT = 1976                      ' Tipi Definiti dall'Utente
Public Const kTviewLogin = 1977                             ' Login
Public Const kTviewDevices = 1978                           ' Device di BackUp
Public Const kTviewLogs = 1979                              ' Log Sql Server
Public Const kTviewLockID = 1980                            ' Blocchi /ID
Public Const kTviewFunction = 1981                          ' Funzioni Utente

Public Const kErrorBatch = 1984                             ' Errore 1% - 2% durante l'esecuzione del batch: 3%
Public Const kErr_Insuf_Data = 1985                         ' Dati Insufficienti
Public Const kMsgBoxError = 1986                            ' Errore
Public Const kNotAvailable = 1987                           ' Non Disponibile
Public Const kErr_AvalilableOnlyOnHostServer = 1988         ' Informazione disponibile solo con accesso dal SqlServer Host # 1% #
Public Const kErr_No_Name = 1989                            ' Nessun Nome Specificato

Public Const k_Running_on = 1990                            ' eseguito su ..
Public Const k_Credits = 1991                               ' Ringraziamenti
Public Const k_Thanks = 1992                                ' Grazie a ...

Public Const k_POP_UP_UNAVAILABLE = 1998                   ' PopUp Menu non disponibile
Public Const k_SYS_INFO_UNAVAILABE = 1999                   ' Informazioni di Sistema non disponibili
'Public Const k_EMail As String = "a_montanari@libero.it"
Public Const k_EMail As String = "montanari_andrea@virgilio.it"
Public Const k_URLwww As String = "http://utenti.lycos.it/asql/index.html"


