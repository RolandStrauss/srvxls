**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@10 : Microsoft Excel - Create
//   Type : *MODULE
// Description :
//   Service procedures to create Microsoft Excel Workbooks.
//   Open Source service program HSSFR4 used - Scott Klement.
// --------------------------------------------------------------------------------------------------
//   Pre-Compiler tags used by STRPREPRC to retrieve creation
//   commands from the source member.
// -------------------------------------------------
// >>PRE-COMPILER<<
//   >>CRTCMD<< CRTRPGMOD    MODULE(&LI/&OB) +
//                           SRCFILE(&SL/&SF) +
//                           SRCMBR(&SM);
//   >>COMPILE<<
//     >>PARM<< TRUNCNBR(*NO);
//     >>PARM<< DBGVIEW(*ALL);
//     >>PARM<< OPTION(*EVENTF);
//   >>END-COMPILE<<
//   >>EXECUTE<<
// >>END-PRE-COMPILER<<
// --------------------------------------------------------------------------------------------------
// Date        Developer         Change
// 2019-01-03  Nico Basson       Initial Code.
// --------------------------------------------------------------------------------------------------
/INCLUDE SRVSRC,SRV@H
ctl-opt nomain
        thread(*concurrent)
        stgmdl(*inherit) ;
// Framework ---------------------------------------------------------------------------------------
/DEFINE  SRVXLS
/INCLUDE SRVSRC,SRV@P
/INCLUDE SRVSRC,SRV@T
/INCLUDE SRVSRC,SRVXLS@T
// --------------------------------------------------------------------------------------------------
// Globals <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-s JNI_Env@     pointer                                                                  import ;
// Workbook ----------------------------------------------------------------------------------------
dcl-s XSSF_Workbook object( *java
                          : XSSF_Workbook_class
                          )                                                                 export ;
// --------------------------------------------------
dcl-s Directory  varchar(512)                                                               export ;
dcl-s ExcelDoc   varchar(512)                                                               export ;
// Sheet -------------------------------------------------------------------------------------------
dcl-s Sheet object ( *java
                   : Sheet_class
                   )                                                                        export ;
// --------------------------------------------------
dcl-s SheetName  varchar(128)                                                               export ;
// Row ---------------------------------------------------------------------------------------------
dcl-s Row object( *java
                : Row_class
                )                                                                           export ;
// --------------------------------------------------
dcl-s RowNumber  int(10) inz(-1)                                                            export ;
// Column ------------------------------------------------------------------------------------------
dcl-s ColNumber  int(5)  inz(-1)                                                            export ;
// Auto Column Width -------------------------------
dcl-s AutoWidth  ind     inz                                                                export ;
// --------------------------------------------------------------------------------------------------
// Open a .xlsx (XSSF) Workbook <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_OpenWorkbook export ;
   dcl-pi   xls_OpenWorkbook ;
      i_ExcelDoc   varchar(512) const ;
      i_Directory  varchar(512) const ;
      i_Create     ind          const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s ExcelPath varchar(1024) ;
   dcl-s Create    ind ;
   // --------------------------------------------------------------------------------------------------
   dcl-ds ErrorData qualified ;
      ExcelDoc     char(512) ;
      Directory    char(512) ;
   end-ds ;
   // Create Default Style Objects oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
   dcl-pr f_CrtDftStyles ;
   end-pr ;
   // Workbook Factory Create Method jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_WorkbookFactory_create  object (*java : Workbook_class )
                                 extproc( *java
                                        : 'org.apache.poi.ss.usermodel.WorkbookFactory'
                                        : 'create'
                                        )
                                 static ;
      i_FileInputStream  like(jInputStream) const  ;
   end-pr ;
   // --------------------------------------------------
   dcl-s FileInputStream  like(jFileInputStream) ;
   dcl-s HSSF_Workbook    object( *java
                             : HSSF_Workbook_class
                             )
                       inz(*null) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      s_CleanupJAVA() ;

      // Excel Document -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      ExcelDoc = i_ExcelDoc ;
      Directory = gen_DirPathName( i_Directory ) ;
      // -----------------------------------------------
      ExcelPath = Directory + ExcelDoc ;

      // JVM Pointer =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (JNI_Env@ = *null) ;
         JNI_Env@ = java_JNI_Env@() ;
         java_BeginObjGroup( JNI_Env@
                        : 1000
                        ) ;
      endif ;

      // Create -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Create = *off ;
      select ;
         when (%parms < %parmnum( i_Create )) ;
         when (%addr( i_Create ) <> *null) ;
            // ----------------------------------------
            Create = i_Create ;
      endsl;
      // File not Found -------------------------------
      select ;
         when (ifs_FileExists( %trim( ExcelPath ) )) ;
         when (NOT Create) ;
            // ----------------------------------------
            ErrorData.ExcelDoc  = ExcelDoc ;
            ErrorData.Directory = Directory ;
            err_SendEsc( 'XLS0002'
                     : 'SRVERR'
                     : ErrorData
                     ) ;
      endsl;
      // -----------------------------------------------
      select ;
         when (NOT Create) ;
         when (ifs_FileExists( %trim( ExcelPath ) )) ;
            // ----------------------------------------
            ifs_DeleteFile( %trim( ExcelPath ) ) ;
            s_CreateWorkbook() ;
            exsr sr_Exit ;
         other ;
            // ----------------------------------------
            s_CreateWorkbook() ;
            exsr sr_Exit ;
      endsl;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      java_BeginObjGroup( JNI_Env@
                     : 1000
                     ) ;
      // -----------------------------------------------
      FileInputStream = java_NewFileInputStream( ExcelPath ) ;
      HSSF_Workbook   = j_WorkbookFactory_create( FileInputStream ) ;
      // Cleanup --------------------------------------
      java_CloseFileInputStream( FileInputStream ) ;
      java_EndObjGroup( JNI_Env@
                   : HSSF_Workbook
                   : XSSF_Workbook
                   ) ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      exsr sr_Exit ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // On-Exit >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   begsr sr_Exit ;
      // ---------------------------------------------------------
      monitor ;
         // Default Style Objects =-=-=-=-=-=-=-=-=-=-=-
         f_CrtDftStyles() ;
         return ;
         // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      on-error ;
         f_CleanupJAVA() ;
         err_BubbleMsg() ;
      endmon ;
      // ---------------------------------------------------------
   endsr;
   // --------------------------------------------------------------------------------------------------
end-proc xls_OpenWorkbook ;
// --------------------------------------------------------------------------------------------------
// Create a New .xlsx (XSSF) Workbook --------------------------------------------------------------
// --------------------------------------------------------------------------------------------------
dcl-proc s_CreateWorkbook ;
   // Create a new XSSF (OOXML Excel) workbook jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_New_XSSF_Workbook like(XSSF_Workbook) extproc( *java
                                                      : XSSF_Workbook_class
                                                      : *constructor
                                                      ) ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      ifs_CreateDirPath( %trim( Directory ) ) ;

      // Instantiate Java Workbook Class =-=-=-=-=-=-=-
      XSSF_Workbook = j_New_XSSF_Workbook() ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc s_CreateWorkbook ;
// --------------------------------------------------------------------------------------------------
// Save .xlsx (XSSF) Workbook <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SaveWorkbook export ;
   dcl-pi   xls_SaveWorkbook ;
      i_Replace      ind          const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s ExcelPath varchar(1024) ;
   dcl-s Replace   ind ;
   // --------------------------------------------------------------------------------------------------
   dcl-ds ErrorData qualified ;
      ExcelDoc     char(512) ;
      Directory    char(512) ;
   end-ds ;
   // --------------------------------------------------------------------------------------------------
   dcl-s FileStream  like(jfileoutputstream) ;
   // Write Workbook to Output Stream jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_XSSF_Workbook_Write extproc( *java
                                    : Workbook_class
                                    : 'write'
                                    ) ;
      i_OutputStream like( jOutputStream ) ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      ExcelPath = Directory + ExcelDoc ;

      // Replace =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Replace = *off ;
      select ;
         when (%parms < %parmnum( i_Replace )) ;
         when (%addr( i_Replace ) <> *null) ;
            // ----------------------------------------
            Replace = i_Replace ;
      endsl;
      // Check if Exists ------------------------------
      select ;
         when (NOT ifs_FileExists( %trim( ExcelPath ) )) ;
         when (Replace) ;
            // Delete Existing -----------------------
            ifs_DeleteFile( %trim( ExcelPath ) ) ;
         other ;
            // ----------------------------------------
            ErrorData.ExcelDoc  = ExcelDoc ;
            ErrorData.Directory = Directory ;
            err_SendEsc( 'XLS0003'
                     : 'SRVERR'
                     : ErrorData
                     ) ;
      endsl;

      // Output to Streamfile (IFS) -=-=-=-=-=-=-=-=-=-
      FileStream = java_NewFileOutputStream( ExcelPath ) ;
      // -----------------------------------------------
      j_XSSF_Workbook_Write( XSSF_Workbook
                        : FileStream
                        ) ;
      // -----------------------------------------------
      java_CloseFileOutputStream( FileStream ) ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( FileStream ) ;
      s_CleanupJAVA() ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SaveWorkbook ;
// --------------------------------------------------------------------------------------------------
// Cleanup Java ------------------------------------------------------------------------------------
// --------------------------------------------------------------------------------------------------
dcl-proc s_CleanupJAVA ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Clear Styles -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      xls_ClearStyle() ;

      // Java Objects -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      // Workbook -------------------------------------
      if (XSSF_Workbook <> *null) ;
         f_FreeLocalRef( XSSF_Workbook ) ;
         XSSF_Workbook = *null ;
      endif ;
      clear Directory     ;
      clear ExcelDoc      ;
      // Sheet ----------------------------------------
      if (Sheet <> *null) ;
         f_FreeLocalRef( Sheet ) ;
         Sheet = *null ;
      endif ;
      clear SheetName     ;
      // Row ------------------------------------------
      if (Row <> *null) ;
         f_FreeLocalRef( Row ) ;
         Row = *null ;
      endif ;
      clear RowNumber     ;
      // -----------------------------------------------
      clear ColNumber     ;

      // Object Group and JNI -=-=-=-=-=-=-=-=-=-=-=-=-
      f_CleanupJAVA() ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc s_CleanupJAVA ;
// --------------------------------------------------------------------------------------------------
// Open Sheet in .xlsx (XSSF) Workbook <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_OpenSheet export ;
   dcl-pi   xls_OpenSheet ;
      i_SheetName  varchar(128) const ;
      i_Create     ind          const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s Create  ind ;
   // --------------------------------------------------------------------------------------------------
   dcl-s NameString  like(jString) ;
   // Retrieve Sheet from a Workbook jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_GetSheet like(Sheet) extproc( *java
                                     : Workbook_class
                                     : 'getSheet'
                                     ) ;
      i_NameString like( jString ) ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   dcl-ds ErrorData qualified ;
      ExcelDoc     char(512) ;
      Directory    char(512) ;
      SheetName    char(128) ;
   end-ds ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Create Sheet -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Create = *off ;
      select ;
         when (%parms < %parmnum( i_Create )) ;
         when (%addr( i_Create ) <> *null) ;
            // ----------------------------------------
            Create = i_Create ;
      endsl;

      // Get Sheet =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      SheetName  = i_SheetName ;
      NameString = java_NewString( SheetName ) ;
      // -----------------------------------------------
      Sheet = j_GetSheet( XSSF_Workbook
                     : NameString
                     ) ;
      // -----------------------------------------------
      select ;
            // No Problems --------------------------------
         when (Sheet <> *null) ;
            // Create Sheet -------------------------------
         when (Create) ;
            xls_AddSheet( i_SheetName ) ;
            // Sheet not Found Error ----------------------
         other ;
            // ----------------------------------------
            ErrorData.ExcelDoc  = ExcelDoc ;
            ErrorData.Directory = Directory ;
            ErrorData.SheetName = i_SheetName ;
            err_SendEsc( 'XLS0004'
                     : 'SRVERR'
                     : ErrorData
                     ) ;
      endsl;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( NameString ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_OpenSheet ;
// --------------------------------------------------------------------------------------------------
// Add Sheet to .xlsx (XSSF) Workbook <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_AddSheet export ;
   dcl-pi   xls_AddSheet ;
      i_SheetName    varchar(128) const ;
      i_Replace      ind          const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s Replace  ind ;
   // --------------------------------------------------------------------------------------------------
   dcl-s NameString  like(jString) ;
   // Create a new sheet in workbook jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_CreateSheet like(Sheet) extproc( *java
                                        : Workbook_class
                                        : 'createSheet'
                                        ) ;
      i_NameString like( jString ) ;
   end-pr ;
   // Error Message Information -----------------------------------------------------------------------
   dcl-ds MsgInfo likeds(MsgInfo_t) ;
   // --------------------------------------------------
   dcl-s  MsgData@  pointer inz(%addr(MsgInfo.RplDta)) ;
   dcl-ds JavaError likeds(JavaErrorRNX0301_t) based(MsgData@) ;
   // --------------------------------------------------
   dcl-ds ErrorData qualified ;
      ExcelDoc     char(512) ;
      Directory    char(512) ;
      SheetName    char(128) ;
   end-ds ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Add Sheet =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      SheetName  = i_SheetName ;
      NameString = java_NewString( SheetName ) ;
      // -----------------------------------------------
      monitor ;
         // ---------------------------------------------
         Sheet = j_CreateSheet( XSSF_Workbook
                          : NameString
                          ) ;
         // ---------------------------------------------
      on-error ;
         // Cleanup ------------------------------------
         f_FreeLocalRef( NameString ) ;
         // Retrieve Error Information -----------------
         MsgInfo = err_RtvLastMsgInfo() ;
         // ---------------------------------------------
         select ;
               // Other Error ------------------------------
            when (%scan( 'The workbook already contains a sheet of this name'
                 : JavaError.Exception
                 ) = 0) ;
               // --------------------------------------
               f_CleanupJAVA() ;
               err_BubbleMsg() ;
               // Replace Sheet ----------------------------
            when (Replace) ;
               xls_RmvSheet( i_SheetName ) ;
               xls_AddSheet( i_SheetName ) ;
               // Sheet already exists ---------------------
            other ;
               // --------------------------------------
               ErrorData.ExcelDoc  = ExcelDoc ;
               ErrorData.Directory = Directory ;
               ErrorData.SheetName = i_SheetName ;
               err_SendEsc( 'XLS0005'
                       : 'SRVERR'
                       : ErrorData
                       ) ;
         endsl;
         // ---------------------------------------------
      endmon ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( NameString ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_AddSheet ;
// --------------------------------------------------------------------------------------------------
// Remove Sheet from .xlsx (XSSF) Workbook <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_RmvSheet export ;
   dcl-pi   xls_RmvSheet ;
      i_SheetName  varchar(128) const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s NameString  like(jString) ;
   // Get sheet index jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_GetSheetIndex like( jInt ) extproc( *java
                                           : Workbook_class
                                           : 'getSheetIndex'
                                           ) ;
      i_NameString like( jString ) ;
   end-pr ;
   // Remove sheet from workbook jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_RemoveSheet extproc( *java
                            : Workbook_class
                            : 'removeSheetAt'
                            ) ;
      i_SheetIndex like( jInt ) value ;
   end-pr ;
   // --------------------------------------------------
   dcl-s SheetIndex int(5) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Sheet Name -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_SheetName )) ;
         when (%addr( i_SheetName ) = *null) ;
         when (SheetName <> i_SheetName) ;
            // ----------------------------------------
            SheetName = i_SheetName ;
      endsl;

      // Get Sheet Index =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      NameString = java_NewString( SheetName ) ;
      SheetIndex = j_GetSheetIndex( XSSF_Workbook
                               : NameString
                               ) ;

      // Remove Sheet -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_RemoveSheet( XSSF_Workbook
                : SheetIndex
                ) ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( NameString ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_RmvSheet ;
// --------------------------------------------------------------------------------------------------
// Add Row to Sheet <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_AddRow export ;
   dcl-pi   xls_AddRow int(10) ;
      i_RowNumber    int(10)      const options( *omit : *nopass ) ;
      i_SheetName    varchar(128) const options( *omit : *nopass ) ;
   end-pi ;
   // Create a row in the sheet jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_CreateRow like(Row) extproc( *java
                                    : Sheet_class
                                    : 'createRow'
                                    ) ;
      i_RowNumber like( jint ) value ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Row Number -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_RowNumber )) ;
            RowNumber += 1 ;
         when (%addr( i_RowNumber ) = *null) ;
            RowNumber += 1 ;
         other ;
            // ----------------------------------------
            RowNumber = i_RowNumber ;
      endsl;

      // Reset Column Number =-=-=-=-=-=-=-=-=-=-=-=-=-
      reset ColNumber ;

      // Sheet Name -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_SheetName )) ;
         when (%addr( i_SheetName ) = *null) ;
         when (i_SheetName <> SheetName) ;
            // ----------------------------------------
            SheetName = i_SheetName ;
      endsl;

      // Create Row -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Row = j_CreateRow( Sheet
                    : RowNumber
                    ) ;
      // -----------------------------------------------
      return RowNumber ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_AddRow ;
// --------------------------------------------------------------------------------------------------
// Get Row from Sheet <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_GetRow export ;
   dcl-pi   xls_GetRow ;
      i_RowNumber  int(10)      const ;
      i_Create     ind          const options( *omit : *nopass ) ;
      i_SheetName  varchar(128) const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s Create  ind ;
   // Retrieve row object from a sheet jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_GetRow like(Row) extproc( *java
                                 : Sheet_class
                                 : 'getRow'
                                 ) ;
      i_RowNumber like( jint ) value ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      RowNumber = i_RowNumber ;

      // Create -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Create = *off ;
      select ;
         when (%parms < %parmnum( i_Create )) ;
         when (%addr( i_Create ) <> *null) ;
            // ----------------------------------------
            Create = i_Create ;
      endsl;

      // Sheet Name -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_SheetName )) ;
         when (%addr( i_SheetName ) = *null) ;
         when (i_SheetName <> SheetName) ;
            // ----------------------------------------
            SheetName = i_SheetName ;
      endsl;

      // Get Row =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Row = j_GetRow( Sheet
                 : RowNumber
                 ) ;
      // -----------------------------------------------
      select ;
            // No Problems --------------------------------
         when (Row <> *null) ;
            // Create Row ---------------------------------
         when (Create) ;
            xls_AddRow( RowNumber ) ;
      endsl;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_GetRow ;
// --------------------------------------------------------------------------------------------------
