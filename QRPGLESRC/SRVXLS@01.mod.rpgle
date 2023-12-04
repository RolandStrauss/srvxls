**free
// ///////////////////////////////////////////////////////////////////////////////////////////// //
//  ____            _                       _     ____    _                                      //
// |  _ \    ___   | |   __ _   _ __     __| |   / ___|  | |_   _ __   __ _   _   _   ___   ___  //
// | |_) |  / _ \  | |  / _` | | '_ \   / _` |   \___ \  | __| | '__| / _` | | | | | / __| / __| //
// |  _ <  | (_) | | | | (_| | | | | | | (_| |    ___) | | |_  | |   | (_| | | |_| | \__ \ \__ \ //
// |_| \_\  \___/  |_|  \__,_| |_| |_|  \__,_|   |____/   \__| |_|    \__,_|  \__,_| |___/ |___/ //
//                                                                                               //
// ///////////////////////////////////////////////////////////////////////////////////////////// //
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@01 : Microsoft Excel Parsing
//   Type : *MODULE
// Description :
//   Service procedures to read Microsoft Excel Workbook data.
//   Open Source service program XLPARSE4 used - Scott Klement and Spaghettiman - Giovanni Perotti.
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
// 2018-06-13  Nico Basson       Initial Code.
// --------------------------------------------------------------------------------------------------
// /INCLUDE SRVSRC,SRV@H
ctl-opt nomain
        thread(*concurrent)
        stgmdl(*inherit) ;
// Framework ---------------------------------------------------------------------------------------
// /DEFINE  SRVXLS
// /INCLUDE SRVSRC,SRV@P
// --------------------------------------------------------------------------------------------------
// Internal Function Prototypes oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
// Enable row start and end callbacks oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
dcl-pr f_ParseNotify ;
   i_NewRowProc@  pointer(*proc) const ;
   i_EndRowProc@  pointer(*proc) const ;
end-pr ;
// Java Cleanup oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
dcl-pr f_CleanupJAVA ;
end-pr ;
// Escape High Level Programs oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
dcl-pr f_Escape ind ;
   StackCnt int(10)   value ;
   MsgTxt   char(256) const ;
end-pr ;
// Sheet Name handler oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
dcl-pr f_SheetName varchar(1024) ;
   i_Name    varchar(1024) const options( *nopass ) ;
end-pr ;
// Sheet Found handler ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
dcl-pr f_SheetFound ind ;
   i_Found    ind const options( *nopass ) ;
end-pr ;
// Static storage procedure pointer handler oooooooooooooooooooooooooooooooooooooooooooooooooooooooo
dcl-pr f_Proc@ pointer(*proc) ;
   i_Type  varchar(15)    const ;
   i_@     pointer(*proc) const options( *nopass ) ;
end-pr ;
// --------------------------------------------------------------------------------------------------
// Cell Callback Procedure Common Prototypes MWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
// --------------------------------------------------------------------------------------------------
// Cell Procedure bpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbp
dcl-pr CellProc ind extproc(CellProc@) ;
   i_Sheet      varchar(1024)  const ;
   i_Row        int(10)        value ;
   i_Column     int(5)         value ;
   i_Value@     pointer        value ;
   i_ValueType  char(1)        const ;
   i_Nan        int(5)         const options( *omit : *nopass ) ;
   i_Formula    varchar(32767) const options( *omit : *nopass ) ;
end-pr ;
// --------------------------------------------------
dcl-s CellProc@ pointer(*proc) ;
// --------------------------------------------------
dcl-s  Numeric@ pointer inz(%addr(Numeric)) ;
dcl-ds Numeric  qualified ;
   Value   float(8) ;
end-ds ;
dcl-s  String@ pointer inz(%addr(String)) ;
dcl-ds String  qualified ;
   Value  varchar(32767) ;
end-ds;
// --------------------------------------------------------------------------------------------------
// Globals <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-s JNI_Env@ pointer export ;
// --------------------------------------------------------------------------------------------------
// Parse the contents of an .xls/.xlsx workbook <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_ParseWorkbook export ;
   dcl-pi   xls_ParseWorkbook ;
      i_Directory    varchar(512)   const ;
      i_ExcelDoc     varchar(512)   const ;
      i_CellProc@    pointer(*proc) const ;
      i_NewRowProc@  pointer(*proc) const options( *omit : *nopass ) ;
      i_EndRowProc@  pointer(*proc) const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s ExcelPath varchar(1024) ;
   // Parse the contents of an .xls/.xlsx workbook <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
   dcl-pr xlparse_workbook int(10) extproc('XLPARSER@1_xlparse_workbook') ;
      i_ExcelPath     varchar(1024)  const ;
      i_NumericProc@  pointer(*proc) value ;
      i_StringProc@   pointer(*proc) value ;
      i_FormulaProc@  pointer(*proc) value options( *nopass ) ;
      i_RetainLog     ind            const options( *nopass : *omit ) ;
   end-pr ;
   // --------------------------------------------------
   dcl-s RC int(10) ;
   // Error Handling ----------------------------------------------------------------------------------
   dcl-ds Error likeds(MsgInfo_t) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // JVM Pointer =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (JNI_Env@ = *null) ;
         JNI_Env@ = java_JNI_Env@() ;
         java_BeginObjGroup( JNI_Env@ ) ;
      endif ;

      // Excel Document Path =-=-=-=-=-=-=-=-=-=-=-=-=-
      ExcelPath = i_Directory ;
      if (%subst( ExcelPath : %len( ExcelPath ) : 1 ) <> '/') ;
         ExcelPath += '/' ;
      endif ;
      ExcelPath += i_ExcelDoc ;

      // Cell callback procedure =-=-=-=-=-=-=-=-=-=-=-
      f_Proc@( 'Cell' : i_CellProc@ ) ;

      // Row Start and End callback procedures =-=-=-=-
      select ;
         when (%parms < %parmnum( i_NewRowProc@ )) ;
         when (%parms < %parmnum( i_EndRowProc@ )) ;
         when (%addr( i_NewRowProc@ ) = *null) ;
         when (%addr( i_EndRowProc@ ) = *null) ;
         when (i_NewRowProc@ <> *null
          AND i_EndRowProc@ <> *null) ;
            // ----------------------------------------
            f_ParseNotify( i_NewRowProc@
                       : i_EndRowProc@
                       ) ;
         other;
      endsl;

      // Parse Workbook -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      RC = xlparse_workbook( ExcelPath
                        : %paddr( cb_NumericCell )
                        : %paddr( cb_StringCell  )
                        : %paddr( cb_FormulaCell )
                        : *off
                        ) ;
      // Error ----------------------------------------
      if (RC < 0) ;
         // Send as escape message --------------------
         Error = err_RtvLastMsgInfo() ;
         err_SendEsc( Error.MessageID
                 : Error.MSGF_Name + Error.MSGF_LibUsed
                 : Error.RplDta
                 ) ;
         // --------------------------------------------
      endif ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_CleanupJAVA() ;
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_ParseWorkbook ;
// --------------------------------------------------------------------------------------------------
// Parse the contents of an .xls/.xlsx workbook sheet <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_ParseSheet export ;
   dcl-pi   xls_ParseSheet ;
      i_Directory    varchar(512)   const ;
      i_ExcelDoc     varchar(512)   const ;
      i_Sheet        varchar(512)   const ;
      i_CellProc@    pointer(*proc) const ;
      i_NewRowProc@  pointer(*proc) const options( *omit : *nopass ) ;
      i_EndRowProc@  pointer(*proc) const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-ds Proc@ qualified static ;
      NewRow   pointer(*proc) ;
      EndRow   pointer(*proc) ;
   end-ds ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Row Start and End callback procedures =-=-=-=-
      clear Proc@ ;
      select ;
         when (%parms < %parmnum( i_NewRowProc@ )) ;
         when (%parms < %parmnum( i_EndRowProc@ )) ;
         when (%addr( i_NewRowProc@ ) = *null) ;
         when (%addr( i_EndRowProc@ ) = *null) ;
         when (i_NewRowProc@ <> *null
          AND i_EndRowProc@ <> *null) ;
            // ----------------------------------------
            Proc@.NewRow = i_NewRowProc@ ;
            Proc@.EndRow = i_EndRowProc@ ;
         other;
      endsl;

      // Set Sheet Name in Static Storage -=-=-=-=-=-=-
      f_SheetName( i_Sheet ) ;

      // Parse Workbook -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      xls_ParseWorkbook( i_Directory
                    : i_ExcelDoc
                    : i_CellProc@
                    : Proc@.NewRow
                    : Proc@.EndRow
                    ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_ParseSheet ;
// --------------------------------------------------------------------------------------------------
// Cell Callback Procedure Wrappers WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
// --------------------------------------------------------------------------------------------------
// Callback for Numeric Cell wmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmw
// --------------------------------------------------------------------------------------------------
dcl-proc cb_NumericCell ;
   dcl-pi   cb_NumericCell ;
      i_Sheet varchar(1024) const ;
      i_Row   int(10)       value ;
      i_Col   int(5)        value ;
      i_Value float(8)      value ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s SheetName varchar(1024) static ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Limit to specific sheet =-=-=-=-=-=-=-=-=-=-=-
      SheetName = f_SheetName() ;
      // -----------------------------------------------
      select ;
         when (SheetName = *blank) ;
         when (i_Sheet <> SheetName) ;
            return ;
         when (i_Sheet = SheetName) ;
            f_SheetFound( *on ) ;
         when (f_SheetFound()) ;
            f_Escape( 1 : 'End Of Data' ) ;
         other;
      endsl;

      // Callback procedure -=-=-=-=-=-=-=-=-=-=-=-=-=-
      CellProc@ = f_Proc@( 'Cell' ) ;
      if (CellProc@ = *null) ;
         return ;
      endif;
      // -----------------------------------------------
      Numeric.Value = i_Value ;
      String.Value  = %trimr( %char( i_Value ) ) ;
      select ;
            // Numeric ------------------------------------
         when (CellProc( i_Sheet
                  : i_Row
                  : i_Col
                  : Numeric@
                  : 'N'
                  ) = *on) ;
            // ----------------------------------------
            f_Escape( 1 : 'End Of Data' ) ;
            // String -------------------------------------
         when (CellProc( i_Sheet
                  : i_Row
                  : i_Col
                  : String@
                  : 'S'
                  ) = *on) ;
            // ----------------------------------------
            f_Escape( 1 : 'End Of Data' ) ;
         other;
      endsl;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc cb_NumericCell ;
// --------------------------------------------------------------------------------------------------
// Callback for String Cell mwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmw
// --------------------------------------------------------------------------------------------------
dcl-proc cb_StringCell export ;
   dcl-pi   cb_StringCell static ;
      i_Sheet varchar(1024)  const ;
      i_Row   int(10)        value ;
      i_Col   int(5)         value ;
      i_Value varchar(32767) const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s SheetName varchar(1024) static ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Limit to specific sheet =-=-=-=-=-=-=-=-=-=-=-
      SheetName = f_SheetName() ;
      // -----------------------------------------------
      select ;
         when (SheetName = *blank) ;
         when (i_Sheet <> SheetName) ;
            return ;
         when (i_Sheet = SheetName) ;
            f_SheetFound( *on ) ;
         when (f_SheetFound()) ;
            f_Escape( 1 : 'End Of Data' ) ;
         other;
      endsl;

      // Callback procedure -=-=-=-=-=-=-=-=-=-=-=-=-=-
      CellProc@ = f_Proc@( 'Cell' ) ;
      if (CellProc@ = *null) ;
         return ;
      endif;
      // -----------------------------------------------
      String.Value = i_Value ;
      select ;
            // String -------------------------------------
         when (CellProc( i_Sheet
                  : i_Row
                  : i_Col
                  : String@
                  : 'S'
                  ) = *on) ;
            // ----------------------------------------
            f_Escape( 1 : 'End Of Data' ) ;
         other;
      endsl;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc cb_StringCell ;
// --------------------------------------------------------------------------------------------------
// Callback for Formula Cell wmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmw
// --------------------------------------------------------------------------------------------------
dcl-proc cb_FormulaCell export ;
   dcl-pi   cb_FormulaCell static ;
      i_Sheet   varchar(1024)  const ;
      i_Row     int(10)        value ;
      i_Col     int(5)         value ;
      i_Value   float(8)       value ;
      i_NaN     int(5)         value ;
      i_Formula varchar(32767) const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s SheetName varchar(1024) static ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Limit to specific sheet =-=-=-=-=-=-=-=-=-=-=-
      SheetName = f_SheetName() ;
      // -----------------------------------------------
      select ;
         when (SheetName = *blank) ;
         when (i_Sheet <> SheetName) ;
            return ;
         when (i_Sheet = SheetName) ;
            f_SheetFound( *on ) ;
         when (f_SheetFound()) ;
            f_Escape( 1 : 'End Of Data' ) ;
         other;
      endsl;

      // Callback procedure -=-=-=-=-=-=-=-=-=-=-=-=-=-
      CellProc@ = f_Proc@( 'Cell' ) ;
      if (CellProc@ = *null) ;
         return ;
      endif;
      // -----------------------------------------------
      Numeric.Value = i_Value ;
      String.Value  = i_Formula ;
      select ;
            // Formula ------------------------------------
         when (CellProc( i_Sheet
                  : i_Row
                  : i_Col
                  : Numeric@
                  : 'F'
                  : i_NaN
                  : i_Formula
                  ) = *on) ;
            // ----------------------------------------
            f_Escape( 1 : 'End Of Data' ) ;
            // Numeric ------------------------------------
         when (CellProc( i_Sheet
                  : i_Row
                  : i_Col
                  : Numeric@
                  : 'N'
                  ) = *on) ;
            // ----------------------------------------
            f_Escape( 1 : 'End Of Data' ) ;
            // String -------------------------------------
         when (CellProc( i_Sheet
                  : i_Row
                  : i_Col
                  : String@
                  : 'S'
                  ) = *on) ;
            // ----------------------------------------
            f_Escape( 1 : 'End Of Data' ) ;
         other;
      endsl;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc cb_FormulaCell ;
// --------------------------------------------------------------------------------------------------
