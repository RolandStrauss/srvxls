**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@PU : Microsoft Excel Parsing - Parsing Utilities
//   Type : *MODULE
// Description :
//   These procedures holds the utility functions used by other modules in this service program.
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
// 2018-07-13  Nico Basson       Initial Code.
// --------------------------------------------------------------------------------------------------
/INCLUDE SRVSRC,SRV@H
ctl-opt nomain
        thread(*concurrent)
        stgmdl(*inherit) ;
// Framework ---------------------------------------------------------------------------------------
/DEFINE  SRVXLS
/INCLUDE SRVSRC,SRV@P
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
// Free Local Reference oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
dcl-pr f_FreeLocalRef ;
   i_localRef  object( *java : 'java.lang.Object' ) value ;
end-pr ;
// --------------------------------------------------------------------------------------------------
// Functions Used oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
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
// Globals <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-s JNI_Env@ pointer                                                                      import ;
// --------------------------------------------------------------------------------------------------
// Enable row start and end callbacks oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_ParseNotify export ;
   dcl-pi   f_ParseNotify  ;
      i_NewRowProc@  pointer(*proc) const ;
      i_EndRowProc@  pointer(*proc) const ;
   end-pi ;
   // Enable row start and end callbacks <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
   dcl-pr xlparse_notify extproc('XLPARSER@1_xlparse_notify') ;
      i_NewRowProc@  pointer(*proc) value ;
      i_EndRowProc@  pointer(*proc) value ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // JVM Pointer =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (JNI_Env@ = *null) ;
         JNI_Env@ = java_JNI_Env@() ;
         java_BeginObjGroup( JNI_Env@ ) ;
      endif ;

      // Set procedure pointers in handler =-=-=-=-=-=-
      f_Proc@( 'NewRow' : i_NewRowProc@ ) ;
      f_Proc@( 'EndRow' : i_EndRowProc@  ) ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      xlparse_notify( %paddr( cb_NewRow )
                 : %paddr( cb_EndRow )
                 ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_ParseNotify ;
// --------------------------------------------------------------------------------------------------
// Java Cleanup oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_CleanupJAVA export ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      if (JNI_Env@ <> *null) ;
         java_EndObjGroup( JNI_Env@ ) ;
      endif ;
      clear JNI_Env@ ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_CleanupJAVA ;
// --------------------------------------------------------------------------------------------------
// Escape High Level Programs oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_Escape export ;
   dcl-pi   f_Escape ind ;
      StackCnt int(10)   value ;
      MsgTxt   char(256) const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-pr QMHSNDPM extpgm('QMHSNDPM') ;
      *n char(7)    const ; // MessageID
      *n char(20)   const ; // QualMsgF
      *n char(256)  const ; // MsgData
      *n int(10)    const ; // MsgDtaLen
      *n char(10)   const ; // MsgType
      *n char(10)   const ; // CallStkEnt
      *n int(10)    const ; // CallStkCnt
      *n char(4)    ;       // MessageKey
      *n char(1024) options(*varsize) ; // ErrorCode
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   dcl-ds dsEC ;
      BytesProv  int(10) inz(0) ;
      BytesAvail int(10) inz(0) ;
   end-ds ;
   // --------------------------------------------------------------------------------------------------
   dcl-s MsgLen int(10) ;
   dcl-s MsgKey char(4) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      MsgLen = %len( %trimr( MsgTxt ) ) ;
      if ( MsgLen < 1 ) ;
         return *off ;
      endif ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      QMHSNDPM( 'CPF9897'
           : 'QCPFMSG *LIBL'
           : MsgTxt
           : MsgLen
           : '*ESCAPE'
           : '*'
           : StackCnt
           : MsgKey
           : dsEC
           ) ;
      // -----------------------------------------------
      return *off ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_Escape ;
// --------------------------------------------------------------------------------------------------
// Free Local Reference oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_FreeLocalRef export ;
   dcl-pi   f_FreeLocalRef ;
      i_localRef  object( *java : 'java.lang.Object' ) value ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Set JNI External Pointer -=-=-=-=-=-=-=-=-=-=-
      jniEnv_P = JNI_Env@ ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      deleteLocalRef( JNI_Env@
                 : i_localRef
                 ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_FreeLocalRef ;
// --------------------------------------------------------------------------------------------------
// Cell Callback Procedure Wrappers WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
// --------------------------------------------------------------------------------------------------
// Callback for New Row Processing wmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmw
// --------------------------------------------------------------------------------------------------
dcl-proc cb_NewRow ;
   dcl-pi   cb_NewRow ;
      i_Sheet varchar(1024) const ;
      i_Row   int(10)       value ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s SheetName varchar(1024) static ;
   // Sheet Name handler oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
   dcl-pr f_SheetName varchar(1024) ;
      i_Name    varchar(1024) const options( *nopass ) ;
   end-pr ;
   // New Row Procedure pbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbp
   dcl-pr NewRow ind extproc(NewRow@) ;
      *n varchar(1024) const ; // Sheet
      *n int(10)       value ; // Row
   end-pr ;
   // --------------------------------------------------
   dcl-s NewRow@ pointer(*proc) ;
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
      endsl;

      // Callback procedure call =-=-=-=-=-=-=-=-=-=-=-
      NewRow@ = f_Proc@( 'NewRow' ) ;
      select ;
         when (NewRow@ = *null) ;
            return ;
         when (NewRow( i_Sheet : i_Row ) = *on) ;
            f_Escape( 1 : 'End Of Data' ) ;
      endsl;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc cb_NewRow ;
// --------------------------------------------------------------------------------------------------
// Callback for End Row Processing wmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmwmw
// --------------------------------------------------------------------------------------------------
dcl-proc cb_EndRow ;
   dcl-pi   cb_EndRow ;
      i_Sheet varchar(1024) const ;
      i_Row   int(10)       value ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s SheetName varchar(1024) static ;
   // End Row Procedure pbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbpbp
   dcl-pr EndRow ind extproc(EndRow@) ;
      *n varchar(1024) const ; // Sheet
      *n int(10)       value ; // Row
   end-pr ;
   // --------------------------------------------------
   dcl-s EndRow@ pointer(*proc) ;
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
      endsl;

      // Callback procedure call =-=-=-=-=-=-=-=-=-=-=-
      EndRow@ = f_Proc@( 'EndRow' ) ;
      select ;
         when (EndRow@ = *null) ;
            return ;
         when (EndRow( i_Sheet : i_Row ) = *on) ;
            f_Escape( 1 : 'End Of Data' ) ;
      endsl;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc cb_EndRow ;
// --------------------------------------------------------------------------------------------------
