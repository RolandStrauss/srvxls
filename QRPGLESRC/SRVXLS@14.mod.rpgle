**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@14 : Microsoft Excel - Formula Functions.
//   Type : *MODULE
// Description :
//   Service procedures for common formulas used in Excel.
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
// 2019-01-29  Nico Basson       Initial Code.
// --------------------------------------------------------------------------------------------------
/INCLUDE SRVSRC,SRV@H
ctl-opt nomain
        thread(*concurrent)
        stgmdl(*inherit) ;
// Framework ---------------------------------------------------------------------------------------
/DEFINE  SRVXLS
/INCLUDE SRVSRC,SRV@P
/INCLUDE SRVSRC,SRVXLS@T
// --------------------------------------------------------------------------------------------------
// Globals <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-s RowNumber  int(10)                                                                    import ;
// --------------------------------------------------------------------------------------------------
// Column SUM() <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SUM export ;
   dcl-pi   xls_SUM ;
      i_StartRow   int(10) const ;
      i_EndRow     int(10) const ;
      i_FromCol    int(5)  const ;
      i_ToCol      int(5)  const options( *omit : *nopass ) ;
      i_TargetRow  int(10) const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s ToCol      int(5)  ;
   dcl-s TargetRow  int(10) ;
   // --------------------------------------------------------------------------------------------------
   dcl-s Column     int(5) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // From and To Columns =-=-=-=-=-=-=-=-=-=-=-=-=-
      ToCol = i_FromCol ;
      // -----------------------------------------------
      select ;
         when (%parms < %parmnum( i_ToCol )) ;
         when (%addr( i_ToCol ) <> *null) ;
            // ----------------------------------------
            ToCol = i_ToCol ;
      endsl;

      // Target row for Forumula =-=-=-=-=-=-=-=-=-=-=-
      TargetRow = i_EndRow + 1 ;
      // -----------------------------------------------
      select ;
         when (%parms < %parmnum( i_TargetRow )) ;
         when (%addr( i_TargetRow ) <> *null) ;
            // ----------------------------------------
            TargetRow = i_TargetRow ;
      endsl;
      // -----------------------------------------------
      if (RowNumber <> TargetRow) ;
         xls_AddRow( TargetRow ) ;
      endif ;

      // Sum a single column =-=-=-=-=-=-=-=-=-=-=-=-=-
      if (i_FromCol = ToCol) ;
         // --------------------------------------------
         xls_FormulaCell( 'SUM(' + xls_CellName( i_StartRow : i_FromCol ) +
                          ':' + xls_CellName( i_EndRow   : i_FromCol ) +
                          ')'
                     : i_FromCol
                     ) ;
         // --------------------------------------------
         return ;
         // --------------------------------------------
      endif;

      // Multiple -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      for Column = i_FromCol TO ToCol ;
         // -------------------------------------------
         xls_FormulaCell( 'SUM(' + xls_CellName( i_StartRow : Column ) +
                           ':' + xls_CellName( i_EndRow   : Column ) +
                           ')'
                      : Column
                      ) ;
         // -------------------------------------------
      endfor;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SUM ;
// --------------------------------------------------------------------------------------------------
