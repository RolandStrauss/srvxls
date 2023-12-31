**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@02 : Microsoft Excel - Utilities
//   Type : *MODULE
// Description :
//   Utilities used in data extraction.
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
// 2018-06-14  Nico Basson       Initial Code.
// --------------------------------------------------------------------------------------------------
/INCLUDE SRVSRC,SRV@H
ctl-opt nomain
        thread(*concurrent)
        stgmdl(*inherit) ;
// Framework ---------------------------------------------------------------------------------------
/DEFINE  SRVXLS
/INCLUDE SRVSRC,SRV@P
// --------------------------------------------------------------------------------------------------
// Return the number of an alpha column (Microsoft Excel) <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
// Note that in Excel columns are counted starting from zero (0) .
// --------------------------------------------------------------------------------------------------
dcl-proc xls_Colnum export ;
   dcl-pi   xls_Colnum int(5) ;
      i_Column char(3) const ;
   end-pi ;
   // --------------------------------------------------
   dcl-s Colnum int(5) ;
   // --------------------------------------------------------------------------------------------------
   dcl-s  Alphabet char(26) inz('ABCDEFGHIJKLMNOPQRSTUVWXYZ') ;
   dcl-ds Column qualified ;
      Alpha1 char(1) ;
      Alpha2 char(1) ;
      Alpha3 char(1) ;
   end-ds;
   dcl-ds ColumnNum qualified ;
      Alpha1 int(10) ;
      Alpha2 int(10) ;
      Alpha3 int(10) ;
   end-ds;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      evalr Column = %trim( gen_uCase( i_Column ) ) ;
      clear ColumnNum ;

      // Calculate Cell Number =-=-=-=-=-=-=-=-=-=-=-=-
      if (Column.Alpha1 <> *blank) ;
         ColumnNum.Alpha1 = %scan( Column.Alpha1 : Alphabet ) * 676 ;
      endif ;
      if (Column.Alpha2 <> *blank) ;
         ColumnNum.Alpha2 = %scan( Column.Alpha2 : Alphabet ) * 26 ;
      endif ;
      if (Column.Alpha3 <> *blank) ;
         ColumnNum.Alpha3 = %scan( Column.Alpha3 : Alphabet ) ;
      endif ;
      // -----------------------------------------------
      Colnum = ColumnNum.Alpha1 + ColumnNum.Alpha2 + ColumnNum.Alpha3 ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      return Colnum - 1 ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_Colnum ;
// --------------------------------------------------------------------------------------------------
// Returns the Cell Name for POI y,x coordinates <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_CellName export ;
   dcl-pi   xls_CellName varchar(16) ;
      i_Row  int(5) value ;
      i_Col  int(5) value ;
   end-pi ;
   // --------------------------------------------------
   dcl-s CellName varchar(16) ;
   // --------------------------------------------------------------------------------------------------
   dcl-ds Alphabet qualified static ;
      whole  char(26) inz('ABCDEFGHIJKLMNOPQRSTUVWXYZ') ;
      letter char(1)  dim(26) overlay(whole) ;
   end-ds ;
   // --------------------------------------------------
   dcl-s Remain int(5) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      i_Row += 1 ;
      i_Col += 1 ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      clear CellName ;
      // -----------------------------------------------
      dou (i_Col = 0) ;
         // -------------------------------------------
         Remain = %rem( i_Col : 26 ) ;
         i_Col  = %div( i_Col : 26 ) ;
         if ( Remain = 0 ) ;
            Remain = 26 ;
            i_Col -= 1 ;
         endif ;
         // -------------------------------------------
         CellName = Alphabet.letter( Remain ) + CellName ;
         // -------------------------------------------
      enddo ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      CellName = CellName + %char( i_Row ) ;
      // -----------------------------------------------
      return CellName ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_CellName ;
// --------------------------------------------------------------------------------------------------
