**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@12 : Microsoft Excel - Style and Formatting
//   Type : *MODULE
// Description :
//   Service procedures to set style and formatting of rows and cells.
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
// 2019-01-04  Nico Basson       Initial Code.
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
// Java Style Objects jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
// --------------------------------------------------------------------------------------------------
// Create a new CellStyle object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
dcl-pr j_CreateCellStyle object( *java : CellStyle_class )
                         extproc( *java
                                : Workbook_class
                                : 'createCellStyle'
                                ) ;
end-pr ;
// Defaults ----------------------------------------
dcl-s StringStyle  object( *java : CellStyle_class )                                        export ;
dcl-s NumberStyle  object( *java : CellStyle_class )                                        export ;
dcl-s DateStyle    object( *java : CellStyle_class )                                        export ;
dcl-s FormulaStyle object( *java : CellStyle_class )                                        export ;
// User Defined ------------------------------------
dcl-s UserStyle    object( *java : CellStyle_class )                                        export ;
// --------------------------------------------------------------------------------------------------
// Globals <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
// Workbook ----------------------------------------------------------------------------------------
dcl-s XSSF_Workbook object( *java
                          : XSSF_Workbook_class
                          )                                                                 import ;
// Sheet -------------------------------------------------------------------------------------------
dcl-s Sheet object ( *java
                   : Sheet_class
                   )                                                                        import ;
// --------------------------------------------------
dcl-s SheetName  varchar(128)                                                               import ;
// Column ------------------------------------------------------------------------------------------
dcl-s ColNumber  int(5)                                                                     import ;
// --------------------------------------------------------------------------------------------------
// Java Style Objects (Stateless) jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
// --------------------------------------------------------------------------------------------------
dcl-s NumberFormat varchar(1024)                                                            export ;
dcl-s ColumnWidth  int(5)                                                                   export ;
// Font --------------------------------------------
dcl-s Bold        ind         inz ;
dcl-s Italic      ind         inz ;
dcl-s Underline   ind         inz ;
dcl-s FontSize    int(5)      inz ;
dcl-s FontName    varchar(64) inz ;
dcl-s FontRGB     char(11)    inz ;
dcl-s StrikeOut   ind         inz ;
// Alignment ---------------------------------------
dcl-s Align       char(1)     inz ;
// Fill --------------------------------------------
dcl-s Pattern     int(5)      inz                                                           export ;
dcl-s BackRGB     char(11)    inz                                                           export ;
dcl-s ForeRGB     char(11)    inz                                                           export ;
// Auto Column Width -------------------------------
dcl-s AutoWidth   ind                                                                       import ;
dcl-s UseMerged   ind         inz                                                           export ;
// --------------------------------------------------------------------------------------------------
// Set Column Width <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_ColWidth export ;
   dcl-pi   xls_ColWidth ;
      i_Width        int(5)       const ;
      i_ColNumber    int(5)       const options( *omit : *nopass ) ;
      i_SheetName    varchar(128) const options( *omit : *nopass ) ;
   end-pi ;
   // Create a new cell in a given row jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_setColumnWidth  extproc( *java
                                : Sheet_class
                                : 'setColumnWidth'
                                ) ;
      i_ColNumber like( jInt ) value ;
      i_Width     like( jInt ) value ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      ColumnWidth = i_Width ;
      // -----------------------------------------------
      if (ColumnWidth > 255) ;
         ColumnWidth = 255 ;
      endif;

      // Column Number =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_ColNumber )) ;
         when (%addr( i_ColNumber ) = *null) ;
         when (i_ColNumber <> ColNumber) ;
            // ----------------------------------------
            ColNumber = i_ColNumber ;
      endsl;

      // Sheet =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_SheetName )) ;
         when (%addr( i_SheetName ) = *null) ;
         when (i_SheetName <> SheetName) ;
            // ----------------------------------------
            SheetName = i_SheetName ;
      endsl;

      // Set Column Width -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_setColumnWidth( Sheet
                   : ColNumber
                   : ColumnWidth * 256
                   ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_ColWidth ;
// --------------------------------------------------------------------------------------------------
// Auto Size Column Width <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_AutoColWidth export ;
   dcl-pi   xls_AutoColWidth ;
      i_Switch       ind    const ;
      i_UseMerged    ind    const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Use Merged Cells to Calculate =-=-=-=-=-=-=-=-
      UseMerged = *off ;
      select ;
         when (%parms < %parmnum( i_UseMerged )) ;
         when (%addr( i_UseMerged ) <> *null) ;
            // ----------------------------------------
            UseMerged = i_UseMerged ;
      endsl;

      // Switch on or off -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      AutoWidth = i_Switch ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_AutoColWidth ;
// --------------------------------------------------------------------------------------------------
// Clear Style (Delete User Style Object) <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_ClearStyle export ;
   dcl-pi   xls_ClearStyle ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      if (UserStyle <> *null) ;
         f_FreeLocalRef( UserStyle ) ;
         UserStyle = *null ;
      endif ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_ResetStyle() ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_ClearStyle ;
// --------------------------------------------------------------------------------------------------
// Set Font <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SetFont export ;
   dcl-pi   xls_SetFont ;
      i_Bold        ind           const options( *omit : *nopass ) ;
      i_Italic      ind           const options( *omit : *nopass ) ;
      i_Underline   ind           const options( *omit : *nopass ) ;
      i_FontSize    int(5)        const options( *omit : *nopass ) ;
      i_FontName    varchar(64)   const options( *omit : *nopass ) ;
      i_FontRGB     char(11)      const options( *omit : *nopass ) ;
      i_StrikeOut   ind           const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Parameters -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_Bold )) ;
         when (%addr( i_Bold ) <> *null) ;
            // ----------------------------------------
            if (Bold <> i_Bold) ;
               s_StyleChange() ;
            endif;
            Bold = i_Bold ;
      endsl;
      select ;
         when (%parms < %parmnum( i_Italic )) ;
         when (%addr( i_Italic ) <> *null) ;
            // ----------------------------------------
            if (Italic <> i_Italic) ;
               s_StyleChange() ;
            endif;
            Italic = i_Italic ;
      endsl;
      select ;
         when (%parms < %parmnum( i_Underline )) ;
         when (%addr( i_Underline ) <> *null) ;
            // ----------------------------------------
            if (Underline <> i_Underline) ;
               s_StyleChange() ;
            endif;
            Underline = i_Underline ;
      endsl;
      select ;
         when (%parms < %parmnum( i_FontSize )) ;
         when (%addr( i_FontSize ) <> *null) ;
            // ----------------------------------------
            if (FontSize <> i_FontSize) ;
               s_StyleChange() ;
            endif;
            FontSize = i_FontSize ;
      endsl;
      select ;
         when (%parms < %parmnum( i_FontName )) ;
         when (%addr( i_FontName ) <> *null) ;
            // ----------------------------------------
            if (FontName <> i_FontName) ;
               s_StyleChange() ;
            endif;
            FontName = i_FontName ;
      endsl;
      select ;
         when (%parms < %parmnum( i_FontRGB )) ;
         when (%addr( i_FontRGB ) <> *null) ;
            // ----------------------------------------
            if (FontRGB <> i_FontRGB) ;
               s_StyleChange() ;
            endif;
            FontRGB = i_FontRGB ;
      endsl;
      select ;
         when (%parms < %parmnum( i_StrikeOut )) ;
         when (%addr( i_StrikeOut ) <> *null) ;
            // ----------------------------------------
            if (StrikeOut <> i_StrikeOut) ;
               s_StyleChange() ;
            endif;
            StrikeOut = i_StrikeOut ;
      endsl;

      // UserStyle Object -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (UserStyle = *null) ;
         UserStyle = j_CreateCellStyle( XSSF_Workbook ) ;
      endif;

      // Apply Style to CellStyle Object =-=-=-=-=-=-=-
      f_StyleFont( UserStyle
              : Bold
              : Italic
              : Underline
              : FontSize
              : FontName
              : FontRGB
              : StrikeOut
              ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SetFont ;
// --------------------------------------------------------------------------------------------------
// Set Number Format <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SetNbrFormat export ;
   dcl-pi   xls_SetNbrFormat ;
      i_NbrFmt  varchar(1024) const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Change in Number Format =-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (NumberFormat = *blank) ;
         when (NumberFormat <> i_NbrFmt) ;
            // ----------------------------------------
            s_StyleChange() ;
      endsl;
      // -----------------------------------------------
      NumberFormat = i_NbrFmt ;

      // UserStyle Object -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (UserStyle = *null) ;
         UserStyle = j_CreateCellStyle( XSSF_Workbook ) ;
      endif;

      // Apply Style to CellStyle Object =-=-=-=-=-=-=-
      f_NumberFormat( UserStyle
                 : NumberFormat
                 ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SetNbrFormat ;
// --------------------------------------------------------------------------------------------------
// Set Alignment <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SetAlignment export ;
   dcl-pi   xls_SetAlignment ;
      i_Align  char(1) const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Change in Alignment =-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (Align = *blank) ;
         when (Align <> i_Align) ;
            // ----------------------------------------
            s_StyleChange() ;
      endsl;
      // -----------------------------------------------
      Align = i_Align ;

      // UserStyle Object -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (UserStyle = *null) ;
         UserStyle = j_CreateCellStyle( XSSF_Workbook ) ;
      endif;

      // Apply Style to CellStyle Object =-=-=-=-=-=-=-
      f_Alignment( UserStyle
              : Align
              ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SetAlignment ;
// --------------------------------------------------------------------------------------------------
// Set Text Wrapping <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SetWrapText export ;
   dcl-pi   xls_SetWrapText ;
      i_WrapText  ind const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // UserStyle Object -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (UserStyle = *null) ;
         UserStyle = j_CreateCellStyle( XSSF_Workbook ) ;
      endif;

      // Apply Style to CellStyle Object =-=-=-=-=-=-=-
      f_WrapText( UserStyle
             : i_WrapText
             ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SetWrapText ;
// --------------------------------------------------------------------------------------------------
// Set Cell Borders <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SetBorder export ;
   dcl-pi   xls_SetBorder ;
      i_BorderTop   ind    const options( *omit : *nopass ) ;
      i_BorderBot   ind    const options( *omit : *nopass ) ;
      i_BorderLeft  ind    const options( *omit : *nopass ) ;
      i_BorderRight ind    const options( *omit : *nopass ) ;
      i_StyleTop    int(3) const options( *omit : *nopass ) ;
      i_StyleBot    int(3) const options( *omit : *nopass ) ;
      i_StyleLeft   int(3) const options( *omit : *nopass ) ;
      i_StyleRight  int(3) const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s BorderTop    ind    ;
   dcl-s BorderBot    ind    ;
   dcl-s BorderLeft   ind    ;
   dcl-s BorderRight  ind    ;
   dcl-s StyleTop     int(3) ;
   dcl-s StyleBot     int(3) ;
   dcl-s StyleLeft    int(3) ;
   dcl-s StyleRight   int(3) ;
   // Set Cell Borders Cell Style Java Object ooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
   // Line Styles -------------------------------------
   // Dash Dot     (-.)          9                    1
   // Dash Dot Dot (-..)        11                    2
   // Dashed       (-)           3                    3
   // Dotted       (...)         7                    4
   // Double       (=)           6                    5
   // Hair         (?)           4                    6
   // Medium                     2                    7
   // Medium Dash Dot           10                    8
   // Medium Dash Dot Dot       12                    9
   // Medium Dashed              8                   10
   // Thin                       1                   11
   // None                       0                   12
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      BorderTop = *off ;
      select ;
         when (%parms < %parmnum( i_BorderTop )) ;
         when (%addr( i_BorderTop ) <> *null) ;
            // ----------------------------------------
            BorderTop = i_BorderTop ;
      endsl;
      BorderBot = *off ;
      select ;
         when (%parms < %parmnum( i_BorderBot )) ;
         when (%addr( i_BorderBot ) <> *null) ;
            // ----------------------------------------
            BorderBot = i_BorderBot ;
      endsl;
      BorderLeft = *off ;
      select ;
         when (%parms < %parmnum( i_BorderLeft )) ;
         when (%addr( i_BorderLeft ) <> *null) ;
            // ----------------------------------------
            BorderLeft = i_BorderLeft ;
      endsl;
      BorderRight = *off ;
      select ;
         when (%parms < %parmnum( i_BorderRight )) ;
         when (%addr( i_BorderRight ) <> *null) ;
            // ----------------------------------------
            BorderRight = i_BorderRight ;
      endsl;
      // -----------------------------------------------
      clear StyleTop ;
      select ;
         when (%parms < %parmnum( i_StyleTop )) ;
         when (%addr( i_StyleTop ) <> *null) ;
            // ----------------------------------------
            StyleTop = i_StyleTop ;
      endsl;
      clear StyleBot ;
      select ;
         when (%parms < %parmnum( i_StyleBot )) ;
         when (%addr( i_StyleBot ) <> *null) ;
            // ----------------------------------------
            StyleBot = i_StyleBot ;
      endsl;
      clear StyleLeft ;
      select ;
         when (%parms < %parmnum( i_StyleLeft )) ;
         when (%addr( i_StyleLeft ) <> *null) ;
            // ----------------------------------------
            StyleLeft = i_StyleLeft ;
      endsl;
      clear StyleRight ;
      select ;
         when (%parms < %parmnum( i_StyleRight )) ;
         when (%addr( i_StyleRight ) <> *null) ;
            // ----------------------------------------
            StyleRight = i_StyleRight ;
      endsl;

      // UserStyle Object -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (UserStyle = *null) ;
         UserStyle = j_CreateCellStyle( XSSF_Workbook ) ;
      endif;

      // Apply Style to CellStyle Object =-=-=-=-=-=-=-
      f_CellBorder( UserStyle
               : BorderTop
               : BorderBot
               : BorderLeft
               : BorderRight
               : StyleTop
               : StyleBot
               : StyleLeft
               : StyleRight
               ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SetBorder ;
// --------------------------------------------------------------------------------------------------
// Set Cell Fill Colours and Pattern <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SetCellFill export ;
   dcl-pi   xls_SetCellFill ;
      i_Pattern    int(5)   const options( *omit : *nopass ) ;
      i_BackRGB    char(11) const options( *omit : *nopass ) ;
      i_ForeRGB    char(11) const options( *omit : *nopass ) ;
   end-pi ;
   // Set Cell Background Colour and Pattern oooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
   // Fill Pattern ------------------------------------
   // PATTERN_NO_FILL                  0
   // PATTERN_SOLID_FOREGROUND         1
   // PATTERN_FINE_DOTS                2
   // PATTERN_ALT_BARS                 3
   // PATTERN_SPARSE_DOTS              4
   // THICK_HORZ_BANDS                 5
   // THICK_VERT_BANDS                 6
   // THICK_BACKWARD_DIAG              7
   // THICK_FORWARD_DIAG               8
   // PATTERN_BIG_SPOTS                9
   // PATTERN_BRICKS                   10
   // THIN_HORZ_BANDS                  11
   // THIN_VERT_BANDS                  12
   // THIN_BACKWARD_DIAG               13
   // THIN_FORWARD_DIAG                14
   // PATTERN_SQUARES                  15
   // PATTERN_DIAMONDS                 16
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      select ;
         when (%parms < %parmnum( i_Pattern )) ;
         when (%addr( i_Pattern ) <> *null) ;
            // ----------------------------------------
            if (Pattern <> i_Pattern) ;
               s_StyleChange() ;
            endif;
            Pattern = i_Pattern ;
      endsl;
      select ;
         when (%parms < %parmnum( i_BackRGB )) ;
         when (%addr( i_BackRGB ) <> *null) ;
            // ----------------------------------------
            if (BackRGB <> i_BackRGB) ;
               s_StyleChange() ;
            endif;
            BackRGB = i_BackRGB ;
      endsl;
      select ;
         when (%parms < %parmnum( i_ForeRGB )) ;
         when (%addr( i_ForeRGB ) <> *null) ;
            // ----------------------------------------
            if (ForeRGB <> i_ForeRGB) ;
               s_StyleChange() ;
            endif;
            ForeRGB = i_ForeRGB ;
      endsl;

      // UserStyle Object -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (UserStyle = *null) ;
         UserStyle = j_CreateCellStyle( XSSF_Workbook ) ;
      endif;

      // Apply Style to CellStyle Object =-=-=-=-=-=-=-
      f_CellFill( UserStyle
             : Pattern
             : BackRGB
             : ForeRGB
             ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SetCellFill ;
// --------------------------------------------------------------------------------------------------
// Create Default Style Objects oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_CrtDftStyles export ;
   dcl-pi   f_CrtDftStyles ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // CellStyle Objects-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      StringStyle  = j_CreateCellStyle( XSSF_Workbook ) ;
      NumberStyle  = j_CreateCellStyle( XSSF_Workbook ) ;
      DateStyle    = j_CreateCellStyle( XSSF_Workbook ) ;
      FormulaStyle = j_CreateCellStyle( XSSF_Workbook ) ;

      // String -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_Alignment( StringStyle
              : 'L'
              ) ;

      // Numeric =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_Alignment( NumberStyle
              : 'R'
              ) ;
      f_NumberFormat( NumberStyle
                 : '# ##0.00'
                 ) ;

      // Date -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_Alignment( DateStyle
              : 'L'
              ) ;
      f_NumberFormat( DateStyle
                 : 'yyyy-mm-dd'
                 ) ;

      // Formula =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_Alignment( FormulaStyle
              : 'L'
              ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_CrtDftStyles ;
// --------------------------------------------------------------------------------------------------
// Handle a Change in Style ------------------------------------------------------------------------
// --------------------------------------------------------------------------------------------------
dcl-proc s_StyleChange ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // UserStyle Object not found -=-=-=-=-=-=-=-=-=-
      if (UserStyle = *null) ;
         return ;
      endif;

      // Local Reference =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( UserStyle ) ;

      // Java Object (pointer) =-=-=-=-=-=-=-=-=-=-=-=-
      UserStyle = *null ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc s_StyleChange ;
// --------------------------------------------------------------------------------------------------
