**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@11 : Microsoft Excel - Cells
//   Type : *MODULE
// Description :
//   Service procedures to handle cells in workbooks.
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
/INCLUDE SRVSRC,SRVXLS@T
// --------------------------------------------------------------------------------------------------
// Internal Functions Used ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
// Auto Size Column Width oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
dcl-pr f_AutoColWidth ;
end-pr ;
// --------------------------------------------------------------------------------------------------
// Java Prototypes jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
// --------------------------------------------------------------------------------------------------
// Create a new cell in a given row jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
dcl-pr j_CreateCell like(Cell) extproc( *java
                                      : Row_class
                                      : 'createCell'
                                      ) ;
   i_ColNumber like( jInt ) value ;
end-pr ;
// Retrieve Cell object from an existing row object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
dcl-pr j_getCell  like(Cell)  extproc( *java
                                     : Row_class
                                     : 'getCell'
                                     ) ;
   i_ColNumber like( jInt ) value ;
end-pr ;
// Set cell type jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
dcl-pr j_SetCellType extproc( *java
                            : Cell_class
                            : 'setCellType'
                            ) ;
   i_CellType like( jInt ) value ;
end-pr ;
// Cell Types --------------------------------------
//   Numeric     0
//   String      1
//   Formula     2
//   Blank       3
//   Boolean     4
//   Error       5
// Set cell value - String jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
dcl-pr j_SetCellString extproc( *java
                              : Cell_class
                              : 'setCellValue'
                              ) ;
   i_CellValue Like( jString ) const ;
end-pr ;
// Set cell value - Floating Point Numeric jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
dcl-pr j_SetCellFloat extproc( *java
                             : Cell_class
                             : 'setCellValue'
                             ) ;
   i_CellValue Like( jDouble ) value ;
end-pr ;
// Set cell value - Formula jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
dcl-pr j_SetCellFormula extproc( *java
                               : Cell_class
                               : 'setCellFormula'
                               ) ;
   i_CellValue Like( jString ) const ;
end-pr ;
// --------------------------------------------------------------------------------------------------
// Java Style Objects jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
// --------------------------------------------------------------------------------------------------
// Associate a CellStyle object with a given cell jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
dcl-pr j_setCellStyle extproc( *java
                             : Cell_class
                             : 'setCellStyle'
                             ) ;
   i_CellStyle   object( *java : CellStyle_class ) ;
end-pr ;
// User Defined ------------------------------------
dcl-s StringStyle  object( *java : CellStyle_class )                                        import ;
dcl-s NumberStyle  object( *java : CellStyle_class )                                        import ;
dcl-s DateStyle    object( *java : CellStyle_class )                                        import ;
dcl-s FormulaStyle object( *java : CellStyle_class )                                        import ;
// User Defined ------------------------------------
dcl-s UserStyle    object( *java : CellStyle_class )                                        import ;
// --------------------------------------------------------------------------------------------------
// Globals <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
// Sheet -------------------------------------------------------------------------------------------
dcl-s Sheet object ( *java
                   : Sheet_class
                   )                                                                        import ;
// Row ---------------------------------------------------------------------------------------------
dcl-s Row object( *java
                : Row_class
                )                                                                           import ;
// --------------------------------------------------
dcl-s RowNumber  int(10)                                                                    import ;
// Column ------------------------------------------------------------------------------------------
dcl-s ColNumber  int(5)                                                                     import ;
// Cell --------------------------------------------------------------------------------------------
dcl-s Cell object( *java
                 : Cell_class
                 ) ;
// Stateless ---------------------------------------------------------------------------------------
dcl-s NumberFormat varchar(1024)                                                            import ;
// Auto Column Width -------------------------------
dcl-s AutoWidth   ind                                                                       import ;
// --------------------------------------------------------------------------------------------------
// Insert Text Value into Cell <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_TextCell export ;
   dcl-pi   xls_TextCell int(5) ;
      i_Text         varchar(4096) const ;
      i_ColNumber    int(5)        const options( *omit : *nopass ) ;
      i_RowNumber    int(10)       const options( *omit : *nopass ) ;
      i_AutoWidth    ind           const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s TextString  like(jstring) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Column Number =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_ColNumber )) ;
            ColNumber += 1 ;
         when (%addr( i_ColNumber ) = *null) ;
            ColNumber += 1 ;
         when (i_ColNumber <> ColNumber) ;
            // ----------------------------------------
            ColNumber = i_ColNumber ;
      endsl;

      // Row Number -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_RowNumber )) ;
         when (%addr( i_RowNumber ) = *null) ;
         when (i_RowNumber <> RowNumber) ;
            // Switch Row ----------------------------
            xls_GetRow( i_RowNumber
                    : *on
                    ) ;
      endsl;

      // Autosize Column =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_AutoWidth )) ;
         when (%addr( i_AutoWidth ) <> *null) ;
            // ----------------------------------------
            AutoWidth = i_AutoWidth ;
      endsl;

      // Get Cell Object =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Cell = j_getCell( Row
                   : ColNumber
                   ) ;
      // Cell Not Found - Create ----------------------
      if (Cell = *null) ;
         Cell = j_CreateCell( Row
                         : ColNumber
                         ) ;
      endif;

      // Set Type to Text -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_SetCellType( Cell
                : 1
                ) ;

      // Set Value =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      TextString = java_NewString( i_Text ) ;
      // -----------------------------------------------
      j_SetCellString( Cell
                  : TextString
                  ) ;

      // Set Style =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (UserStyle <> *null) ;
            j_setCellStyle( Cell
                        : UserStyle
                        ) ;
            // Default ------------------------------------
         other ;
            j_setCellStyle( Cell
                        : StringStyle
                        ) ;
      endsl;

      // Autosize Column =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (AutoWidth) ;
         f_AutoColWidth() ;
      endif ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( TextString ) ;
      f_FreeLocalRef( Cell ) ;
      // -----------------------------------------------
      return ColNumber ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_TextCell ;
// --------------------------------------------------------------------------------------------------
// Insert Numeric Value into Cell <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_NumCell export ;
   dcl-pi   xls_NumCell int(5) ;
      i_Number       float(8)      value ;
      i_ColNumber    int(5)        const options( *omit : *nopass ) ;
      i_RowNumber    int(10)       const options( *omit : *nopass ) ;
      i_AutoWidth    ind           const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Column Number =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_ColNumber )) ;
            ColNumber += 1 ;
         when (%addr( i_ColNumber ) = *null) ;
            ColNumber += 1 ;
         when (i_ColNumber <> ColNumber) ;
            // ----------------------------------------
            ColNumber = i_ColNumber ;
      endsl;

      // Row Number -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_RowNumber )) ;
         when (%addr( i_RowNumber ) = *null) ;
         when (i_RowNumber <> RowNumber) ;
            // Switch Row ----------------------------
            xls_GetRow( i_RowNumber
                    : *on
                    ) ;
      endsl;

      // Autosize Column =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_AutoWidth )) ;
         when (%addr( i_AutoWidth ) <> *null) ;
            // ----------------------------------------
            AutoWidth = i_AutoWidth ;
      endsl;

      // Get Cell Object =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Cell = j_getCell( Row
                   : ColNumber
                   ) ;
      // Cell Not Found - Create ----------------------
      if (Cell = *null) ;
         Cell = j_CreateCell( Row
                         : ColNumber
                         ) ;
      endif;

      // Set Type to Numeric =-=-=-=-=-=-=-=-=-=-=-=-=-
      j_SetCellType( Cell
                : 0
                ) ;

      // Set Value =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_SetCellFloat( Cell
                 : i_Number
                 ) ;

      // Set Style =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (UserStyle <> *null) ;
            j_setCellStyle( Cell
                        : UserStyle
                        ) ;
            // Default ------------------------------------
         other ;
            j_setCellStyle( Cell
                        : NumberStyle
                        ) ;
      endsl;

      // Autosize Column =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (AutoWidth) ;
         f_AutoColWidth() ;
      endif ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( Cell ) ;
      // -----------------------------------------------
      return ColNumber ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_NumCell ;
// --------------------------------------------------------------------------------------------------
// Insert Date into Cell <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_DateCell export ;
   dcl-pi   xls_DateCell int(5) ;
      i_Date         date          value ;
      i_ColNumber    int(5)        const options( *omit : *nopass ) ;
      i_RowNumber    int(10)       const options( *omit : *nopass ) ;
      i_AutoWidth    ind           const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s FormatSave   varchar(1024) ;
   dcl-s DateFormat   varchar(1024) inz('yyyy-mm-dd') ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Column Number =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_ColNumber )) ;
            ColNumber += 1 ;
         when (%addr( i_ColNumber ) = *null) ;
            ColNumber += 1 ;
         when (i_ColNumber <> ColNumber) ;
            // ----------------------------------------
            ColNumber = i_ColNumber ;
      endsl;

      // Row Number -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_RowNumber )) ;
         when (%addr( i_RowNumber ) = *null) ;
         when (i_RowNumber <> RowNumber) ;
            // Switch Row ----------------------------
            xls_GetRow( i_RowNumber
                    : *on
                    ) ;
      endsl;

      // Autosize Column =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_AutoWidth )) ;
         when (%addr( i_AutoWidth ) <> *null) ;
            // ----------------------------------------
            AutoWidth = i_AutoWidth ;
      endsl;

      // Get Cell Object =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Cell = j_getCell( Row
                   : ColNumber
                   ) ;
      // Cell Not Found - Create ----------------------
      if (Cell = *null) ;
         Cell = j_CreateCell( Row
                         : ColNumber
                         ) ;
      endif;

      // Set Type to Numeric =-=-=-=-=-=-=-=-=-=-=-=-=-
      j_SetCellType( Cell
                : 0
                ) ;

      // Set Value =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_SetCellFloat( Cell
                 : s_ExcelDate( i_Date )
                 ) ;

      // Set Style =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (UserStyle <> *null) ;
            // Set to Date Format --------------------
            FormatSave = NumberFormat ;
            xls_SetNbrFormat( DateFormat ) ;
            j_setCellStyle( Cell
                        : UserStyle
                        ) ;
            // Reset ---------------------------------
            xls_SetNbrFormat( FormatSave ) ;
            // Default ------------------------------------
         other ;
            j_setCellStyle( Cell
                        : DateStyle
                        ) ;
      endsl;

      // Autosize Column =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (AutoWidth) ;
         f_AutoColWidth() ;
      endif ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( Cell ) ;
      // -----------------------------------------------
      return ColNumber ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_DateCell ;
// --------------------------------------------------------------------------------------------------
// Return Excel Date -------------------------------------------------------------------------------
// --------------------------------------------------------------------------------------------------
dcl-proc s_ExcelDate ;
   dcl-pi   s_ExcelDate like(jDouble) ;
      i_Date  Date value ;
   end-pi ;
   // --------------------------------------------------
   dcl-s Date  Date ;
   // --------------------------------------------------------------------------------------------------
   dcl-s StartDate Date inz(d'1900-01-01') ;
   dcl-s ExcelDays like(jDouble) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      Date = i_Date ;
      if (Date = *loval) ;
         Date = StartDate ;
      endif;
      // -----------------------------------------------
      ExcelDays = %diff( Date
                    : StartDate
                    : *days
                    ) + 1 ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      // Excel incorrectly thinks that 1900-02-29 is a valid date.
      if ( Date > d'1900-02-28' ) ;
         ExcelDays = ExcelDays + 1 ;
      endif ;
      // -----------------------------------------------
      return ExcelDays ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc s_ExcelDate ;
// --------------------------------------------------------------------------------------------------
// Insert Formula into Cell <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_FormulaCell export ;
   dcl-pi   xls_FormulaCell int(5) ;
      i_Formula      varchar(4096) const ;
      i_ColNumber    int(5)        const options( *omit : *nopass ) ;
      i_RowNumber    int(10)       const options( *omit : *nopass ) ;
      i_AutoWidth    ind           const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s FormulaString  like(jstring) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Column Number =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_ColNumber )) ;
            ColNumber += 1 ;
         when (%addr( i_ColNumber ) = *null) ;
            ColNumber += 1 ;
         when (i_ColNumber <> ColNumber) ;
            // ----------------------------------------
            ColNumber = i_ColNumber ;
      endsl;

      // Row Number -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_RowNumber )) ;
         when (%addr( i_RowNumber ) = *null) ;
         when (i_RowNumber <> RowNumber) ;
            // Switch Row ----------------------------
            xls_GetRow( i_RowNumber
                    : *on
                    ) ;
      endsl;

      // Autosize Column =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_AutoWidth )) ;
         when (%addr( i_AutoWidth ) <> *null) ;
            // ----------------------------------------
            AutoWidth = i_AutoWidth ;
      endsl;

      // Get Cell Object =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Cell = j_getCell( Row
                   : ColNumber
                   ) ;
      // Cell Not Found - Create ----------------------
      if (Cell = *null) ;
         Cell = j_CreateCell( Row
                         : ColNumber
                         ) ;
      endif;

      // Set Type to Formula =-=-=-=-=-=-=-=-=-=-=-=-=-
      j_SetCellType( Cell
                : 2
                ) ;

      // Set Value =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      FormulaString = java_NewString( i_Formula ) ;
      // -----------------------------------------------
      j_SetCellFormula( Cell
                   : FormulaString
                   ) ;

      // Set Style =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (UserStyle <> *null) ;
            j_setCellStyle( Cell
                        : UserStyle
                        ) ;
            // Default ------------------------------------
         other ;
            j_setCellStyle( Cell
                        : FormulaStyle
                        ) ;
      endsl;

      // Autosize Column =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (AutoWidth) ;
         f_AutoColWidth() ;
      endif ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( FormulaString ) ;
      f_FreeLocalRef( Cell ) ;
      // -----------------------------------------------
      return ColNumber ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_FormulaCell ;
// --------------------------------------------------------------------------------------------------
// Merge Cells on a Sheet <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_MergeCells export ;
   dcl-pi   xls_MergeCells ;
      i_RowFrom      int(5)       const options( *omit : *nopass ) ;
      i_ColFrom      int(5)       const options( *omit : *nopass ) ;
      i_RowTo        int(5)       const options( *omit : *nopass ) ;
      i_ColTo        int(5)       const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s RowFrom  int(5) ;
   dcl-s ColFrom  int(5) ;
   dcl-s RowTo    int(5) ;
   dcl-s ColTo    int(5) ;
   // Error Data --------------------------------------
   dcl-ds Char qualified ;
      To    char(20) ;
      From  char(20) ;
   end-ds ;
   // Create a new CellRangeAddress Object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_New_CellRangeAddress like(CellRangeAddress) extproc( *java
                                                            : CellRangeAddress_class
                                                            : *constructor
                                                            ) ;
      i_RowFrom  like( jInt ) value ;
      i_RowTo    like( jInt ) value ;
      i_ColFrom  like( jInt ) value ;
      i_ColTo    like( jInt ) value ;
   end-pr ;
   // --------------------------------------------------
   dcl-s CellRangeAddress object( *java
                             : CellRangeAddress_class
                             ) ;
   // --------------------------------------------------
   dcl-s MergeRegion             like(CellRangeAddress) ;
   // Merge all cells in sheet region jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_addMergedRegion  like(jInt)  extproc( *java
                                             : Sheet_class
                                             : 'addMergedRegion'
                                             ) ;
      i_MergeRegion like(CellRangeAddress) const ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // From and To =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      RowFrom = RowNumber ;
      select ;
         when (%parms < %parmnum( i_RowFrom )) ;
         when (%addr( i_RowFrom ) <> *null) ;
            // ----------------------------------------
            RowFrom = i_RowFrom ;
      endsl;
      // -----------------------------------------------
      ColFrom = ColNumber ;
      select ;
         when (%parms < %parmnum( i_ColFrom )) ;
         when (%addr( i_ColFrom ) <> *null) ;
            // ----------------------------------------
            ColFrom = i_ColFrom ;
      endsl;
      // -----------------------------------------------
      RowTo = RowNumber ;
      select ;
         when (%parms < %parmnum( i_RowTo )) ;
         when (%addr( i_RowTo ) <> *null) ;
            // ----------------------------------------
            RowTo = i_RowTo ;
      endsl;
      // -----------------------------------------------
      ColFrom = ColNumber ;
      select ;
         when (%parms < %parmnum( i_ColTo )) ;
         when (%addr( i_ColTo ) <> *null) ;
            // ----------------------------------------
            ColTo = i_ColTo ;
      endsl;

      // Validate Row and Column Values -=-=-=-=-=-=-=-
      if (RowTo < RowFrom) ;
         // --------------------------------------------
         Char.From = '"Row From"' ;
         Char.To   = '"Row To"' ;
         // --------------------------------------------
         err_SendEsc( 'XLS0001'
                 : 'SRVERR'
                 : Char
                 ) ;
      endif;
      // -----------------------------------------------
      if (ColTo < ColFrom) ;
         // --------------------------------------------
         Char.From = '"Column From"' ;
         Char.To   = '"Column To"' ;
         // --------------------------------------------
         err_SendEsc( 'XLS0001'
                 : 'SRVERR'
                 : Char
                 ) ;
      endif;

      // Create Region Object -=-=-=-=-=-=-=-=-=-=-=-=-
      MergeRegion = j_New_CellRangeAddress( RowFrom
                                       : RowTo
                                       : ColFrom
                                       : ColTo
                                       ) ;

      // Add Merged Region to Sheet -=-=-=-=-=-=-=-=-=-
      j_addMergedRegion( Sheet
                    : MergeRegion
                    ) ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( MergeRegion ) ;
      // Update the Current Row and Column Numbers ----
      RowNumber = RowTo ;
      ColNumber = ColTo ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_MergeCells ;
// --------------------------------------------------------------------------------------------------
// Return Cell Value <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_CellValue export ;
   dcl-pi   xls_CellValue varchar(1024) ;
      i_ColNumber    int(5)        const options( *omit : *nopass ) ;
      i_RowNumber    int(10)       const options( *omit : *nopass ) ;
      o_DataType     char(1)             options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s CellValue varchar(1024) ;
   dcl-s DataType  char(1) ;
   // Determine the type of data in a Cell object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_getCellType  like(jInt)  extproc( *java
                                         : Cell_class
                                         : 'getCellType'
                                         ) ;
   end-pr ;
   // Type Codes --------------------------------------
   // 0 =>> Numeric
   // 1 =>> String
   // 2 =>> Formula
   // 3 =>> Blank
   // 4 =>> Boolean
   // 5 =>> Error
   // --------------------------------------------------
   dcl-s CellType  int(10) ;
   // Retrieve String Value stored in a Cell object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_getStringValue  like(jString)  extproc( *java
                                               : Cell_class
                                               : 'getStringCellValue'
                                               ) ;
   end-pr ;
   // Retrieve Formula stored in a Cell object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_getFormula  like(jString)  extproc( *java
                                           : Cell_class
                                           : 'getCellFormula'
                                           ) ;
   end-pr ;
   // Retrieve Numeric Value stored in a Cell object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_getNumericValue  like(jDouble)  extproc( *java
                                                : Cell_class
                                                : 'getNumericCellValue'
                                                ) ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   dcl-s NumValue  float(8) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      clear DataType ;
      select ;
         when (%parms < %parmnum( o_DataType )) ;
         when (%addr( o_DataType ) <> *null) ;
            // ----------------------------------------
            clear o_DataType ;
      endsl;

      // Column Number =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_ColNumber )) ;
         when (%addr( i_ColNumber ) = *null) ;
         when (i_ColNumber <> ColNumber) ;
            // ----------------------------------------
            ColNumber = i_ColNumber ;
      endsl;

      // Row Number -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_RowNumber )) ;
         when (%addr( i_RowNumber ) = *null) ;
         when (i_RowNumber <> RowNumber) ;
            // Switch Row ----------------------------
            clear Row ;
      endsl;
      // -----------------------------------------------
      if (Row = *null) ;
         xls_GetRow( i_RowNumber
                : *on
                ) ;
      endif;

      // Get Cell Object =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Cell = j_getCell( Row
                   : ColNumber
                   ) ;
      // Cell Not Found -------------------------------
      if (Cell = *null) ;
         return *blank ;
      endif;

      // Get Cell Type =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      CellType = j_getCellType( Cell ) ;
      // Set Value ------------------------------------
      select ;
            // String -------------------------------------
         when (CellType = 1) ;
            CellValue = java_GetBytes( j_getStringValue( Cell ) ) ;
            DataType = 'C' ;
            // Formula ------------------------------------
         when (CellType = 2) ;
            CellValue = java_GetBytes( j_getFormula( Cell ) ) ;
            DataType = 'F' ;
            // Numeric ------------------------------------
         when (CellType = 3) ;
            NumValue = j_getNumericValue( Cell ) ;
            CellValue = %char( %dech( NumValue
                                  : 15
                                  : 2
                                  )
                           ) ;
            DataType = 'N' ;
            // Other type ---------------------------------
         other ;
            return *blank ;
      endsl;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( o_DataType )) ;
         when (%addr( o_DataType ) <> *null) ;
            // ----------------------------------------
            o_DataType = DataType ;
      endsl;
      // -----------------------------------------------
      return CellValue ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_CellValue ;
// --------------------------------------------------------------------------------------------------
// Switch Auto Size Column Width On or Off <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_SwitchAutoWidth export ;
   dcl-pi   xls_SwitchAutoWidth ;
      i_AutoWidth  ind const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      AutoWidth = i_AutoWidth ;
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_SwitchAutoWidth ;
// --------------------------------------------------------------------------------------------------
