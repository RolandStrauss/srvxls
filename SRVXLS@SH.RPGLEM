**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@SH : Microsoft Excel - Style Handler
//   Type : *MODULE
// Description :
//   These procedures handles Java style objects.
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
// Colour Codes ------------------------------------------------------------------------------------
// AQUA...                       49
// BLACK...                      8
// BLUE...                       12
// BLUE_GREY...                  54
// BRIGHT_GREEN...               11
// BROWN...                      60
// CORAL...                      29
// CORNFLOWER_BLUE...            24
// DARK_BLUE...                  18
// DARK_RED...                   16
// DARK_TEAL...                  56
// DARK_YELLOW...                19
// DARK_GOLD...                  51
// DARK_GREEN...                 17
// GREY_25...                    22
// GREY_40...                    55
// GREY_50...                    23
// GREY_80...                    63
// INDIGO...                     62
// LAVENDER...                   46
// LEMON_CHIFFON...              26
// LIGHT_BLUE...                 48
// LIGHT_CORNFLOWER_BLUE...      31
// LIGHT_GREEN...                42
// LIGHT_ORANGE...               52
// LIGHT_TURQUOISE...            27
// LIGHT_YELLOW...               43
// LIME...                       50
// MAROON...                     25
// OLIVE_GREEN...                59
// ORANGE...                     53
// ORCHID...                     28
// PALE_BLUE...                  44
// PINK...                       14
// PLUM...                       61
// RED...                        10
// ROSE...                       45
// ROYAL_BLUE...                 30
// SEA_GREEN...                  57
// SKY_BLUE...                   40
// TAN...                        47
// TEAL...                       21
// TURQUOISE...                  15
// VIOLET...                     20
// WHITE...                      9
// YELLOW...                     13
// NORMAL...                     32767
// AUTOMATIC...                  64           - DEFAULT
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
// Workbook ----------------------------------------------------------------------------------------
dcl-s XSSF_Workbook object( *java
                          : XSSF_Workbook_class
                          )                                                                 import ;
// Sheet -------------------------------------------------------------------------------------------
dcl-s Sheet object ( *java
                   : Sheet_class
                   )                                                                        import ;
// Column ------------------------------------------------------------------------------------------
dcl-s ColNumber  int(5)                                                                     import ;
// --------------------------------------------------------------------------------------------------
// Java Style Objects (Stateless) jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
// --------------------------------------------------------------------------------------------------
// Data Format -------------------------------------------------------------------------------------
dcl-s DataFormat object( *java
                       : DataFormat_class
                       ) ;
// --------------------------------------------------
dcl-s FormatInt  int(5) ;
// Font --------------------------------------------------------------------------------------------
dcl-s FontStyle object( *java
                      : Font_class
                      ) ;
// Alignment ---------------------------------------------------------------------------------------
dcl-s AlignInt    int(5) ;
// Text Wrapping -----------------------------------------------------------------------------------
dcl-s WrapText    ind ;
// Cell Borders ------------------------------------------------------------------------------------
dcl-s BorderTop   ind    inz ;
dcl-s BorderBot   ind    inz ;
dcl-s BorderLeft  ind    inz ;
dcl-s BorderRight ind    inz ;
dcl-s StyleTop    int(3) inz ;
dcl-s StyleBot    int(3) inz ;
dcl-s StyleLeft   int(3) inz ;
dcl-s StyleRight  int(3) inz ;
// Fill --------------------------------------------------------------------------------------------
dcl-s Pattern     int(5) inz ;
dcl-s BackRGB     char(11) inz ;
dcl-s ForeRGB     char(11) inz ;
// Auto Column Width -------------------------------------------------------------------------------
dcl-s AutoWidth   ind                                                                       import ;
dcl-s UseMerged   ind                                                                       import ;
// --------------------------------------------------------------------------------------------------
// Set Number Format in Cell Style Java Object ooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_NumberFormat export ;
   dcl-pi   f_NumberFormat ;
      b_CellStyle  object( *java : CellStyle_class ) ;
      i_NbrFmt     varchar(1024)                     const ;
   end-pi ;
   // Create a DataFormat object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_CreateDataFormat like(DataFormat) extproc( *java
                                                  : Workbook_class
                                                  : 'createDataFormat'
                                                  ) ;
   end-pr ;
   // --------------------------------------------------
   dcl-s FormatString  like(jString) ;
   // Get the Internal representation of a data format jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_GetFormat like(jShort) extproc( *java
                                       : DataFormat_class
                                       : 'getFormat'
                                       ) ;
      i_FormatString  like(jString) ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Create DataFormat Java Object =-=-=-=-=-=-=-=-
      if (DataFormat <> *null) ;
         f_FreeLocalRef( DataFormat ) ;
      endif ;
      // -----------------------------------------------
      DataFormat = j_CreateDataFormat( XSSF_Workbook ) ;

      // Data Format Internal Representation =-=-=-=-=-
      FormatString = java_NewString( i_NbrFmt ) ;
      FormatInt    = j_GetFormat( DataFormat
                             : FormatString
                             ) ;

      // Apply All Current Styles to Object -=-=-=-=-=-
      s_ApplyStyles( b_CellStyle ) ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( FormatString ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_NumberFormat ;
// --------------------------------------------------------------------------------------------------
// Set Font Style in Cell Style Java Object oooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_StyleFont export ;
   dcl-pi   f_StyleFont ;
      b_CellStyle   object( *java : CellStyle_class ) ;
      i_Bold        ind                               value ;
      i_Italic      ind                               value ;
      i_Underline   ind                               value ;
      i_FontSize    int(5)                            value ;
      i_FontName    varchar(64)                       value ;
      i_FontRGB     char(11)                          const ;
      i_StrikeOut   ind                               value ;
   end-pi ;
   // Create a Font Style object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_CreateFontStyle like(FontStyle) extproc( *java
                                                : Workbook_class
                                                : 'createFont'
                                                ) ;
   end-pr ;
   // Set Font Weight jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetFontWeight  extproc( *java
                               : Font_class
                               : 'setBoldweight'
                               ) ;
      i_FontWeight  like(jShort) value ;
   end-pr ;
   // --------------------------------------------------
   dcl-s FontWeight  int(5) ;
   // Set Font Height in points jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetFontHeight  extproc( *java
                               : Font_class
                               : 'setFontHeightInPoints'
                               ) ;
      i_FontWeight  like(jShort) value ;
   end-pr ;
   // Set Font to Italic jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetItalic  extproc( *java
                           : Font_class
                           : 'setItalic'
                           ) ;
      i_Italic  ind value ;
   end-pr ;
   // Underline jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetUnderline  extproc( *java
                              : Font_class
                              : 'setUnderline'
                              ) ;
      i_Underline  char(1) value ;
   end-pr ;
   // Set Font Face by Name jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_setFontName  extproc( *java
                             : Font_class
                             : 'setFontName'
                             ) ;
      i_FontName  like(jString) ;
   end-pr ;
   // --------------------------------------------------
   dcl-s NameString  like(jString) ;
   // Set Font Colour jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetColour  extproc( *java
                           : XSSF_Font_class
                           : 'setColor'
                           ) ;
      i_Color  object( *java : 'org.apache.poi.xssf.usermodel.XSSFColor' ) const ;
   end-pr ;
   // Set Strikeout on/off jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetStrikeout  extproc( *java
                           : Font_class
                           : 'setStrikeout'
                           ) ;
      i_StrikeOut  ind value ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Create FontStyle Java Object -=-=-=-=-=-=-=-=-
      if (FontStyle <> *null) ;
         f_FreeLocalRef( FontStyle ) ;
      endif ;
      // -----------------------------------------------
      FontStyle = j_CreateFontStyle( XSSF_Workbook ) ;

      // Font Weight =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      FontWeight = 190 ;
      if (i_Bold) ;
         FontWeight = 700 ;
      endif ;
      // -----------------------------------------------
      j_SetFontWeight( FontStyle
                  : FontWeight
                  ) ;

      // Font Size (Height pt) =-=-=-=-=-=-=-=-=-=-=-=-
      if (i_FontSize > 0) ;
         j_SetFontHeight( FontStyle
                     : i_FontSize
                     ) ;
      endif ;

      // Italic -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (i_Italic) ;
         j_SetItalic( FontStyle
                 : *on
                 ) ;
      endif ;

      // Underline =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (i_Underline) ;
         j_SetUnderline( FontStyle
                    : *on
                    ) ;
      endif ;

      // Font Face Name -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (i_FontName <> *blank) ;
         // --------------------------------------------
         NameString = java_NewString( i_FontName ) ;
         j_setFontName( FontStyle
                   : NameString
                   ) ;
         f_FreeLocalRef( NameString ) ;
         // --------------------------------------------
      endif;

      // Font Colour =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (i_FontRGB <> *blank) ;
         j_SetColour( FontStyle
                 : s_XSSFColor( i_FontRGB )
                 ) ;
      endif ;

      // Strikeout =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (i_StrikeOut) ;
         j_SetStrikeout( FontStyle
                    : *on
                    ) ;
      endif ;

      // Apply All Current Styles to Object -=-=-=-=-=-
      s_ApplyStyles( b_CellStyle ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_StyleFont ;
// --------------------------------------------------------------------------------------------------
// Returns Java Colour Object for Font Colour ------------------------------------------------------
// --------------------------------------------------------------------------------------------------
dcl-proc s_XSSFColor ;
   dcl-pi   s_XSSFColor object( *java : Colour_class ) ;
      i_FontRGB     char(11) const ;
   end-pi ;
   // --------------------------------------------------
   dcl-s R int(5) ;
   dcl-s G int(5) ;
   dcl-s B int(5) ;
   // --------------------------------------------------
   dcl-s Start int(5) ;
   dcl-s End   int(5) ;
   // Java Class Constants jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-c COLOUR_CLASS          'java.awt.Color'                                    ;
   dcl-c XSSFCOLOUR_CLASS      'org.apache.poi.xssf.usermodel.XSSFColor'           ;
   // Create a new Colour Object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_NewColour object ( *java : Colour_class )
                   extproc( *java
                          : Colour_class
                          : *constructor
                          ) ;
      i_R  like( jInt ) value ;
      i_G  like( jInt ) value ;
      i_B  like( jInt ) value ;
   end-pr ;
   // --------------------------------------------------
   dcl-s Color  object( *java : 'java.awt.Color' ) ;
   // Create a new XSSF Colour Object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_NewXSSFColour object ( *java : XSSFColour_class )
                       extproc( *java
                              : XSSFColour_class
                              : *constructor
                              ) ;
      i_Color  object( *java : 'java.awt.Color' ) const ;
   end-pr ;
   // --------------------------------------------------
   dcl-s XSSFColor  object( *java : 'org.apache.poi.xssf.usermodel.XSSFColor' ) ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      Start = 1 ;
      End   = %scan( ',' : i_FontRGB ) ;
      R = %int( %subst( i_FontRGB
                   : Start
                   : End - 1
                   )
           ) ;
      // -----------------------------------------------
      Start = End + 1 ;
      End   = %scan( ',' : i_FontRGB : Start ) ;
      G = %int( %subst( i_FontRGB
                   : Start
                   : End - 1
                   )
           ) ;
      // -----------------------------------------------
      Start = End + 1 ;
      B = %int( %subst( i_FontRGB
                   : Start
                   )
           ) ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Color = j_NewColour( R
                      : G
                      : B
                      ) ;
      // -----------------------------------------------
      XSSFColor = j_NewXSSFColour( Color ) ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( Color ) ;
      // -----------------------------------------------
      return XSSFColor ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc s_XSSFColor ;
// --------------------------------------------------------------------------------------------------
// Set Text Alignment in Cell Style Java Object oooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_Alignment export ;
   dcl-pi   f_Alignment ;
      b_CellStyle  object( *java : CellStyle_class ) ;
      i_Align      char(1)                           const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Set Integer =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
            // C - Center (2) -----------------------------
         when (i_Align  = 'C') ;
            AlignInt = 2 ;
            // S - Center Selection (6) -------------------
         when (i_Align  = 'S') ;
            AlignInt = 6 ;
            // F - Fill (4) -------------------------------
         when (i_Align  = 'F') ;
            AlignInt = 4 ;
            // G - General (0) ----------------------------
         when (i_Align  = 'G') ;
            AlignInt = 0 ;
            // J - Justify (5) ----------------------------
         when (i_Align  = 'J') ;
            AlignInt = 5 ;
            // L - Left (1) -------------------------------
         when (i_Align  = 'L') ;
            AlignInt = 1 ;
            // R - Right (3) ------------------------------
         when (i_Align  = 'R') ;
            AlignInt = 3 ;
            // Default (General) --------------------------
         other ;
            AlignInt = 0 ;
      endsl;

      // Apply All Current Styles to Object -=-=-=-=-=-
      s_ApplyStyles( b_CellStyle ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_Alignment ;
// --------------------------------------------------------------------------------------------------
// Set Text Wrapping in Cell Style Java Object ooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_WrapText export ;
   dcl-pi   f_WrapText ;
      b_CellStyle  object( *java : CellStyle_class ) ;
      i_WrapText   ind                               const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      WrapText = i_WrapText ;

      // Apply All Current Styles to Object -=-=-=-=-=-
      s_ApplyStyles( b_CellStyle ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_WrapText ;
// --------------------------------------------------------------------------------------------------
// Set Cell Borders Cell Style Java Object ooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
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
dcl-proc f_CellBorder export ;
   dcl-pi   f_CellBorder ;
      b_CellStyle   object( *java : CellStyle_class ) ;
      i_BorderTop   ind                              const ;
      i_BorderBot   ind                              const ;
      i_BorderLeft  ind                              const ;
      i_BorderRight ind                              const ;
      i_StyleTop    int(3)                           const ;
      i_StyleBot    int(3)                           const ;
      i_StyleLeft   int(3)                           const ;
      i_StyleRight  int(3)                           const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      BorderTop   = i_BorderTop   ;
      BorderBot   = i_BorderBot   ;
      BorderLeft  = i_BorderLeft  ;
      BorderRight = i_BorderRight ;
      StyleTop    = i_StyleTop    ;
      StyleBot    = i_StyleBot    ;
      StyleLeft   = i_StyleLeft   ;
      StyleRight  = i_StyleRight  ;

      // Apply All Current Styles to Object -=-=-=-=-=-
      s_ApplyStyles( b_CellStyle ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_CellBorder ;
// --------------------------------------------------------------------------------------------------
// Set Cell Background Colour and Pattern oooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_CellFill export ;
   dcl-pi   f_CellFill ;
      b_CellStyle  object( *java : CellStyle_class ) ;
      i_Pattern    int(5)                            const ;
      i_BackRGB    char(11)                          const ;
      i_ForeRGB    char(11)                          const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      Pattern = i_Pattern  ;
      BackRGB = i_BackRGB ;
      ForeRGB = i_ForeRGB ;

      // Apply All Current Styles to Object -=-=-=-=-=-
      s_ApplyStyles( b_CellStyle ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_CellFill ;
// --------------------------------------------------------------------------------------------------
// Apply All Styles to Style Object ----------------------------------------------------------------
// --------------------------------------------------------------------------------------------------
dcl-proc s_ApplyStyles ;
   dcl-pi   s_ApplyStyles ;
      b_CellStyle   object( *java : CellStyle_class ) ;
   end-pi ;
   // Set Data Format in CellStyle object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetDataFormat extproc( *java
                              : CellStyle_class
                              : 'setDataFormat'
                              ) ;
      i_DataFormat  like(jShort) value ;
   end-pr ;
   // Associate a font object with a CellStyle object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetFontCellStyle extproc( *java
                                 : CellStyle_class
                                 : 'setFont'
                                 ) ;
      i_FontStyle  like(FontStyle) const ;
   end-pr ;
   // Set Text Alignment for a Cell jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetAlignment extproc( *java
                             : CellStyle_class
                             : 'setAlignment'
                             ) ;
      i_AlignInt  like(jShort) value ;
   end-pr ;
   // Set Text Wrapping for a Cell jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetWrapText extproc( *java
                            : CellStyle_class
                            : 'setWrapText'
                            ) ;
      i_WrapText  ind value ;
   end-pr ;
   // Set Cell Border Type (Top) jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetBorderTop extproc( *java
                             : CellStyle_class
                             : 'setBorderTop'
                             ) ;
      i_StyleInt   like(jShort) value ;
   end-pr ;
   // Set Cell Border Type (Bottom) jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetBorderBottom extproc( *java
                                : CellStyle_class
                                : 'setBorderBottom'
                                ) ;
      i_StyleInt   like(jShort) value ;
   end-pr ;
   // Set Cell Border Type (Left) jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetBorderLeft extproc( *java
                              : CellStyle_class
                              : 'setBorderLeft'
                              ) ;
      i_StyleInt   like(jShort) value ;
   end-pr ;
   // Set Cell Border Type (Right) jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetBorderRight extproc( *java
                               : CellStyle_class
                               : 'setBorderRight'
                               ) ;
      i_StyleInt   like(jShort) value ;
   end-pr ;
   // Set Cell Fill Pattern jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetFillPattern extproc( *java
                               : CellStyle_class
                               : 'setFillPattern'
                               ) ;
      i_Pattern  like(jShort) value ;
   end-pr ;
   // Set Background Colour of a Fill Pattern jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetBackColour extproc( *java
                              : XSSF_CellStyle_class
                              : 'setFillBackgroundColor'
                              ) ;
      i_Color  object( *java : 'org.apache.poi.xssf.usermodel.XSSFColor' ) const ;
   end-pr ;
   // Set Foreground Colour of a Fill Pattern jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_SetForeGroundColour extproc( *java
                                    : XSSF_CellStyle_class
                                    : 'setFillForegroundColor'
                                    ) ;
      i_Color  object( *java : 'org.apache.poi.xssf.usermodel.XSSFColor' ) const ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Data Format =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (DataFormat <> *null) ;
         j_SetDataFormat( b_CellStyle
                     : FormatInt
                     ) ;
      endif ;

      // Font -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (FontStyle <> *null) ;
         j_SetFontCellStyle( b_CellStyle
                        : FontStyle
                        ) ;
      endif ;

      // Alignment =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_SetAlignment( b_CellStyle
                 : AlignInt
                 ) ;

      // Wrap Text =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_SetWrapText( b_CellStyle
                : WrapText
                ) ;

      // Borders =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      // Top =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (NOT BorderTop) ;
         when (StyleTop <> 12) ;
            // ----------------------------------------
            j_SetBorderTop( b_CellStyle
                        : s_StyleInt( StyleTop )
                        ) ;
      endsl;

      // Bottom -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (NOT BorderBot) ;
         when (StyleBot <> 12) ;
            // ----------------------------------------
            j_SetBorderBottom( b_CellStyle
                           : s_StyleInt( StyleBot )
                           ) ;
      endsl;

      // Left -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (NOT BorderLeft) ;
         when (StyleLeft <> 12) ;
            // ----------------------------------------
            j_SetBorderLeft( b_CellStyle
                         : s_StyleInt( StyleLeft )
                         ) ;
      endsl;

      // Right =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (NOT BorderRight) ;
         when (StyleRight <> 12) ;
            // ----------------------------------------
            j_SetBorderRight( b_CellStyle
                          : s_StyleInt( StyleRight )
                          ) ;
      endsl;

      // Fill -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      // Set Fill Pattern -----------------------------
      if (Pattern > 0);
         j_SetFillPattern( b_CellStyle
                      : Pattern
                      ) ;
      endif;
      // Background Fill Colour -----------------------
      if (BackRGB <> *blank) ;
         j_SetBackColour( b_CellStyle
                     : s_XSSFColor( BackRGB )
                     ) ;
      endif;
      // Foreground Fill Colour -----------------------
      if (ForeRGB <> *blank) ;
         j_SetForeGroundColour( b_CellStyle
                           : s_XSSFColor( ForeRGB )
                           ) ;
      endif;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc s_ApplyStyles ;
// --------------------------------------------------------------------------------------------------
// Return Line Style Integer -----------------------------------------------------------------------
// --------------------------------------------------------------------------------------------------
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
dcl-proc s_StyleInt ;
   dcl-pi   s_StyleInt int(3) ;
      i_Style  int(3) const ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      select ;
         when (i_Style =  1) ;
            return 9 ;
         when (i_Style =  2) ;
            return 11 ;
         when (i_Style =  3) ;
            return 3 ;
         when (i_Style =  4) ;
            return 7 ;
         when (i_Style =  5) ;
            return 6 ;
         when (i_Style =  6) ;
            return 4 ;
         when (i_Style =  7) ;
            return 2 ;
         when (i_Style =  8) ;
            return 10 ;
         when (i_Style =  9) ;
            return 12 ;
         when (i_Style = 10) ;
            return 8 ;
         when (i_Style = 11) ;
            return 1 ;
         when (i_Style = 12) ;
            return 0 ;
      endsl;

      // Default =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      return 0 ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc s_StyleInt ;
// --------------------------------------------------------------------------------------------------
// Reset All Styles oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_ResetStyle export ;
   dcl-pi   f_ResetStyle ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Data Format =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (DataFormat <> *null) ;
         f_FreeLocalRef( DataFormat ) ;
         DataFormat = *null ;
      endif ;

      // Font -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (FontStyle <> *null) ;
         f_FreeLocalRef( FontStyle ) ;
         FontStyle = *null ;
      endif ;

      // Data Format =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      FormatInt = 0 ;

      // Alignment =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      AlignInt = 0 ;

      // Text Wrapping =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      clear WrapText ;

      // Borders =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      clear BorderTop   ;
      clear BorderBot   ;
      clear BorderLeft  ;
      clear BorderRight ;
      clear StyleTop    ;
      clear StyleBot    ;
      clear StyleLeft   ;
      clear StyleRight  ;

      // Fill -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      clear Pattern  ;
      clear BackRGB  ;
      clear ForeRGB  ;

      // Auto Width -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      clear AutoWidth ;
      clear UseMerged ;

      // -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_ResetStyle ;
// --------------------------------------------------------------------------------------------------
// Auto Size Column Width oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_AutoColWidth export ;
   dcl-pi   f_AutoColWidth ;
   end-pi ;
   // Autosize a column jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_autoSizeColumn  extproc( *java
                                : Sheet_class
                                : 'autoSizeColumn'
                                ) ;
      i_ColNumber  like( jInt ) value ;
   end-pr ;
   // Autosize a column (use merged cells to calculate) jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_autoSizeMerged  extproc( *java
                                : Sheet_class
                                : 'autoSizeColumn'
                                ) ;
      i_ColNumber  like( jInt ) value ;
      i_UseMerged  ind          value ;
   end-pr ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      if (NOT AutoWidth) ;
         return ;
      endif;

      // Autosize (not merged) =-=-=-=-=-=-=-=-=-=-=-=-
      if (NOT UseMerged) ;
         // --------------------------------------------
         j_autoSizeColumn( Sheet
                      : ColNumber
                      ) ;
         // --------------------------------------------
         return ;
         // --------------------------------------------
      endif;

      // Autosize Merged =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_autoSizeMerged( Sheet
                   : ColNumber
                   : *on
                   ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_AutoColWidth ;
// --------------------------------------------------------------------------------------------------
