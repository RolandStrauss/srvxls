**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@12 : Microsoft Excel - Pictures.
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
// 2019-01-09  Nico Basson       Initial Code.
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
// Sheet -------------------------------------------------------------------------------------------
dcl-s Sheet         object( *java : Sheet_class )                                           import ;
// --------------------------------------------------------------------------------------------------
// Add Picture to Sheet <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_AddImage export ;
   dcl-pi   xls_AddImage ;
      i_Picture     varchar(512) const ;
      i_Directory   varchar(512) const ;
      i_Row         int(10)      const ;
      i_Col         int(10)      const ;
      i_Width       int(10)      const options( *omit : *nopass ) ;
      i_Height      int(10)      const options( *omit : *nopass ) ;
      i_KeepAspect  ind          const options( *omit : *nopass ) ;
   end-pi ;
   // --------------------------------------------------
   dcl-s Directory   varchar(512)  ;
   dcl-s PicPath     varchar(1024) ;
   dcl-s Width       float(8)      ;
   dcl-s Height      float(8)      ;
   dcl-s KeepAspect  ind           ;
   // --------------------------------------------------
   dcl-s ErrorData char(128)  ;
   // Java Class Constants jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-c FILE_CLASS  'java.io.File' ;
   // Create a drawing patriarch to draw pictures on a sheet jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_createDrawingPatriarch  object (*java : Drawing_class )
                                 extproc( *java
                                        : Sheet_class
                                        : 'createDrawingPatriarch'
                                        ) ;
   end-pr ;
   // --------------------------------------------------
   dcl-s Drawing  object( *java : Drawing_class ) ;
   // Create a Java File Object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_File  object (*java : File_class )
               extproc( *java
                      : File_class
                      : *constructor
                      ) ;
      i_filePath  like(jString) const ;
   end-pr ;
   // --------------------------------------------------
   dcl-s Path  like(jString) ;
   dcl-s File  object( *java : File_class ) ;
   // Create a Java URL Object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_URL  object (*java : URL_class )
              extproc( *java
                     : URL_class
                     : *constructor
                     ) ;
      i_filePath  like(jString) const ;
   end-pr ;
   // --------------------------------------------------
   dcl-s URL   object( *java : URL_class ) ;
   // Create a Java File Object jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_AddDimensionedImageClass object (*java : DimImage_class )
                                  extproc( *java
                                         : DimImage_class
                                         : *constructor
                                         ) ;
   end-pr ;
   // --------------------------------------------------
   dcl-s AddDimensionedImageClass  object (*java : DimImage_class ) ;
   // Add Image to a Sheet jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_addImageToSheet extproc( *java
                                : DimImage_class
                                : 'addImageToSheet'
                                ) ;
      i_Col           like(jInt)    value ;
      i_Row           like(jInt)    value ;
      i_Sheet         like(Sheet)   ;
      i_Drawing       like(Drawing) ;
      i_imageFile     like(URL)     ;
      i_Width         like(jDouble) value ;
      i_Height        like(jDouble) value ;
      i_Resize        like(jInt)    value ;
   end-pr ;
   // Pixels to Millimetres Factor --------------------------------------------------------------------
   dcl-c PX_MM_FACTOR  0,2646135265700483 ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Picture Path -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      Directory = gen_DirPathName( i_Directory ) ;
      PicPath = Directory + i_Picture ;
      // Check Exist ----------------------------------
      if (NOT ifs_FileExists( %trim( PicPath ) )) ;
         ErrorData = PicPath ;
         err_SendEsc( 'IFS0002'
                 : 'SRVERR'
                 : ErrorData
                 ) ;
      endif;

      // Create Drawing Patriarch -=-=-=-=-=-=-=-=-=-=-
      Drawing = j_createDrawingPatriarch( Sheet ) ;

      // URL from Picture Path =-=-=-=-=-=-=-=-=-=-=-=-
      PicPath = 'file:///' + PicPath ;
      // -----------------------------------------------
      Path = java_NewString( PicPath ) ;
      File = j_File( Path ) ;
      URL  = j_URL( Path ) ;
      // -----------------------------------------------
      f_FreeLocalRef( Path ) ;
      f_FreeLocalRef( File ) ;

      // Image Width and Height -=-=-=-=-=-=-=-=-=-=-=-
      clear Width ;
      select ;
         when (%parms < %parmnum( i_Width )) ;
         when (%addr( i_Width ) <> *null) ;
            // ----------------------------------------
            Width = i_Width ;
      endsl;
      clear Height ;
      select ;
         when (%parms < %parmnum( i_Height )) ;
         when (%addr( i_Height ) <> *null) ;
            // ----------------------------------------
            Height = i_Height ;
      endsl;
      // Keep Aspect Ratio ----------------------------
      KeepAspect = *on ;
      select ;
         when (%parms < %parmnum( i_KeepAspect )) ;
         when (%addr( i_KeepAspect ) <> *null) ;
            // ----------------------------------------
            KeepAspect = i_KeepAspect ;
      endsl;
      // Set Image Dimension --------------------------
      s_SetImageDimension( URL
                      : Width
                      : Height
                      : KeepAspect
                      ) ;

      // Instantiate Class =-=-=-=-=-=-=-=-=-=-=-=-=-=-
      AddDimensionedImageClass = j_AddDimensionedImageClass() ;

      // Add Image to Sheet -=-=-=-=-=-=-=-=-=-=-=-=-=-
      j_addImageToSheet( AddDimensionedImageClass
                    : i_Col
                    : i_Row
                    : Sheet
                    : Drawing
                    : URL
                    // Width in mm
                    : Width  * px_mm_Factor
                    // Height in mm
                    : Height * px_mm_Factor
                    // resizeBehaviour
                    //    EXPAND_ROW = 1;
                    //    EXPAND_COLUMN = 2;
                    //    EXPAND_ROW_AND_COLUMN = 3;
                    //    OVERLAY_ROW_AND_COLUMN = 7;
                    : 7
                    ) ;

      // Cleanup =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      f_FreeLocalRef( AddDimensionedImageClass ) ;
      f_FreeLocalRef( Drawing ) ;
      f_FreeLocalRef( URL ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      f_CleanupJAVA() ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_AddImage ;
// --------------------------------------------------------------------------------------------------
// Set Image Dimension -----------------------------------------------------------------------------
// --------------------------------------------------------------------------------------------------
dcl-proc s_SetImageDimension ;
   dcl-pi   s_SetImageDimension ;
      i_URL         object( *java : URL_class ) ;
      b_Width       float(8)                    ;
      b_Height      float(8)                    ;
      i_KeepAspect  ind                         const ;
   end-pi ;
   // Java Class Constants jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-c BUFFIMAGE_CLASS 'java.awt.image.BufferedImage' ;
   dcl-c IMAGEIO_CLASS   'javax.imageio.ImageIO' ;
   // Image IO Read jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_imageIO_Read object (*java : BuffImage_Class )
                      extproc( *java
                             : ImageIO_Class
                             : 'read'
                             )
                      static ;
      i_URL  object( *java : URL_class ) ;
   end-pr ;
   // --------------------------------------------------
   dcl-s BuffImage  object (*java : BuffImage_Class ) ;
   // Get Image Height jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_Height  like(jInt)
                 extproc( *java
                        : BuffImage_Class
                        : 'getHeight'
                        ) ;
   end-pr ;
   // --------------------------------------------------
   dcl-s Height    int(10)       ;
   // Get Image Width jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj
   dcl-pr j_Width  like(jInt)
                extproc( *java
                       : BuffImage_Class
                       : 'getWidth'
                       ) ;
   end-pr ;
   // --------------------------------------------------
   dcl-s Width  int(10) ;
   // Keep Aspect with Width and Height Specified -----------------------------------------------------
   dcl-ds AspectRatio qualified ;
      Picture   float(8) ;
      Required  float(8) ;
   end-ds ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Height and Width Specified (Do not Keep Aspect)
      if (b_Width <> 0 AND b_Height <> 0
      AND NOT i_KeepAspect) ;
         // --------------------------------------------
         return ;
      endif;

      // Get Image Dimension Ratio =-=-=-=-=-=-=-=-=-=-
      BuffImage = *null ;
      BuffImage = j_imageIO_Read( i_URL ) ;
      // -----------------------------------------------
      Height = j_Height( BuffImage ) ;
      Width  = j_Width ( BuffImage ) ;
      // -----------------------------------------------
      f_FreeLocalRef( BuffImage ) ;

      // Aspect Ratios =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      AspectRatio.Picture  =   Width /   Height ;
      AspectRatio.Required = b_Width / b_Height ;

      // Original Size =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (b_Width = 0 AND b_Height = 0) ;
         // --------------------------------------------
         b_Width  = Width ;
         b_Height = Height ;
         return ;
         // --------------------------------------------
      endif;

      // Proportional Width -=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (b_Width = 0 AND b_Height <> 0) ;
         // --------------------------------------------
         b_Width = b_Height * AspectRatio.Picture ;
         return ;
         // --------------------------------------------
      endif;

      // Proportional Height =-=-=-=-=-=-=-=-=-=-=-=-=-
      if (b_Width <> 0 AND b_Height = 0) ;
         // --------------------------------------------
         b_Height = b_Width / AspectRatio.Picture ;
         return ;
         // --------------------------------------------
      endif;

      // Fit - Keep Aspect =-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
            // Picture Portrait - Required Landscape ------
         when (AspectRatio.Picture < 1
          AND AspectRatio.Required > 1) ;
            // Scale to Height -----------------------
            b_Width = b_Height * AspectRatio.Picture ;
            return ;
            // Picture Landscape - Required Portrait ------
         when (AspectRatio.Picture > 1
          AND AspectRatio.Required < 1) ;
            // Scale to Width ------------------------
            b_Height = b_Width / AspectRatio.Picture ;
            return ;
      endsl;
      // Both Portrait or Landscape -------------------
      select ;
         when (AspectRatio.Picture < AspectRatio.Required) ;
            // Scale to Height -----------------------
            b_Width = b_Height * AspectRatio.Picture ;
         other ;
            // Scale to Width ------------------------
            b_Height = b_Width / AspectRatio.Picture ;
      endsl;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc s_SetImageDimension ;
// --------------------------------------------------------------------------------------------------
