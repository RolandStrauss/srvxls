**free
// --------------------------------------------------------------------------------------------------
// Program File:
//   Name : SRVXLS@20 : Standard Headers
//   Type : *MODULE
// Description :
//   Add header and footer (totals) to sheet.
// --------------------------------------------------------------------------------------------------
//   Pre-Compiler tags used by STRPREPRC to retrieve creation
//   commands from the source member.
// --------------------------------------------------------------------------------------------------
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
// 2019-08-12  Nico Basson       Initial Code.
// --------------------------------------------------------------------------------------------------
/INCLUDE SRVSRC,SRV@H
ctl-opt nomain
        thread(*concurrent)
        stgmdl(*inherit) ;
// Framework ---------------------------------------------------------------------------------------
/DEFINE  SRVXLS
/INCLUDE SRVSRC,SRV@P
// --------------------------------------------------------------------------------------------------
// Add Standard Outsource Company (UMA) Logo to Sheet <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
// --------------------------------------------------------------------------------------------------
dcl-proc xls_StdHdrLogo export ;
   dcl-pi   xls_StdHdrLogo ;
      i_Outsource  packed(3) const ;
   end-pi ;
   // Outsource Company (UMA) -------------------------------------------------------------------------
   dcl-ds UMA_Logo qualified ;
      Name   varchar(128) inz('_Excel_Header_Full.png') ;
      Folder varchar(128) inz('/BusinessDocuments/Images/Logos/UMA/') ;
   end-ds ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      reset UMA_Logo ;

      // Set Logo File Name -=-=-=-=-=-=-=-=-=-=-=-=-=-
      UMA_Logo.Name = 'UMA' + %editc( i_Outsource : 'X' ) + UMA_Logo.Name ;
      // Not Found ------------------------------------
      if (NOT ifs_FileExists( %trim( UMA_Logo.Folder + UMA_Logo.Name ) )) ;
         return ;
      endif;

      // UMA Logo -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      xls_AddRow( 0 ) ;
      xls_AddImage( UMA_Logo.Name
               : UMA_Logo.Folder
               : 0
               : 0
               : *omit
               : 150
               ) ;
      // -----------------------------------------------
      return ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc xls_StdHdrLogo ;
// --------------------------------------------------------------------------------------------------
