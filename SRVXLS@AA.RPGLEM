**free
// --------------------------------------------------------------------------------------------------
// Program Object:
//   Name : SRVXLS@AA : Microsoft Excel Parsing - Static Storage Handler
//   Type : *MODULE
// Description :
//   These procedures provide static storage to callback procedures for stateless variable retrieval.
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
// Internal Function Prototypes oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
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
// Sheet Name handler oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_SheetName export ;
   dcl-pi   f_SheetName varchar(1024) ;
      i_Name    varchar(1024) const options( *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s SheetName varchar(1024) static ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Set Name -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (%parms >= %parmnum( i_Name )) ;
         SheetName = i_Name ;
         return *blank ;
      endif;

      // Provide Name -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      return SheetName ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_SheetName ;
// --------------------------------------------------------------------------------------------------
// Sheet Found handler ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_SheetFound export ;
   dcl-pi   f_SheetFound ind ;
      i_Found    ind const options( *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-s Found ind static ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Set =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      if (%parms >= %parmnum( i_Found )) ;
         Found = i_Found ;
         return Found ;
      endif;

      // Provide Name -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      return Found ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_SheetFound ;
// --------------------------------------------------------------------------------------------------
// Static storage procedure pointer handler oooooooooooooooooooooooooooooooooooooooooooooooooooooooo
// --------------------------------------------------------------------------------------------------
dcl-proc f_Proc@ export ;
   dcl-pi   f_Proc@ pointer(*proc) ;
      i_Type  varchar(15)    const ;
      i_@     pointer(*proc) const options( *nopass ) ;
   end-pi ;
   // --------------------------------------------------------------------------------------------------
   dcl-ds Proc@ qualified static ;
      Cell     pointer(*proc) ;
      NewRow   pointer(*proc) ;
      EndRow   pointer(*proc) ;
   end-ds ;
   // --------------------------------------------------------------------------------------------------
   monitor ;
      // ===========================================================
      // Set Pointer =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (%parms < %parmnum( i_@ )) ;
         when (i_Type = 'Cell') ;
            Proc@.Cell = i_@ ;
            return *null ;
         when (i_Type = 'NewRow') ;
            Proc@.NewRow = i_@ ;
            return *null ;
         when (i_Type = 'EndRow') ;
            Proc@.EndRow = i_@ ;
            return *null ;
      endsl;

      // Return Pointer -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
      select ;
         when (i_Type = 'Cell') ;
            return Proc@.Cell ;
         when (i_Type = 'NewRow') ;
            return Proc@.NewRow ;
         when (i_Type = 'EndRow') ;
            return Proc@.EndRow ;
      endsl;
      // -----------------------------------------------
      return *null ;

      // ===========================================================
   on-error ;
      err_BubbleMsg() ;
   endmon ;
   // --------------------------------------------------------------------------------------------------
end-proc f_Proc@ ;
// --------------------------------------------------------------------------------------------------
