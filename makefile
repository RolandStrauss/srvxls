
BIN_LIB=$(LIBRARY)
LIBRARY=LSUTILS
LIBLIST=$(LIBRARY) LSUTILSDEV BBLIB
SHELL=/QOpenSys/usr/bin/qsh
SYSTEM_PARMS=-s

all: srvxls.bnddir srvxls.srv.sqlrpgle srvxls@p.rpgleinc srvxls@t.rpgleinc srvxls@01.mod.rpgle srvxls@02.mod.rpgle srvxls@10.mod.rpgle srvxls@11.mod.rpgle srvxls@01.mod.rpgle srvxls@12.mod.rpgle srvxls@13.mod.rpgle srvxls@14.mod.rpgle srvxls@20.mod.rpgle  srvxls@aa.mod.rpgle srvxls@au.mod.rpgle srvxls@sh.mod.rpgle  srvxls@sp.mod.rpgle

## Targets

srvxls.srv.sqlrpgle: srvxls.bnd srvxls@01.mod.rpgle srvxls@02.mod.rpgle srvxls@10.mod.rpgle srvxls@11.mod.rpgle srvxls@01.mod.rpgle srvxls@12.mod.rpgle srvxls@13.mod.rpgle srvxls@14.mod.rpgle srvxls@20.mod.rpgle  srvxls@aa.mod.rpgle srvxls@au.mod.rpgle srvxls@sh.mod.rpgle  srvxls@sp.mod.rpgle

## Rules
   ## system "CRTRPGMOD MODULE($(BIN_LIB)/$*) SRCSTMF('/home/ROLAND/builds/LSPS-1162/QRPGLESRC/$*.pgm.rpgle') DBGVIEW(*SOURCE) OPTION(*EVENTF)"

%.rpgleinc: qrpglesrc/%.rpgleinc
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	@touch $@

%.rpgleinc: qrpglesrc/%.rpgleinc
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	@touch $@
	
%.pgm.rpgle: qrpglesrc/%.pgm.rpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	liblist -a $(LIBLIST);\
	system "CRTRPGMOD MODULE($(BIN_LIB)/$*) SRCSTMF('$<') DBGVIEW(*SOURCE) OPTION(*EVENTF)"
	system "CRTPGM PGM($(BIN_LIB)/$*) MODULE($(BIN_LIB)/$*) ACTGRP(*NEW)"
	@touch $@

%.mod.rpgle: qrpglesrc/%.mod.rpgle
	system $(SYSTEM_PARMS) "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	liblist -a $(LIBLIST);\
	system $(SYSTEM_PARMS) "CRTRPGMOD MODULE($(BIN_LIB)/$*) SRCSTMF('$<') DBGVIEW(*SOURCE) OPTION(*EVENTF)"

%.pgm.sqlrpgle: qrpglesrc/%.pgm.sqlrpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	liblist -a $(LIBLIST);\
	system "CRTSQLRPGI OBJ($(BIN_LIB)/$*) SRCSTMF('$<') OBJTYPE(*MODULE) OPTION(*EVENTF) RPGPPOPT(*LVL2) DBGVIEW(*SOURCE)"
	system "CRTPGM PGM($(BIN_LIB)/$*) MODULE($(BIN_LIB)/$*) ACTGRP(*NEW)"
	@touch $@	

%.mod.sqlrpgle: qrpglesrc/%.mod.sqlrpgle
	system $(SYSTEM_PARMS) "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	liblist -a $(LIBLIST);\
	system $(SYSTEM_PARMS) "CRTSQLRPGI OBJ($(BIN_LIB)/$*) SRCSTMF('$<') CLOSQLCSR(*ENDMOD) OPTION(*EVENTF) DBGVIEW(*SOURCE) OBJTYPE(*MODULE) RPGPPOPT(*LVL2) "

%.srv.rpgle: qrpglesrc/%.srv.rpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	liblist -a $(LIBLIST);\
	system "CRTRPGMOD MODULE($(BIN_LIB)/$*) SRCSTMF('$<') DBGVIEW(*SOURCE) OPTION(*EVENTF)"
	system "CRTSRVPGM SRVPGM($(BIN_LIB)/$*) MODULE($(BIN_LIB)/$*) EXPORT(*SRCFILE) SRCSTMF('$<')"
	system "DLTOBJ OBJ($(BIN_LIB)/$*) OBJTYPE(*MODULE)"
	@touch $@
	
%.srv.sqlrpgle: qrpglesrc/%.srv.sqlrpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	liblist -a $(LIBLIST);\
	system "CRTSQLRPGI OBJ($(BIN_LIB)/$*) SRCSTMF('$<') OBJTYPE(*MODULE) OPTION(*EVENTF) RPGPPOPT(*LVL2) DBGVIEW(*SOURCE)"
	system "CRTSRVPGM SRVPGM($(BIN_LIB)/$*) MODULE($(BIN_LIB)/$*) EXPORT(*SRCFILE) SRCSTMF('$<')"
	system "DLTOBJ OBJ($(BIN_LIB)/$*) OBJTYPE(*MODULE)"
	@touch $@

%.dspf: qddssrc/%.dspf
	-system -qi "CRTSRCPF FILE($(BIN_LIB)/QDDSSRC) RCDLEN(112)"
	system "CPYFRMSTMF FROMSTMF('./qddssrc/$*.dspf') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QDDSSRC.file/$*.mbr') MBROPT(*REPLACE)"
	system -s "CRTDSPF FILE($(BIN_LIB)/$*) SRCFILE($(BIN_LIB)/QDDSSRC) SRCMBR($*)"
	@touch $@

%.sqltabl: qsqlsrc/%.sqltabl
	liblist -c $(BIN_LIB);\
	system "RUNSQLSTM SRCSTMF('$<') COMMIT(*NONE)"
	@touch $@

%.rpgleinc: qrpgleref/%.rpgleinc
	@touch $@

%.bnd: qsrvsrc/%.bnd
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	@touch $@

%.cmd: qcmdsrc/%.cmd
	-system -q "CRTSRCPF FILE($(BIN_LIB)/QCMDSRC) RCDLEN(112)"
	system $(SYSTEM_PARMS) "CPYFRMSTMF FROMSTMF('$<') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QCMDSRC.file/$*.mbr') MBROPT(*REPLACE)"
	system $(SYSTEM_PARMS) "CRTCMD CMD($(BIN_LIB)/$*) PGM($(BIN_LIB)/$*) SRCFILE($(BIN_LIB)/QCMDSRC)"

%.bnddir:
	-system -q "CRTBNDDIR BNDDIR($(BIN_LIB)/$*)"
	-system -q "ADDBNDDIRE BNDDIR($(BIN_LIB)/$*) OBJ($(patsubst %.srvpgm,($(BIN_LIB)/% *SRVPGM *DEFER),$^))"
