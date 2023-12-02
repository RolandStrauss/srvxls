
BIN_LIB=CMPSYS
LIBLIST=$(BIN_LIB) COBRA COBRAS BBLIB
SHELL=/QOpenSys/usr/bin/qsh

all: wmssrv.srv.sqlrpgle wms_db_bld.sqlrpgle wmsdriver.pgm.sqlrpgle wms001a.pgm.sqlrpgle wms001b.pgm.sqlrpgle wms001c.pgm.sqlrpgle wms002a.pgm.sqlrpgle wms002b.pgm.sqlrpgle wms010.pgm.sqlrpgle wms011.pgm.sqlrpgle wms011a.pgm.sqlrpgle wms012.pgm.sqlrpgle wms013.pgm.sqlrpgle wms014.pgm.sqlrpgle wms015.pgm.sqlrpgle wms016.pgm.sqlrpgle wms020a.pgm.sqlrpgle wms020b.pgm.sqlrpgle wms020c.pgm.sqlrpgle wms020d.pgm.sqlrpgle wms021.pgm.sqlrpgle wms030.pgm.sqlrpgle

## Targets

wmssrv.srv.sqlrpgle: wmssrv.bnd wmssrv.cb.rpgle wmssrv.pr.rpgle

wms001a.pgm.sqlrpgle: wms001a.dspf
wms001b.pgm.sqlrpgle: wms001b.dspf
wms001c.pgm.sqlrpgle: wms001c.dspf

wms002a.pgm.sqlrpgle: wms002a.dspf
wms002b.pgm.sqlrpgle: wms002b.dspf

wms010.pgm.sqlrpgle: wms010.dspf

wms011.pgm.sqlrpgle: wms011.dspf
wms011a.pgm.sqlrpgle: wms011a.dspf

wms012.pgm.sqlrpgle: wms012.dspf

wms013.pgm.sqlrpgle: wms013.dspf

wms014.pgm.sqlrpgle: wms014.dspf

wms015.pgm.sqlrpgle: wms015.dspf

wms016.pgm.sqlrpgle: wms016.dspf

wms020a.pgm.sqlrpgle: wms020a.dspf
wms020b.pgm.sqlrpgle: wms020b.dspf
wms020c.pgm.sqlrpgle: wms020c.dspf
wms020d.pgm.sqlrpgle: wms020d.dspf
wms021.pgm.sqlrpgle: wms021.dspf

wms030.pgm.sqlrpgle: wms030.dspf

## Rules
   ## system "CRTRPGMOD MODULE($(BIN_LIB)/$*) SRCSTMF('/home/ROLAND/builds/LSPS-1162/QRPGLESRC/$*.pgm.rpgle') DBGVIEW(*SOURCE) OPTION(*EVENTF)"

%.cb.rpgle: qrpglesrc/%.cb.rpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	@touch $@

%.pr.rpgle: qrpglesrc/%.pr.rpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	@touch $@
	
%.pgm.rpgle: qrpglesrc/%.pgm.rpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	system "CPYFRMSTMF FROMSTMF('./qrpglesrc/$*.rpgle') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QRPGLESRC.file/$*.mbr') MBROPT(*REPLACE)"
	SDMIM/ISETLIBL ENV(PDSCOBDEV) CMD(SDMIM/ICOMPPDM MBROBJ($*) LIBRARY($(BIN_LIB)) SRCF(QRPGLESRC) SUBMIT(*NO)) LIB1(BBLIB) 
	@touch $@
	
%.pgm.sqlrpgle: qrpglesrc/%.pgm.sqlrpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	system "CPYFRMSTMF FROMSTMF('./qrpglesrc/$*.sqlrpgle') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QRPGLESRC.file/$*.mbr') MBROPT(*REPLACE)"
	SDMIM/ISETLIBL ENV(PDSCOBDEV) CMD(SDMIM/ICOMPPDM MBROBJ($*) LIBRARY($(BIN_LIB)) SRCF(QRPGLESRC) SUBMIT(*NO)) LIB1(BBLIB) 
	@touch $@	

%.srv.rpgle: qrpglesrc/%.srv.rpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	system "CPYFRMSTMF FROMSTMF('./qrpglesrc/$*.rpgle') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QRPGLESRC.file/$*.mbr') MBROPT(*REPLACE)"
	SDMIM/ISETLIBL ENV(PDSCOBDEV) CMD(SDMIM/ICOMPPDM MBROBJ($*) LIBRARY($(BIN_LIB)) SRCF(QRPGLESRC) SUBMIT(*NO)) LIB1(BBLIB) 
	@touch $@
	
%.srv.sqlrpgle: qrpglesrc/%.srv.sqlrpgle
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	system "CPYFRMSTMF FROMSTMF('./qrpglesrc/$*.sqlrpgle') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QRPGLESRC.file/$*.mbr') MBROPT(*REPLACE)"
	SDMIM/ISETLIBL ENV(PDSCOBDEV) CMD(SDMIM/ICOMPPDM MBROBJ($*) LIBRARY($(BIN_LIB)) SRCF(QRPGLESRC) SUBMIT(*NO)) LIB1(BBLIB)
	@touch $@

%.dspf: qddssrc/%.dspf
	-system -qi "CRTSRCPF FILE($(BIN_LIB)/QDDSSRC) RCDLEN(112)"
	system "CPYFRMSTMF FROMSTMF('./qddssrc/$*.dspf') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QDDSSRC.file/$*.mbr') MBROPT(*REPLACE)"
	system -s "CRTDSPF FILE($(BIN_LIB)/$*) SRCFILE($(BIN_LIB)/QDDSSRC) SRCMBR($*)"
	@touch $@

%.sqltabl: qsqlsrc/%.sqltabl
	liblist -c $(BIN_LIB);\
	system "CPYFRMSTMF FROMSTMF('./qddssrc/$*.rpgleinc') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QSQLSRC.file/$*.mbr') MBROPT(*REPLACE)"
	SDMIM/ISETLIBL ENV(PDSCOBDEV) CMD(SDMIM/ICOMPPDM MBROBJ($*) LIBRARY($(BIN_LIB)) SRCF(QSQLSRC) SUBMIT(*NO)) LIB1(BBLIB)
	@touch $@

%.sql: qsqlsrc/%.sql
	liblist -c $(BIN_LIB);\
	system "RUNSQLSTM SRCSTMF('$<') COMMIT(*NONE)"
	@touch $@

%.table: qsqlsrc/%.table
	liblist -c $(BIN_LIB);\
	system "RUNSQLSTM SRCSTMF('$<') COMMIT(*NONE)"
	@touch $@

%.rpgleinc: qrpgleref/%.rpgleinc
	system "CPYFRMSTMF FROMSTMF('./qddssrc/$*.rpgleinc') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QRPGLEREF.file/$*.mbr') MBROPT(*REPLACE)"
	@touch $@

%.bnd: qsrvsrc/%.bnd
	system -s "CHGATR OBJ('$<') ATR(*CCSID) VALUE(1252)"
	system "CPYFRMSTMF FROMSTMF('./qddssrc/$*.bnd') TOMBR('/QSYS.lib/$(BIN_LIB).lib/QSRVSRC.file/$*.mbr') MBROPT(*REPLACE)"
	@touch $@
