const { PT_TOKEN_EXCEL: TOKEN_EXCEL, PT_TOKEN_EXCEL_MAP: TOKEN_EXCEL_MAP, FT_CODE_FUNCTION_EXCEL: CODE_FUNCTION_EXCEL, RECORD_OP_CODES: OP_CODES, BUILT_IN_LABEL_NAMES, OP_CODES_MAP, STREAM_TYPES, get_sheet_type_descriptive, get_sheet_state_descriptive } = require('./biff-constants')
const { process_RGCE, structure_RGCE, RGCE_to_javasript, processXLUnicodeStringNoCch, processXLUnicodeString } = require('./biff-formulas')

const ieee754 = require('./ieee754')

const RECORD_CONSTRUCTOR = {
    0x06: (data, context) => { //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/8e3c6978-6c9f-4915-a826-07613204b244
        let row = data.readUInt16LE(0)
        let col = data.readUInt16LE(2)
        let ixfe = data.readUInt16LE(4)
        let formula_value = parseFormulaValue(data.slice(6, 14))
        let fShrFmla = data.readUInt8(14) & 0x10
        // https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/7dd67f0a-671d-4905-b87b-4cc07295e442
        // CellParsedFormula
        // 20
        let cce = data.readUInt16LE(20)
        let rgce = data.slice(22, 22 + cce)
        let rgcb = data.slice(22 + cce)
        let cell = {
            rw: row,
            col: col,
            value: {}
        }
        let exp = process_RGCE(rgce, { row, col, rgcb, sheet: context.current_sheet.name })
        let rev = exp.reverse()
        try {
            cell.value.exp = exp
            cell.value.exp_str = structure_RGCE(rev),
                cell.value.js = RGCE_to_javasript(rev)
            if (typeof cell.value.js === 'object') {
                cell.value.js = cell.value.js.value
            }
            cell.js_v = cell.value.js
        } catch (err) {
            cell.value.err = err.toString()
        }
        context.add_cell(cell)
        return context
    },//FORMULA : Cell Formula,
    0x0A: (data, context) => { return context },//EOF : End of File,
    0x0C: (data, context) => { return context },//CALCCOUNT : Iteration Count,
    0x0D: (data, context) => { return context },//CALCMODE : Calculation Mode,
    0x0E: (data, context) => { return context },//PRECISION : Precision,
    0x0F: (data, context) => { return context },//REFMODE : Reference Mode,
    0x10: (data, context) => { return context },//DELTA : Iteration Increment,
    0x11: (data, context) => { return context },//ITERATION : Iteration Mode,
    0x12: (data, context) => {
        let fLock = data.readUInt16LE(0);
        return context
    },//PROTECT : Protection Flag,
    0x13: (data, context) => {
        let wPassword = data.readUInt16LE(0);
        return context
    },//PASSWORD : Protection Password,
    0x14: (data, context) => { return context },//HEADER : Print Header on Each Page,
    0x15: (data, context) => { return context },//FOOTER : Print Footer on Each Page,
    0x16: (data, context) => { return context },//EXTERNCOUNT : Number of External References,
    0x17: (data, context) => {
        //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/475df8d4-a3be-47a5-9a01-4ec828059f43
        let n_entries = data.readUInt16LE()
        let array = data.slice(2)
        try {
            for (let i = 0; i < n_entries; i++) {
                //Overcomplicated
                //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/5adbad90-093d-4bc6-acc1-b662270bc0d7
                //XTI: [iSupBook,itabFirst, itabLast]
                let iSupBook = array.readUInt16LE(i * 6)
                let itabFirst = array.readUInt16LE((i * 6) + 2)
                let itabLast = array.readUInt16LE((i * 6) + 4)
                context.external_sheets.push({
                    iSupBook,
                    itabFirst,
                    itabLast
                })
            }
        } catch (err) { }

        return context
    },//EXTERNSHEET : External Reference,
    0x18: (data, context) => {
        //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/d148e898-4504-4841-a793-ee85f3ea9eef
        let flags = data.readUInt16LE(0)
        let chkey = data.readUInt8(2)
        let cch = data.readUInt8(3)
        let cce = data.readUInt16LE(4)
        let itab = data.readUInt16LE(8)
        let fBuiltin = flags & 0x20
        let name;
        let ofst = 14;
        if (data.readUInt8(ofst) === 0) {
            ofst = 15
        }
        if (fBuiltin === 0) {
            name = data.slice(ofst, ofst + cch)
        } else {
            name = BUILT_IN_LABEL_NAMES[data.readUInt8(ofst)]

        }

        //If Auto_Open label... then its malware...

        let rgce = data.slice(ofst + 1, ofst + 1 + cce)
        let rgcb = data.slice(ofst + 1 + cce)
        let named_formula = null
        let label = {
            name: name.toString()
        }
        try {
            label.named_formula = process_RGCE(rgce, { row: 0, col: 0, rgcb, named_formula: true, sheet: context.current_sheet.name });
        } catch (err) {
            label.err = err.toString()
        }
        context.current_sheet.labels.push(label)
        return context
    },//LABEL : Cell Value, String Constant,
    0x19: (data, context) => { return context },//WINDOWPROTECT : Windows Are Protected,
    0x1A: (data, context) => { return context },//VERTICALPAGEBREAKS : Explicit Column Page Breaks,
    0x1B: (data, context) => { return context },//HORIZONTALPAGEBREAKS : Explicit Row Page Breaks,
    0x1C: (data, context) => { return context },//NOTE : Comment Associated with a Cell,
    0x1D: (data, context) => { return context },//SELECTION : Current Selection,
    0x22: (data, context) => { return context },//1904 : 1904 Date System,
    0x26: (data, context) => { return context },//LEFTMARGIN : Left Margin Measurement,
    0x27: (data, context) => { return context },//RIGHTMARGIN : Right Margin Measurement,
    0x28: (data, context) => { return context },//TOPMARGIN : Top Margin Measurement,
    0x29: (data, context) => { return context },//BOTTOMMARGIN : Bottom Margin Measurement,
    0x2A: (data, context) => { return context },//PRINTHEADERS : Print Row/Column Labels,
    0x2B: (data, context) => { return context },//PRINTGRIDLINES : Print Gridlines Flag,
    0x2F: (data, context) => { return context },//FILEPASS : File Is Password-Protected,
    0x31: (data, context) => { return context },//FONT,
    0x32: (data, context) => { return context },//FONT2,
    0x3C: (data, context) => {
        return RECORD_CONSTRUCTOR[context.last_opcode](data, context)
    },//CONTINUE : Continues Long Records,
    0x3D: (data, context) => { return context },//WINDOW1 : Window Information,
    0x40: (data, context) => { return context },//BACKUP : Save Backup Version of the File,
    0x41: (data, context) => { return context },//PANE : Number of Panes and Their Position,
    0x42: (data, context) => { return context },//CODENAME : VBE Object Name,
    0x42: (data, context) => { return context },//CODEPAGE : Default Code Page,
    0x4D: (data, context) => { return context },//PLS : Environment-Specific Print Record,
    0x50: (data, context) => { return context },//DCON : Data Consolidation Information,
    0x51: (data, context) => { return context },//DCONREF : Data Consolidation References,
    0x52: (data, context) => { return context },//DCONNAME : Data Consolidation Named References,
    0x55: (data, context) => { return context },//DEFCOLWIDTH : Default Width for Columns,
    0x59: (data, context) => { return context },//XCT : CRN Record Count,
    0x5A: (data, context) => { return context },//CRN : Nonresident Operands,
    0x5B: (data, context) => { return context },//FILESHARING : File-Sharing Information,
    0x5C: (data, context) => {
        //Info about the user
        let size = data.readUInt16LE(0)
        let unicode = data[2] & 0x01
        context.user_name = data.slice(3, 3 + size).toString(unicode ? 'utf16le' : 'utf8')
        return context
    },//WRITEACCESS : Write Access User Name,
    0x5D: (data, context) => { return context },//OBJ : Describes a Graphic Object,
    0x5E: (data, context) => { return context },//UNCALCED : Recalculation Status,
    0x5F: (data, context) => { return context },//SAVERECALC : Recalculate Before Save,
    0x60: (data, context) => { return context },//TEMPLATE : Workbook Is a Template,
    0x63: (data, context) => { return context },//OBJPROTECT : Objects Are Protected,
    0x7D: (data, context) => { return context },//COLINFO : Column Formatting Information,
    0x7E: (data, context) => { return context },//RK : Cell Value, RK Number,
    0x7F: (data, context) => { return context },//IMDATA : Image Data,
    0x80: (data, context) => { return context },//GUTS : Size of Row and Column Gutters,
    0x81: (data, context) => {
        let fDialog = (data.readUInt8(0) & 16) > 0;
        context.current_sheet.type_desc = fDialog ? "Dialog sheet" : "Work sheet"
        return context
    },//WSBOOL : Additional Workspace Information,
    0x82: (data, context) => { return context },//GRIDSET : State Change of Gridlines Option,
    0x83: (data, context) => { return context },//HCENTER : Center Between Horizontal Margins,
    0x84: (data, context) => { return context },//VCENTER : Center Between Vertical Margins,
    0x85: (data, context) => {
        let positionBOF = data.readInt32LE(0)
        let sheetState = data.readUInt8(4)
        let sheetType = data.readUInt8(5)

        let size = data[6]
        let sheetName = data.slice(8, 8 + size).toString((data[7] & 0x01) ? 'utf16le' : 'utf8')
        let sheet = {
            name: sheetName,
            type: sheetType,
            type_desc: get_sheet_type_descriptive(sheetType),
            state: sheetState,
            state_desc: get_sheet_state_descriptive(sheetState),
            pos: positionBOF,
            cells: [],
            labels: []
        }
        if (context.current_sheet == null) {
            context.current_sheet = sheet
        }
        context.add_sheet(sheet)
        return context
    },//BOUNDSHEET : Sheet Information,
    0x86: (data, context) => { return context },//WRITEPROT : Workbook Is Write-Protected,
    0x87: (data, context) => { return context },//ADDIN : Workbook Is an Add-in Macro,
    0x88: (data, context) => { return context },//EDG : Edition Globals,
    0x89: (data, context) => { return context },//PUB : Publisher,
    0x8C: (data, context) => { return context },//COUNTRY : Default Country and WIN.INI Country,
    0x8D: (data, context) => { return context },//HIDEOBJ : Object Display Options,
    0x90: (data, context) => { return context },//SORT : Sorting Options,
    0x91: (data, context) => { return context },//SUB : Subscriber,
    0x92: (data, context) => { return context },//PALETTE : Color Palette Definition,
    0x94: (data, context) => { return context },//LHRECORD : .WK? File Conversion Information,
    0x95: (data, context) => { return context },//LHNGRAPH : Named Graph Information,
    0x96: (data, context) => { return context },//SOUND : Sound Note,
    0x98: (data, context) => { return context },//LPR : Sheet Was Printed Using LINE.PRINT(,
    0x99: (data, context) => { return context },//STANDARDWIDTH : Standard Column Width,
    0x9A: (data, context) => { return context },//FNGROUPNAME : Function Group Name,
    0x9B: (data, context) => { return context },//FILTERMODE : Sheet Contains Filtered List,
    0x9C: (data, context) => { return context },//FNGROUPCOUNT : Built-in Function Group Count,
    0x9D: (data, context) => { return context },//AUTOFILTERINFO : Drop-Down Arrow Count,
    0x9E: (data, context) => { return context },//AUTOFILTER : AutoFilter Data,
    0xA0: (data, context) => { return context },//SCL : Window Zoom Magnification,
    0xA1: (data, context) => { return context },//SETUP : Page Setup,
    0xA9: (data, context) => { return context },//COORDLIST : Polygon Object Vertex Coordinates,
    0xAB: (data, context) => { return context },//GCW : Global Column-Width Flags,
    0xAE: (data, context) => { return context },//SCENMAN : Scenario Output Data,
    0xAF: (data, context) => { return context },//SCENARIO : Scenario Data,
    0xB0: (data, context) => { return context },//SXVIEW : View Definition,
    0xB1: (data, context) => { return context },//SXVD : View Fields,
    0xB2: (data, context) => { return context },//SXVI : View Item,
    0xB4: (data, context) => { return context },//SXIVD : Row/Column Field IDs,
    0xB5: (data, context) => { return context },//SXLI : Line Item Array,
    0xB6: (data, context) => { return context },//SXPI : Page Item,
    0xB8: (data, context) => { return context },//DOCROUTE : Routing Slip Information,
    0xB9: (data, context) => { return context },//RECIPNAME : Recipient Name,
    0xBC: (data, context) => { return context },//SHRFMLA : Shared Formula,
    0xBD: (data, context) => { return context },//MULRK : Multiple  RK Cells,
    0xBE: (data, context) => { return context },//MULBLANK : Multiple Blank Cells,
    0xC1: (data, context) => { return context },//MMS :  ADDMENU / DELMENU Record Group Count,
    0xC2: (data, context) => { return context },//ADDMENU : Menu Addition,
    0xC3: (data, context) => { return context },//DELMENU : Menu Deletion,
    0xC5: (data, context) => { return context },//SXDI : Data Item,
    0xC6: (data, context) => { return context },//SXDB : PivotTable Cache Data,
    0xCD: (data, context) => { return context },//SXSTRING : String,
    0xD0: (data, context) => { return context },//SXTBL : Multiple Consolidation Source Info,
    0xD1: (data, context) => { return context },//SXTBRGIITM : Page Item Name Count,
    0xD2: (data, context) => { return context },//SXTBPG : Page Item Indexes,
    0xD3: (data, context) => { return context },//OBPROJ : Visual Basic Project,
    0xD5: (data, context) => { return context },//SXIDSTM : Stream ID,
    0xD6: (data, context) => { return context },//RSTRING : Cell with Character Formatting,
    0xD7: (data, context) => { return context },//DBCELL : Stream Offsets,
    0xDA: (data, context) => {
        context.update_links = (data.readUInt8() >> 5) & 0x3
        return context
    },//BOOKBOOL : Workbook Option Flag,
    0xDC: (data, context) => { return context },//PARAMQRY : Query Parameters,
    0xDC: (data, context) => { return context },//SXEXT : External Source Information,
    0xDD: (data, context) => { return context },//SCENPROTECT : Scenario Protection,
    0xDE: (data, context) => { return context },//OLESIZE : Size of OLE Object,
    0xDF: (data, context) => { return context },//UDDESC : Description String for Chart Autoformat,
    0xE0: (data, context) => { return context },//XF : Extended Format,
    0xE1: (data, context) => { return context },//INTERFACEHDR : Beginning of User Interface Records,
    0xE2: (data, context) => { return context },//INTERFACEEND : End of User Interface Records,
    0xE3: (data, context) => { return context },//SXVS : View Source,
    0xE5: (data, context) => { return context },//MERGECELLS : Merged Cells,
    0xEA: (data, context) => { return context },//TABIDCONF : Sheet Tab ID of Conflict History,
    0xEB: (data, context) => { return context },//MSODRAWINGGROUP : Microsoft Office Drawing Group,
    0xEC: (data, context) => { return context },//MSODRAWING : Microsoft Office Drawing,
    0xED: (data, context) => { return context },//MSODRAWINGSELECTION : Microsoft Office Drawing Selection,
    0xF0: (data, context) => { return context },//SXRULE : PivotTable Rule Data,
    0xF1: (data, context) => { return context },//SXEX : PivotTable View Extended Information,
    0xF2: (data, context) => { return context },//SXFILT : PivotTable Rule Filter,
    0xF4: (data, context) => { return context },//SXDXF : Pivot Table Formatting,
    0xF5: (data, context) => { return context },//SXITM : Pivot Table Item Indexes,
    0xF6: (data, context) => { return context },//SXNAME : PivotTable Name,
    0xF7: (data, context) => { return context },//SXSELECT : PivotTable Selection Information,
    0xF8: (data, context) => { return context },//SXPAIR : PivotTable Name Pair,
    0xF9: (data, context) => { return context },//SXFMLA : Pivot Table Parsed Expression,
    0xFB: (data, context) => { return context },//SXFORMAT : PivotTable Format Record,
    0xFC: (data, context) => {
        let cstTotal = data.readUInt32LE(0);//Number of references to strings
        let cstUnique = data.readUInt32LE(4);//Number of unique strings
        let rgboriginal = data.slice(8);
        let slicedXL = data.slice(8);
        for (let i = 0; i < cstUnique; i++) {
            //XLUnicodeRichExtendedString
            let cch = slicedXL.readUInt16LE(0);
            let flags = slicedXL.readUInt8(2)
            let fHighByte = flags >> 7;
            let fExtSt = (flags >> 5) & 0x1;
            let fRichSt = (flags >> 4) & 0x1;
            let cRun = fRichSt == 1 ? slicedXL.readUInt16LE(3) : 0;
            let cbExtRst = fExtSt == 1 ? slicedXL.readUInt32LE(5) : 0;
            let extractOffset = 3 + (cRun > 0 ? 2 : 0) + (cbExtRst > 0 ? 4 : 0);
            //rgb
            let rgb = slicedXL.slice(extractOffset, extractOffset + cch * (fHighByte + 1)).toString(fHighByte == 1 ? 'utf16le' : 'utf8');
            extractOffset += cch * (fHighByte + 1);
            extractOffset += cRun * 4;
            extractOffset += cbExtRst
            slicedXL = slicedXL.slice(extractOffset);
            context.sst.push(rgb);
        }
        return context
    },//SST : Shared String Table,
    0xFD: (data, context) => {
        //Stores a plain string
        let rw = data.readUInt16LE(0);
        let col = data.readUInt16LE(2);
        let ixfe = data.readUInt16LE(4);//Index to the XFrecord
        let ixsst = data.readUInt32LE(6);//Index into the SSTrecord where actual string is stored
        let vl = context.sst[ixsst] || ""
        context.current_sheet.cells.push({
            rw,
            col,
            value: vl,
            js_v: `'${vl.replace(/\\/g, "\\\\").replace(/'/g, "\\'")}'`
        })
        return context
    },//LABELSST : Cell Value, String Constant/ SST,
    0xFF: (data, context) => {
        let dsst = data.readUInt16LE(0);//Number of strings in each bucket
        let rfisstinf = data.slice(2, 8 * dsst + 2);// Array of ISSTINF structures.
        let string_positions = []
        for (let i = 0; i < dsst; i++) {
            string_positions.push({
                ib: rfisstinf.readUInt32LE(0),// Stream position where the strings begin
                cb: rfisstinf.readUInt16LE(4),// Offset into the SST record that points where the bucket begins
            })
        }
        return context
    },//EXTSST : Extended Shared String Table,
    0x100: (data, context) => { return context },//SXVDEX : Extended PivotTable View Fields,
    0x103: (data, context) => { return context },//SXFORMULA : PivotTable Formula Record,
    0x122: (data, context) => { return context },//SXDBEX : PivotTable Cache Data,
    0x13D: (data, context) => { return context },//TABID : Sheet Tab Index Array,
    0x160: (data, context) => { return context },//USESELFS : Natural Language Formulas Flag,
    0x161: (data, context) => { return context },//DSF : Double Stream File,
    0x162: (data, context) => { return context },//XL5MODIFY : Flag for  DSF,
    0x1A5: (data, context) => { return context },//FILESHARING2 : File-Sharing Information for Shared Lists,
    0x1A9: (data, context) => { return context },//USERBVIEW : Workbook Custom View Settings,
    0x1AA: (data, context) => { return context },//USERSVIEWBEGIN : Custom View Settings,
    0x1AB: (data, context) => { return context },//USERSVIEWEND : End of Custom View Records,
    0x1AD: (data, context) => { return context },//QSI : External Data Range,
    0x1AE: (data, context) => {
        //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/31ed3738-e4ff-4b60-804c-ac49ac1ee6c0
        //TODO:
        let ctab = data.readUInt16LE()
        let cch = data.readUInt16LE(2)
        let dt_ofst = 4;
        if (cch >= 0x0001 && cch <= 0x00FF) {
            let virt_path = processXLUnicodeStringNoCch(data.slice(dt_ofst, dt_ofst + type))
            dt_ofst = dt_ofst + type
            if (virt_path.value === " ") {
                //This record specifies an unused supporting link.
            } else if (virt_path.value === "") {
                //This record specifies a same-sheet referencing type of supporting link.
            } else {
                //TODO
                for (let i = 0; i < ctab; i++) {
                    let str = processXLUnicodeString(data.slice(dt_ofst))
                    context.external_sheets.push({
                        cch,
                        ctab,
                        virt_path,
                        rgst: str.value
                    })
                    dt_ofst = dt_ofst + str.byte_length
                }
            }

        }
        context.superbook_type = cch
        return context
    },//SUPBOOK : Supporting WorkbooTABIDk,
    0x1AF: (data, context) => { return context },//PROT4REV : Shared Workbook Protection Flag,
    0x1B0: (data, context) => { return context },//CONDFMT : Conditional Formatting Range Information,
    0x1B1: (data, context) => { return context },//CF : Conditional Formatting Conditions,
    0x1B2: (data, context) => { return context },//DVAL : Data Validation Information,
    0x1B5: (data, context) => { return context },//DCONBIN : Data Consolidation Information,
    0x1B6: (data, context) => { return context },//TXO : Text Object,
    0x1B7: (data, context) => { return context },//REFRESHALL : Refresh Flag,
    0x1B8: (data, context) => { return context },//HLINK : Hyperlink,
    0x1BB: (data, context) => { return context },//SXFDBTYPE : SQL Datatype Identifier,
    0x1BC: (data, context) => { return context },//PROT4REVPASS : Shared Workbook Protection Password,
    0x1BE: (data, context) => { return context },//DV : Data Validation Criteria,
    0x1C0: (data, context) => { return context },//EXCEL9FILE : Excel 9 File,
    0x1C1: (data, context) => { return context },//RECALCID : Recalc Information,
    0x1C2: (data, context) => { return context },//EntExU2: Application-specific cache,
    0x200: (data, context) => { return context },//DIMENSIONS : Cell Table Size,
    0x201: (data, context) => { return context },//BLANK : Cell Value, Blank Cell,
    0x203: (data, context) => {
        return context
    },//NUMBER : Cell Value, Floating-Point Number,
    0x204: (data, context) => {
        return context
    },//LABEL : Cell Value, String Constant,
    0x205: (data, context) => {
        return context
    },//BOOLERR : Cell Value, Boolean or Error,
    0x207: (data, context) => {
        return context
    },//STRING : String Value of a Formula,
    0x208: (data, context) => {
        //Stores a plain string
        let rw = data.readUInt16LE(0);
        let colMic = data.readUInt16LE(2);//First column
        let colMac = data.readUInt16LE(4);// Last column + 1
        let miyRw = data.readUInt16LE(6);
        let irwMac = data.readUInt16LE(8);
        let grbit = data.readUInt16LE(12);//Option flags
        if (((grbit >> 8) & 0x80) != 0) {
            //fGhostDirty = 1
            //Index XF
            let ixfe = data.readUInt16LE(14)
        }
        return context
    },//ROW : Describes a Row,
    0x20B: (data, context) => { return context },//INDEX : Index Record,
    0x218: (data, context) => { return context },//NAME : Defined Name,
    0x221: (data, context) => { return context },//ARRAY : Array-Entered Formula,
    0x223: (data, context) => { return context },//EXTERNNAME : Externally Referenced Name,
    0x225: (data, context) => { return context },//DEFAULTROWHEIGHT : Default Row Height,
    0x231: (data, context) => { return context },//FONT : Font Description,
    0x236: (data, context) => { return context },//TABLE : Data Table,
    0x23E: (data, context) => {
        context.work_on_next_sheet()
        return context
    },//WINDOW2 : Sheet Window Information,
    0x27E: (data, context) => {
        let rw = data.readUInt16LE(0);
        let col = data.readUInt16LE(2);
        let number = data.readUInt32LE(6)//[0,0,8,64]
        let fx100 = number & 0x01
        let fInt = number & 0x02
        number = number >> 2
        if (fInt == 0) {
            // LE >> 2 
            let num_buff = Buffer.alloc(8)
            num_buff.writeBigUInt64LE(BigInt(number) << BigInt(34))
            number = ieee754.fromIEEE754Double(num_buff.reverse())

            if (fx100) {
                number = number / 100
            }
        } else {
            number = (data.readInt32LE(6) >> 2)
            if (fx100) {
                number = number / 100
            }
        }


        context.add_cell({
            rw,
            col,
            value: number,
            js_v: number
        })
        return context
    },//RK : Cell Value, RK Number,
    0x293: (data, context) => { return context },//STYLE : Style Information,
    0x406: (data, context) => {
        return context
    },//FORMULA : Cell Formula,
    0x41E: (data, context) => { return context },//FORMAT : Number Format,
    0x800: (data, context) => { return context },//HLINKTOOLTIP : Hyperlink Tooltip,
    0x801: (data, context) => { return context },//WEBPUB : Web Publish Item,
    0x802: (data, context) => { return context },//QSISXTAG : PivotTable and Query Table Extensions,
    0x803: (data, context) => { return context },//DBQUERYEXT : Database Query Extensions,
    0x804: (data, context) => { return context },//EXTSTRING :  FRT String,
    0x805: (data, context) => { return context },//TXTQUERY : Text Query Information,
    0x806: (data, context) => { return context },//QSIR : Query Table Formatting,
    0x807: (data, context) => { return context },//QSIF : Query Table Field Formatting,
    0x809: (data, context) => {
        let ver = data.readUInt16BE(0);
        if (ver != 6) {
            throw Error(`Invalid BIFF version in BOF: ${ver}`);
        }
        let stream_type = data.readUInt16LE(2);
        //0x0005 workbook, 0x0010 dialog, 0x0020 cchart, 0x0040 macro


        let rupBuild = data.readUInt16BE(4);
        let rupYear = data.readUInt16BE(6);
        let flags = data.readUInt32BE(8);
        let verXLHigh = (flags >> 14) & 0x7;
        let verLowestBiff = data.readUInt8(12);
        let verLastXLSaved = data.readUInt8(13) >> 4;
        context.excel_version = excelVersionName(verLastXLSaved)
        return context
    },//BOF : Beginning of File,
    0x80A: (data, context) => { return context },//OLEDBCONN : OLE Database Connection,
    0x80B: (data, context) => { return context },//WOPT : Web Options,
    0x80C: (data, context) => { return context },//SXVIEWEX : Pivot Table OLAP Extensions,
    0x80D: (data, context) => { return context },//SXTH : PivotTable OLAP Hierarchy,
    0x80E: (data, context) => { return context },//SXPIEX : OLAP Page Item Extensions,
    0x80F: (data, context) => { return context },//SXVDTEX : View Dimension OLAP Extensions,
    0x810: (data, context) => { return context },//SXVIEWEX9 : Pivot Table Extensions,
    0x812: (data, context) => {
        return context
    },//CONTINUEFRT : Continued  FRT,
    0x813: (data, context) => { return context },//REALTIMEDATA : Real-Time Data (RTD),
    0x862: (data, context) => { return context },//SHEETEXT : Extra Sheet Info,
    0x863: (data, context) => { return context },//BOOKEXT : Extra Book Info,
    0x864: (data, context) => { return context },//SXADDL : Pivot Table Additional Info,
    0x865: (data, context) => { return context },//CRASHRECERR : Crash Recovery Error,
    0x866: (data, context) => { return context },//HFPicture : Header / Footer Picture,
    0x867: (data, context) => { return context },//FEATHEADR : Shared Feature Header,
    0x868: (data, context) => { return context },//FEAT : Shared Feature Record,
    0x86A: (data, context) => { return context },//DATALABEXT : Chart Data Label Extension,
    0x86B: (data, context) => { return context },//DATALABEXTCONTENTS : Chart Data Label Extension Contents,
    0x86C: (data, context) => { return context },//CELLWATCH : Cell Watch,
    0x86d: (data, context) => { return context },//FEATINFO : Shared Feature Info Record,
    0x871: (data, context) => { return context },//FEATHEADR11 : Shared Feature Header 11,
    0x872: (data, context) => { return context },//FEAT11 : Shared Feature 11 Record,
    0x873: (data, context) => { return context },//FEATINFO11 : Shared Feature Info 11 Record,
    0x874: (data, context) => { return context },//DROPDOWNOBJIDS : Drop Down Object,
    0x875: (data, context) => { return context },//CONTINUEFRT11 : Continue  FRT 11,
    0x876: (data, context) => { return context },//DCONN : Data Connection,
    0x877: (data, context) => { return context },//LIST12 : Extra Table Data Introduced in Excel 2007,
    0x878: (data, context) => { return context },//FEAT12 : Shared Feature 12 Record,
    0x879: (data, context) => { return context },//CONDFMT12 : Conditional Formatting Range Information 12,
    0x87A: (data, context) => { return context },//CF12 : Conditional Formatting Condition 12,
    0x87B: (data, context) => { return context },//CFEX : Conditional Formatting Extension,
    0x87C: (data, context) => { return context },//XFCRC : XF Extensions Checksum,
    0x87D: (data, context) => { return context },//XFEXT : XF Extension,
    0x87E: (data, context) => { return context },//EZFILTER12 : AutoFilter Data Introduced in Excel 2007,
    0x87F: (data, context) => { return context },//CONTINUEFRT12 : Continue FRT 12,
    0x881: (data, context) => { return context },//SXADDL12 : Additional Workbook Connections Information,
    0x884: (data, context) => { return context },//MDTINFO : Information about a Metadata Type,
    0x885: (data, context) => { return context },//MDXSTR : MDX Metadata String,
    0x886: (data, context) => { return context },//MDXTUPLE : Tuple MDX Metadata,
    0x887: (data, context) => { return context },//MDXSET : Set MDX Metadata,
    0x888: (data, context) => { return context },//MDXPROP : Member Property MDX Metadata,
    0x889: (data, context) => { return context },//MDXKPI : Key Performance Indicator MDX Metadata,
    0x88A: (data, context) => { return context },//MDTB : Block of Metadata Records,
    0x88B: (data, context) => { return context },//PLV : Page Layout View Settings in Excel 2007,
    0x88C: (data, context) => { return context },//COMPAT12 : Compatibility Checker 12,
    0x88D: (data, context) => { return context },//DXF : Differential XF,
    0x88E: (data, context) => { return context },//TABLESTYLES : Table Styles,
    0x88F: (data, context) => { return context },//TABLESTYLE : Table Style,
    0x890: (data, context) => { return context },//TABLESTYLEELEMENT : Table Style Element,
    0x892: (data, context) => { return context },//STYLEEXT : Named Cell Style Extension,
    0x893: (data, context) => { return context },//NAMEPUBLISH : Publish To Excel Server Data for Name,
    0x894: (data, context) => { return context },//NAMECMT : Name Comment,
    0x895: (data, context) => { return context },//SORTDATA12 : Sort Data 12,
    0x896: (data, context) => { return context },//THEME : Theme,
    0x897: (data, context) => { return context },//GUIDTYPELIB : VB Project Typelib GUID,
    0x898: (data, context) => { return context },//FNGRP12 : Function Group,
    0x899: (data, context) => { return context },//NAMEFNGRP12 : Extra Function Group,
    0x89A: (data, context) => { return context },//MTRSETTINGS : Multi-Threaded Calculation Settings,
    0x89B: (data, context) => { return context },//COMPRESSPICTURES : Automatic Picture Compression Mode,
    0x89C: (data, context) => { return context },//HEADERFOOTER : Header Footer,
    0x8A3: (data, context) => { return context },//FORCEFULLCALCULATION : Force Full Calculation Settings,
    0x8c1: (data, context) => { return context },//LISTOBJ : List Object,
    0x8c2: (data, context) => { return context },//LISTFIELD : List Field,
    0x8c3: (data, context) => { return context },//LISTDV : List Data Validation,
    0x8c4: (data, context) => { return context },//LISTCONDFMT : List Conditional Formatting,
    0x8c5: (data, context) => { return context },//LISTCF : List Cell Formatting,
    0x8c6: (data, context) => { return context },//FMQRY : Filemaker queries,
    0x8c7: (data, context) => { return context },//FMSQRY : File maker queries,
    0x8c8: (data, context) => { return context },//PLV : Page Layout View in Mac Excel 11,
    0x8c9: (data, context) => { return context },//LNEXT : Extension information for borders in Mac Office 11,
    0x8ca: (data, context) => { return context },//MKREXT : Extension information for markers in Mac Office 11
}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/39a0757a-c7bb-4e85-b144-3e7837b059d7
 * @param {Buffer} data 
 */
function parseFormulaValue(data) {
    let fExprO = data.readUInt16LE(6)
    if (fExprO === 0xFFFF) {
        let byte1_type = data.readUInt8(0)
        return {
            type: byte1_type === 0 ? "STRING" : byte1_type === 1 ? "BOOLEAN" : byte1_type === 2 ? "ERROR" : "BLANK"
        }
    } else {
        return {
            type: "IEEE754",
            value: ieee754.fromIEEE754Double(data)
        }

    }

}

function excelVersionName(verXLHigh) {
    let verName = {
        0: "Excel 97",
        1: "Excel 2000",
        2: "Excel 2002",
        3: "Office Excel 2003",
        4: "Office Excel 2007",
        5: "Excel 2010",
        6: "Excel 2013"

    }[verXLHigh]

    return verName ? verName : "Not a valid version"
}

module.exports.RECORD_CONSTRUCTOR = RECORD_CONSTRUCTOR