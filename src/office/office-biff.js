const ieee754 = require('./ieee754')
const { PT_TOKEN_EXCEL: TOKEN_EXCEL, PT_TOKEN_EXCEL_MAP: TOKEN_EXCEL_MAP, FT_CODE_FUNCTION_EXCEL: CODE_FUNCTION_EXCEL, RECORD_OP_CODES: OP_CODES, BUILT_IN_LABEL_NAMES, OP_CODES_MAP, STREAM_TYPES } = require('./biff-constants')

const { process_RGCE, structure_RGCE, processXLUnicodeStringNoCch, processXLUnicodeString } = require('./biff-formulas')
const {RECORD_CONSTRUCTOR} = require('./biff-record')
class BiffDocument {
    constructor(){
        this.sheet_pos = 0;
        this.current_sheet = null;
        this.biff_version = 0;
        this.excel_version = "";
        this.sheet_list = []
        this.user_name = ""
        this.last_opcode = 0;
        this.sst = [];
        this.update_links = 0;
        this.superbook = []
        this.superbook_type = 0;
        this.external_sheets = []
    }
    work_on_next_sheet(){
        this.sheet_pos++;
        this.current_sheet = this.sheet_list[this.sheet_pos];
    }
    add_cell(cell){
        this.current_sheet.cells.push(cell)
    }
    add_sheet(sheet){
        this.sheet_list.push(sheet)
    }

    generate_javascript(){
        let auto_open = null
        let excel_jscript = "let CELL_LIST_OBJ = {}\n";
        excel_jscript += `
let CELL_LIST = new Proxy(CELL_LIST_OBJ,{
    get: function(target,prop){
        if(prop.startsWith("cell_")){
            let fnc = target[prop];
            if(fnc){
                return fnc
            }else{
                return function(){return ""}
            }
        }
        
    }
})
let FUNCTIONS = {
    RUN : function(...args){
        console.log(args)
    },
    FORMULA: function(...args){
        console.log(args)
    },
}

let context = new Proxy({},{
    get: function(target,prop){
        if(prop.startsWith("functions_")){
            let fun_name = prop.slice(10)
            console.log(prop)
            return FUNCTIONS[fun_name]
        }
        
    }
})
        `
        let sheet_list =  {}
        for(let i_sheet =0; i_sheet < this.sheet_list.length; i_sheet++){
            let sheet = this.sheet_list[i_sheet]
            excel_jscript += `//----------------- ${sheet.name} -----------------\n`
            for(let i_cell =0; i_cell < sheet.cells.length; i_cell++){
                let cell = sheet.cells[i_cell]
                excel_jscript += `CELL_LIST_OBJ.cell_${sheet.name}_${cell.rw}_${cell.col} = function(){return ${cell.js_v || cell.value}}\n`
            }
            for(let i_label =0; i_label < sheet.labels.length; i_label++){
                let lbl = sheet.labels[i_label]
                let label_values = lbl.named_formula
                let label_code = ""
                for(let i_named = 0; i_named < label_values.length; i_named++){
                    label_code += `${label_values[i_named].js_v};`
                }
                excel_jscript += `function label_${sheet.name}_${lbl.name}(){return ${label_code}}\n`
                if(lbl.name.toLowerCase() === "auto_open"){
                    auto_open = `label_${sheet.name}_${lbl.name}`
                }
            }
        }
        if(auto_open){
            excel_jscript += `${auto_open}()`
        }
        return excel_jscript
    }

    /**
     * 
     * @param {Buffer} buff 
     */
    static from_buffer(buff) {
        let streamType = 0;
        let userName = "";
        let sheets = []
        let current_sheet = null
        let sheet_pos = 0;
        let offset = 0;
        let biffversion = 0;//0=BIFF2, 2=BIFF3, 4 = BIFF3, 8 = BIFF5/7/8
        let excelVersion = "";
        let xfrecords = []
        let section_list = []
        let sst = []

        let shared_protection = null

        let update_links = 0;
        let last_opcode = 0;
        let external_sheets = []
        let sup_book = []
        let superbook_type = 0;
        let context = new BiffDocument()

        while (offset < buff.length) {
            try {
                let opcode = buff.readUInt16LE(offset)
                let lngth = buff.readUInt16LE(offset + 2)
                let data = buff.slice(offset + 4, offset + 4 + lngth)
                offset += 4 + lngth

                RECORD_CONSTRUCTOR[opcode](data,context)

                if(opcode != OP_CODES_MAP.CONTINUE){
                    context.last_opcode = opcode
                }
            } catch (err) { }
        }
        return context
    }
}




const C_TAB_MAP = {
    0x0000: "BEEP",
    0x0001: "OPEN",
    0x0002: "OPEN.LINKS",
    0x0003: "CLOSE.ALL",
    0x0004: "SAVE",
    0x0005: "SAVE.AS",
    0x0006: "FILE.DELETE",
    0x0007: "PAGE.SETUP",
    0x0008: "PRINT",
    0x0009: "PRINTER.SETUP",
    0x000A: "QUIT",
    0x000B: "NEW.WINDOW",
    0x000C: "ARRANGE.ALL",
    0x000D: "WINDOW.SIZE",
    0x000E: "WINDOW.MOVE",
    0x000F: "FULL",
    0x0010: "CLOSE",
    0x0011: "RUN",
    0x0016: "SET.PRINT.AREA",
    0x0017: "SET.PRINT.TITLES",
    0x0018: "SET.PAGE.BREAK",
    0x0019: "REMOVE.PAGE.BREAK",
    0x001A: "FONT",
    0x001B: "DISPLAY",
    0x001C: "PROTECT.DOCUMENT",
    0x001D: "PRECISION",
    0x001E: "A1.R1C1",
    0x001F: "CALCULATE.NOW",
    0x0020: "CALCULATION",
    0x0022: "DATA.FIND",
    0x0023: "EXTRACT",
    0x0024: "DATA.DELETE",
    0x0025: "SET.DATABASE",
    0x0026: "SET.CRITERIA",
    0x0027: "SORT",
    0x0028: "DATA.SERIES",
    0x0029: "TABLE",
    0x002A: "FORMAT.NUMBER",
    0x002B: "ALIGNMENT",
    0x002C: "STYLE",
    0x002D: "BORDER",
    0x002E: "CELL.PROTECTION",
    0x002F: "COLUMN.WIDTH",
    0x0030: "UNDO",
    0x0031: "CUT",
    0x0032: "COPY",
    0x0033: "PASTE",
    0x0034: "CLEAR",
    0x0035: "PASTE.SPECIAL",
    0x0036: "EDIT.DELETE",
    0x0037: "INSERT",
    0x0038: "FILL.RIGHT",
    0x0039: "FILL.DOWN",
    0x003D: "DEFINE.NAME",
    0x003E: "CREATE.NAMES",
    0x003F: "FORMULA.GOTO",
    0x0040: "FORMULA.FIND",
    0x0041: "SELECT.LAST.CELL",
    0x0042: "SHOW.ACTIVE.CELL",
    0x0043: "GALLERY.AREA",
    0x0044: "GALLERY.BAR",
    0x0045: "GALLERY.COLUMN",
    0x0046: "GALLERY.LINE",
    0x0047: "GALLERY.PIE",
    0x0048: "GALLERY.SCATTER",
    0x0049: "COMBINATION",
    0x004A: "PREFERRED",
    0x004B: "ADD.OVERLAY",
    0x004C: "GRIDLINES",
    0x004D: "SET.PREFERRED",
    0x004E: "AXES",
    0x004F: "LEGEND",
    0x0050: "ATTACH.TEXT",
    0x0051: "ADD.ARROW",
    0x0052: "SELECT.CHART",
    0x0053: "SELECT.PLOT.AREA",
    0x0054: "PATTERNS",
    0x0055: "MAIN.CHART",
    0x0056: "OVERLAY",
    0x0057: "SCALE",
    0x0058: "FORMAT.LEGEND",
    0x0059: "FORMAT.TEXT",
    0x005A: "EDIT.REPEAT",
    0x005B: "PARSE",
    0x005C: "JUSTIFY",
    0x005D: "HIDE",
    0x005E: "UNHIDE",
    0x005F: "WORKSPACE",
    0x0060: "FORMULA",
    0x0061: "FORMULA.FILL",
    0x0062: "FORMULA.ARRAY",
    0x0063: "DATA.FIND.NEXT",
    0x0064: "DATA.FIND.PREV",
    0x0065: "FORMULA.FIND.NEXT",
    0x0066: "FORMULA.FIND.PREV",
    0x0067: "ACTIVATE",
    0x0068: "ACTIVATE.NEXT",
    0x0069: "ACTIVATE.PREV",
    0x006A: "UNLOCKED.NEXT",
    0x006B: "UNLOCKED.PREV",
    0x006C: "COPY.PICTURE",
    0x006D: "SELECT",
    0x006E: "DELETE.NAME",
    0x006F: "DELETE.FORMAT",
    0x0070: "VLINE",
    0x0071: "HLINE",
    0x0072: "VPAGE",
    0x0073: "HPAGE",
    0x0074: "VSCROLL",
    0x0075: "HSCROLL",
    0x0076: "ALERT",
    0x0077: "NEW",
    0x0078: "CANCEL.COPY",
    0x0079: "SHOW.CLIPBOARD",
    0x007A: "MESSAGE",
    0x007C: "PASTE.LINK",
    0x007D: "APP.ACTIVATE",
    0x007E: "DELETE.ARROW",
    0x007F: "ROW.HEIGHT",
    0x0080: "FORMAT.MOVE",
    0x0081: "FORMAT.SIZE",
    0x0082: "FORMULA.REPLACE",
    0x0083: "SEND.KEYS",
    0x0084: "SELECT.SPECIAL",
    0x0085: "APPLY.NAMES",
    0x0086: "REPLACE.FONT",
    0x0087: "FREEZE.PANES",
    0x0088: "SHOW.INFO",
    0x0089: "SPLIT",
    0x008A: "ON.WINDOW",
    0x008B: "ON.DATA",
    0x008C: "DISABLE.INPUT",
    0x008E: "OUTLINE",
    0x008F: "LIST.NAMES",
    0x0090: "FILE.CLOSE",
    0x0091: "SAVE.WORKBOOK",
    0x0092: "DATA.FORM",
    0x0093: "COPY.CHART",
    0x0094: "ON.TIME",
    0x0095: "WAIT",
    0x0096: "FORMAT.FONT",
    0x0097: "FILL.UP",
    0x0098: "FILL.LEFT",
    0x0099: "DELETE.OVERLAY",
    0x009B: "SHORT.MENUS",
    0x009F: "SET.UPDATE.STATUS",
    0x00A1: "COLOR.PALETTE",
    0x00A2: "DELETE.STYLE",
    0x00A3: "WINDOW.RESTORE",
    0x00A4: "WINDOW.MAXIMIZE",
    0x00A6: "CHANGE.LINK",
    0x00A7: "CALCULATE.DOCUMENT",
    0x00A8: "ON.KEY",
    0x00A9: "APP.RESTORE",
    0x00AA: "APP.MOVE",
    0x00AB: "APP.SIZE",
    0x00AC: "APP.MINIMIZE",
    0x00AD: "APP.MAXIMIZE",
    0x00AE: "BRING.TO.FRONT",
    0x00AF: "SEND.TO.BACK",
    0x00B9: "MAIN.CHART.TYPE",
    0x00BA: "OVERLAY.CHART.TYPE",
    0x00BB: "SELECT.END",
    0x00BC: "OPEN.MAIL",
    0x00BD: "SEND.MAIL",
    0x00BE: "STANDARD.FONT",
    0x00BF: "CONSOLIDATE",
    0x00C0: "SORT.SPECIAL",
    0x00C1: "GALLERY.3D.AREA",
    0x00C2: "GALLERY.3D.COLUMN",
    0x00C3: "GALLERY.3D.LINE",
    0x00C4: "GALLERY.3D.PIE",
    0x00C5: "VIEW.3D",
    0x00C6: "GOAL.SEEK",
    0x00C7: "WORKGROUP",
    0x00C8: "FILL.GROUP",
    0x00C9: "UPDATE.LINK",
    0x00CA: "PROMOTE",
    0x00CB: "DEMOTE",
    0x00CC: "SHOW.DETAIL",
    0x00CE: "UNGROUP",
    0x00CF: "OBJECT.PROPERTIES",
    0x00D0: "SAVE.NEW.OBJECT",
    0x00D1: "SHARE",
    0x00D2: "SHARE.NAME",
    0x00D3: "DUPLICATE",
    0x00D4: "APPLY.STYLE",
    0x00D5: "ASSIGN.TO.OBJECT",
    0x00D6: "OBJECT.PROTECTION",
    0x00D7: "HIDE.OBJECT",
    0x00D8: "SET.EXTRACT",
    0x00D9: "CREATE.PUBLISHER",
    0x00DA: "SUBSCRIBE.TO",
    0x00DB: "ATTRIBUTES",
    0x00DC: "SHOW.TOOLBAR",
    0x00DE: "PRINT.PREVIEW",
    0x00DF: "EDIT.COLOR",
    0x00E0: "SHOW.LEVELS",
    0x00E1: "FORMAT.MAIN",
    0x00E2: "FORMAT.OVERLAY",
    0x00E3: "ON.RECALC",
    0x00E4: "EDIT.SERIES",
    0x00E5: "DEFINE.STYLE",
    0x00F0: "LINE.PRINT",
    0x00F3: "ENTER.DATA",
    0x00F9: "GALLERY.RADAR",
    0x00FA: "MERGE.STYLES",
    0x00FB: "EDITION.OPTIONS",
    0x00FC: "PASTE.PICTURE",
    0x00FD: "PASTE.PICTURE.LINK",
    0x00FE: "SPELLING",
    0x0100: "ZOOM",
    0x0103: "INSERT.OBJECT",
    0x0104: "WINDOW.MINIMIZE",
    0x0109: "SOUND.NOTE",
    0x010A: "SOUND.PLAY",
    0x010B: "FORMAT.SHAPE",
    0x010C: "EXTEND.POLYGON",
    0x010D: "FORMAT.AUTO",
    0x0110: "GALLERY.3D.BAR",
    0x0111: "GALLERY.3D.SURFACE",
    0x0112: "FILL.AUTO",
    0x0114: "CUSTOMIZE.TOOLBAR",
    0x0115: "ADD.TOOL",
    0x0116: "EDIT.OBJECT",
    0x0117: "ON.DOUBLECLICK",
    0x0118: "ON.ENTRY",
    0x0119: "WORKBOOK.ADD",
    0x011A: "WORKBOOK.MOVE",
    0x011B: "WORKBOOK.COPY",
    0x011C: "WORKBOOK.OPTIONS",
    0x011D: "SAVE.WORKSPACE",
    0x0120: "CHART.WIZARD",
    0x0121: "DELETE.TOOL",
    0x0122: "MOVE.TOOL",
    0x0123: "WORKBOOK.SELECT",
    0x0124: "WORKBOOK.ACTIVATE",
    0x0125: "ASSIGN.TO.TOOL",
    0x0127: "COPY.TOOL",
    0x0128: "RESET.TOOL",
    0x0129: "CONSTRAIN.NUMERIC",
    0x012A: "PASTE.TOOL",
    0x012E: "WORKBOOK.NEW",
    0x0131: "SCENARIO.CELLS",
    0x0132: "SCENARIO.DELETE",
    0x0133: "SCENARIO.ADD",
    0x0134: "SCENARIO.EDIT",
    0x0135: "SCENARIO.SHOW",
    0x0136: "SCENARIO.SHOW.NEXT",
    0x0137: "SCENARIO.SUMMARY",
    0x0138: "PIVOT.TABLE.WIZARD",
    0x0139: "PIVOT.FIELD.PROPERTIES",
    0x013A: "PIVOT.FIELD",
    0x013B: "PIVOT.ITEM",
    0x013C: "PIVOT.ADD.FIELDS",
    0x013E: "OPTIONS.CALCULATION",
    0x013F: "OPTIONS.EDIT",
    0x0140: "OPTIONS.VIEW",
    0x0141: "ADDIN.MANAGER",
    0x0142: "MENU.EDITOR",
    0x0143: "ATTACH.TOOLBARS",
    0x0144: "VBAActivate",
    0x0145: "OPTIONS.CHART",
    0x0148: "VBA.INSERT.FILE",
    0x014A: "VBA.PROCEDURE.DEFINITION",
    0x0150: "ROUTING.SLIP",
    0x0152: "ROUTE.DOCUMENT",
    0x0153: "MAIL.LOGON",
    0x0156: "INSERT.PICTURE",
    0x0157: "EDIT.TOOL",
    0x0158: "GALLERY.DOUGHNUT",
    0x015E: "CHART.TREND",
    0x0160: "PIVOT.ITEM.PROPERTIES",
    0x0162: "WORKBOOK.INSERT",
    0x0163: "OPTIONS.TRANSITION",
    0x0164: "OPTIONS.GENERAL",
    0x0172: "FILTER.ADVANCED",
    0x0175: "MAIL.ADD.MAILER",
    0x0176: "MAIL.DELETE.MAILER",
    0x0177: "MAIL.REPLY",
    0x0178: "MAIL.REPLY.ALL",
    0x0179: "MAIL.FORWARD",
    0x017A: "MAIL.NEXT.LETTER",
    0x017B: "DATA.LABEL",
    0x017C: "INSERT.TITLE",
    0x017D: "FONT.PROPERTIES",
    0x017E: "MACRO.OPTIONS",
    0x017F: "WORKBOOK.HIDE",
    0x0180: "WORKBOOK.UNHIDE",
    0x0181: "WORKBOOK.DELETE",
    0x0182: "WORKBOOK.NAME",
    0x0184: "GALLERY.CUSTOM",
    0x0186: "ADD.CHART.AUTOFORMAT",
    0x0187: "DELETE.CHART.AUTOFORMAT",
    0x0188: "CHART.ADD.DATA",
    0x0189: "AUTO.OUTLINE",
    0x018A: "TAB.ORDER",
    0x018B: "SHOW.DIALOG",
    0x018C: "SELECT.ALL",
    0x018D: "UNGROUP.SHEETS",
    0x018E: "SUBTOTAL.CREATE",
    0x018F: "SUBTOTAL.REMOVE",
    0x0190: "RENAME.OBJECT",
    0x019C: "WORKBOOK.SCROLL",
    0x019D: "WORKBOOK.NEXT",
    0x019E: "WORKBOOK.PREV",
    0x019F: "WORKBOOK.TAB.SPLIT",
    0x01A0: "FULL.SCREEN",
    0x01A1: "WORKBOOK.PROTECT",
    0x01A4: "SCROLLBAR.PROPERTIES",
    0x01A5: "PIVOT.SHOW.PAGES",
    0x01A6: "TEXT.TO.COLUMNS",
    0x01A7: "FORMAT.CHARTTYPE",
    0x01A8: "LINK.FORMAT",
    0x01A9: "TRACER.DISPLAY",
    0x01AE: "TRACER.NAVIGATE",
    0x01AF: "TRACER.CLEAR",
    0x01B0: "TRACER.ERROR",
    0x01B1: "PIVOT.FIELD.GROUP",
    0x01B2: "PIVOT.FIELD.UNGROUP",
    0x01B3: "CHECKBOX.PROPERTIES",
    0x01B4: "LABEL.PROPERTIES",
    0x01B5: "LISTBOX.PROPERTIES",
    0x01B6: "EDITBOX.PROPERTIES",
    0x01B7: "PIVOT.REFRESH",
    0x01B8: "LINK.COMBO",
    0x01B9: "OPEN.TEXT",
    0x01BA: "HIDE.DIALOG",
    0x01BB: "SET.DIALOG.FOCUS",
    0x01BC: "ENABLE.OBJECT",
    0x01BD: "PUSHBUTTON.PROPERTIES",
    0x01BE: "SET.DIALOG.DEFAULT",
    0x01BF: "FILTER",
    0x01C0: "FILTER.SHOW.ALL",
    0x01C1: "CLEAR.OUTLINE",
    0x01C2: "FUNCTION.WIZARD",
    0x01C3: "ADD.LIST.ITEM",
    0x01C4: "SET.LIST.ITEM",
    0x01C5: "REMOVE.LIST.ITEM",
    0x01C6: "SELECT.LIST.ITEM",
    0x01C7: "SET.CONTROL.VALUE",
    0x01C8: "SAVE.COPY.AS",
    0x01CA: "OPTIONS.LISTS.ADD",
    0x01CB: "OPTIONS.LISTS.DELETE",
    0x01CC: "SERIES.AXES",
    0x01CD: "SERIES.X",
    0x01CE: "SERIES.Y",
    0x01CF: "ERRORBAR.X",
    0x01D0: "ERRORBAR.Y",
    0x01D1: "FORMAT.CHART",
    0x01D2: "SERIES.ORDER",
    0x01D3: "MAIL.LOGOFF",
    0x01D4: "CLEAR.ROUTING.SLIP",
    0x01D5: "APP.ACTIVATE.MICROSOFT",
    0x01D6: "MAIL.EDIT.MAILER",
    0x01D7: "ON.SHEET",
    0x01D8: "STANDARD.WIDTH",
    0x01D9: "SCENARIO.MERGE",
    0x01DA: "SUMMARY.INFO",
    0x01DB: "FIND.FILE",
    0x01DC: "ACTIVE.CELL.FONT",
    0x01DD: "ENABLE.TIPWIZARD",
    0x01DE: "VBA.MAKE.ADDIN",
    0x01E0: "INSERTDATATABLE",
    0x01E1: "WORKGROUP.OPTIONS",
    0x01E2: "MAIL.SEND.MAILER",
    0x01E5: "AUTOCORRECT",
    0x01E9: "POST.DOCUMENT",
    0x01EB: "PICKLIST",
    0x01ED: "VIEW.SHOW",
    0x01EE: "VIEW.DEFINE",
    0x01EF: "VIEW.DELETE",
    0x01FD: "SHEET.BACKGROUND",
    0x01FE: "INSERT.MAP.OBJECT",
    0x01FF: "OPTIONS.MENONO",
    0x0205: "MSOCHECKS",
    0x0206: "NORMAL",
    0x0207: "LAYOUT",
    0x0208: "RM.PRINT.AREA",
    0x0209: "CLEAR.PRINT.AREA",
    0x020A: "ADD.PRINT.AREA",
    0x020B: "MOVE.BRK",
    0x0221: "HIDECURR.NOTE",
    0x0222: "HIDEALL.NOTES",
    0x0223: "DELETE.NOTE",
    0x0224: "TRAVERSE.NOTES",
    0x0225: "ACTIVATE.NOTES",
    0x026C: "PROTECT.REVISIONS",
    0x026D: "UNPROTECT.REVISIONS",
    0x0287: "OPTIONS.ME",
    0x028D: "WEB.PUBLISH",
    0x029B: "NEWWEBQUERY",
    0x02A1: "PIVOT.TABLE.CHART",
    0x02F1: "OPTIONS.SAVE",
    0x02F3: "OPTIONS.SPELL",
    0x0328: "HIDEALL.INKANNOTS"
}

const F_TAB_MAP = {
    0x0000: "COUNT",
    0x0001: "IF",
    0x0002: "ISNA",
    0x0003: "ISERROR",
    0x0004: "SUM",
    0x0005: "AVERAGE",
    0x0006: "MIN",
    0x0007: "MAX",
    0x0008: "ROW",
    0x0009: "COLUMN",
    0x000A: "NA",
    0x000B: "NPV",
    0x000C: "STDEV",
    0x000D: "DOLLAR",
    0x000E: "FIXED",
    0x000F: "SIN",
    0x0010: "COS",
    0x0011: "TAN",
    0x0012: "ATAN",
    0x0013: "PI",
    0x0014: "SQRT",
    0x0015: "EXP",
    0x0016: "LN",
    0x0017: "LOG10",
    0x0018: "ABS",
    0x0019: "INT",
    0x001A: "SIGN",
    0x001B: "ROUND",
    0x001C: "LOOKUP",
    0x001D: "INDEX",
    0x001E: "REPT",
    0x001F: "MID",
    0x0020: "LEN",
    0x0021: "VALUE",
    0x0022: "TRUE",
    0x0023: "FALSE",
    0x0024: "AND",
    0x0025: "OR",
    0x0026: "NOT",
    0x0027: "MOD",
    0x0028: "DCOUNT",
    0x0029: "DSUM",
    0x002A: "DAVERAGE",
    0x002B: "DMIN",
    0x002C: "DMAX",
    0x002D: "DSTDEV",
    0x002E: "VAR",
    0x002F: "DVAR",
    0x0030: "TEXT",
    0x0031: "LINEST",
    0x0032: "TREND",
    0x0033: "LOGEST",
    0x0034: "GROWTH",
    0x0035: "GOTO",
    0x0036: "HALT",
    0x0037: "RETURN",
    0x0038: "PV",
    0x0039: "FV",
    0x003A: "NPER",
    0x003B: "PMT",
    0x003C: "RATE",
    0x003D: "MIRR",
    0x003E: "IRR",
    0x003F: "RAND",
    0x0040: "MATCH",
    0x0041: "DATE",
    0x0042: "TIME",
    0x0043: "DAY",
    0x0044: "MONTH",
    0x0045: "YEAR",
    0x0046: "WEEKDAY",
    0x0047: "HOUR",
    0x0048: "MINUTE",
    0x0049: "SECOND",
    0x004A: "NOW",
    0x004B: "AREAS",
    0x004C: "ROWS",
    0x004D: "COLUMNS",
    0x004E: "OFFSET",
    0x004F: "ABSREF",
    0x0050: "RELREF",
    0x0051: "ARGUMENT",
    0x0052: "SEARCH",
    0x0053: "TRANSPOSE",
    0x0054: "ERROR",
    0x0055: "STEP",
    0x0056: "TYPE",
    0x0057: "ECHO",
    0x0058: "SET.NAME",
    0x0059: "CALLER",
    0x005A: "DEREF",
    0x005B: "WINDOWS",
    0x005C: "SERIES",
    0x005D: "DOCUMENTS",
    0x005E: "ACTIVE.CELL",
    0x005F: "SELECTION",
    0x0060: "RESULT",
    0x0061: "ATAN2",
    0x0062: "ASIN",
    0x0063: "ACOS",
    0x0064: "CHOOSE",
    0x0065: "HLOOKUP",
    0x0066: "VLOOKUP",
    0x0067: "LINKS",
    0x0068: "INPUT",
    0x0069: "ISREF",
    0x006A: "GET.FORMULA",
    0x006B: "GET.NAME",
    0x006C: "SET.VALUE",
    0x006D: "LOG",
    0x006E: "EXEC",
    0x006F: "CHAR",
    0x0070: "LOWER",
    0x0071: "UPPER",
    0x0072: "PROPER",
    0x0073: "LEFT",
    0x0074: "RIGHT",
    0x0075: "EXACT",
    0x0076: "TRIM",
    0x0077: "REPLACE",
    0x0078: "SUBSTITUTE",
    0x0079: "CODE",
    0x007A: "NAMES",
    0x007B: "DIRECTORY",
    0x007C: "FIND",
    0x007D: "CELL",
    0x007E: "ISERR",
    0x007F: "ISTEXT",
    0x0080: "ISNUMBER",
    0x0081: "ISBLANK",
    0x0082: "T",
    0x0083: "N",
    0x0084: "FOPEN",
    0x0085: "FCLOSE",
    0x0086: "FSIZE",
    0x0087: "FREADLN",
    0x0088: "FREAD",
    0x0089: "FWRITELN",
    0x008A: "FWRITE",
    0x008B: "FPOS",
    0x008C: "DATEVALUE",
    0x008D: "TIMEVALUE",
    0x008E: "SLN",
    0x008F: "SYD",
    0x0090: "DDB",
    0x0091: "GET.DEF",
    0x0092: "REFTEXT",
    0x0093: "TEXTREF",
    0x0094: "INDIRECT",
    0x0095: "REGISTER",
    0x0096: "CALL",
    0x0097: "ADD.BAR",
    0x0098: "ADD.MENU",
    0x0099: "ADD.COMMAND",
    0x009A: "ENABLE.COMMAND",
    0x009B: "CHECK.COMMAND",
    0x009C: "RENAME.COMMAND",
    0x009D: "SHOW.BAR",
    0x009E: "DELETE.MENU",
    0x009F: "DELETE.COMMAND",
    0x00A0: "GET.CHART.ITEM",
    0x00A1: "DIALOG.BOX",
    0x00A2: "CLEAN",
    0x00A3: "MDETERM",
    0x00A4: "MINVERSE",
    0x00A5: "MMULT",
    0x00A6: "FILES",
    0x00A7: "IPMT",
    0x00A8: "PPMT",
    0x00A9: "COUNTA",
    0x00AA: "CANCEL.KEY",
    0x00AB: "FOR",
    0x00AC: "WHILE",
    0x00AD: "BREAK",
    0x00AE: "NEXT",
    0x00AF: "INITIATE",
    0x00B0: "REQUEST",
    0x00B1: "POKE",
    0x00B2: "EXECUTE",
    0x00B3: "TERMINATE",
    0x00B4: "RESTART",
    0x00B5: "HELP",
    0x00B6: "GET.BAR",
    0x00B7: "PRODUCT",
    0x00B8: "FACT",
    0x00B9: "GET.CELL",
    0x00BA: "GET.WORKSPACE",
    0x00BB: "GET.WINDOW",
    0x00BC: "GET.DOCUMENT",
    0x00BD: "DPRODUCT",
    0x00BE: "ISNONTEXT",
    0x00BF: "GET.NOTE",
    0x00C0: "NOTE",
    0x00C1: "STDEVP",
    0x00C2: "VARP",
    0x00C3: "DSTDEVP",
    0x00C4: "DVARP",
    0x00C5: "TRUNC",
    0x00C6: "ISLOGICAL",
    0x00C7: "DCOUNTA",
    0x00C8: "DELETE.BAR",
    0x00C9: "UNREGISTER",
    0x00CC: "USDOLLAR",
    0x00CD: "FINDB",
    0x00CE: "SEARCHB",
    0x00CF: "REPLACEB",
    0x00D0: "LEFTB",
    0x00D1: "RIGHTB",
    0x00D2: "MIDB",
    0x00D3: "LENB",
    0x00D4: "ROUNDUP",
    0x00D5: "ROUNDDOWN",
    0x00D6: "ASC",
    0x00D7: "DBCS",
    0x00D8: "RANK",
    0x00DB: "ADDRESS",
    0x00DC: "DAYS360",
    0x00DD: "TODAY",
    0x00DE: "VDB",
    0x00DF: "ELSE",
    0x00E0: "ELSE.IF",
    0x00E1: "END.IF",
    0x00E2: "FOR.CELL",
    0x00E3: "MEDIAN",
    0x00E4: "SUMPRODUCT",
    0x00E5: "SINH",
    0x00E6: "COSH",
    0x00E7: "TANH",
    0x00E8: "ASINH",
    0x00E9: "ACOSH",
    0x00EA: "ATANH",
    0x00EB: "DGET",
    0x00EC: "CREATE.OBJECT",
    0x00ED: "VOLATILE",
    0x00EE: "LAST.ERROR",
    0x00EF: "CUSTOM.UNDO",
    0x00F0: "CUSTOM.REPEAT",
    0x00F1: "FORMULA.CONVERT",
    0x00F2: "GET.LINK.INFO",
    0x00F3: "TEXT.BOX",
    0x00F4: "INFO",
    0x00F5: "GROUP",
    0x00F6: "GET.OBJECT",
    0x00F7: "DB",
    0x00F8: "PAUSE",
    0x00FB: "RESUME",
    0x00FC: "FREQUENCY",
    0x00FD: "ADD.TOOLBAR",
    0x00FE: "DELETE.TOOLBAR",
    0x00FF: "User Defined Function",
    0x0100: "RESET.TOOLBAR",
    0x0101: "EVALUATE",
    0x0102: "GET.TOOLBAR",
    0x0103: "GET.TOOL",
    0x0104: "SPELLING.CHECK",
    0x0105: "ERROR.TYPE",
    0x0106: "APP.TITLE",
    0x0107: "WINDOW.TITLE",
    0x0108: "SAVE.TOOLBAR",
    0x0109: "ENABLE.TOOL",
    0x010A: "PRESS.TOOL",
    0x010B: "REGISTER.ID",
    0x010C: "GET.WORKBOOK",
    0x010D: "AVEDEV",
    0x010E: "BETADIST",
    0x010F: "GAMMALN",
    0x0110: "BETAINV",
    0x0111: "BINOMDIST",
    0x0112: "CHIDIST",
    0x0113: "CHIINV",
    0x0114: "COMBIN",
    0x0115: "CONFIDENCE",
    0x0116: "CRITBINOM",
    0x0117: "EVEN",
    0x0118: "EXPONDIST",
    0x0119: "FDIST",
    0x011A: "FINV",
    0x011B: "FISHER",
    0x011C: "FISHERINV",
    0x011D: "FLOOR",
    0x011E: "GAMMADIST",
    0x011F: "GAMMAINV",
    0x0120: "CEILING",
    0x0121: "HYPGEOMDIST",
    0x0122: "LOGNORMDIST",
    0x0123: "LOGINV",
    0x0124: "NEGBINOMDIST",
    0x0125: "NORMDIST",
    0x0126: "NORMSDIST",
    0x0127: "NORMINV",
    0x0128: "NORMSINV",
    0x0129: "STANDARDIZE",
    0x012A: "ODD",
    0x012B: "PERMUT",
    0x012C: "POISSON",
    0x012D: "TDIST",
    0x012E: "WEIBULL",
    0x012F: "SUMXMY2",
    0x0130: "SUMX2MY2",
    0x0131: "SUMX2PY2",
    0x0132: "CHITEST",
    0x0133: "CORREL",
    0x0134: "COVAR",
    0x0135: "FORECAST",
    0x0136: "FTEST",
    0x0137: "INTERCEPT",
    0x0138: "PEARSON",
    0x0139: "RSQ",
    0x013A: "STEYX",
    0x013B: "SLOPE",
    0x013C: "TTEST",
    0x013D: "PROB",
    0x013E: "DEVSQ",
    0x013F: "GEOMEAN",
    0x0140: "HARMEAN",
    0x0141: "SUMSQ",
    0x0142: "KURT",
    0x0143: "SKEW",
    0x0144: "ZTEST",
    0x0145: "LARGE",
    0x0146: "SMALL",
    0x0147: "QUARTILE",
    0x0148: "PERCENTILE",
    0x0149: "PERCENTRANK",
    0x014A: "MODE",
    0x014B: "TRIMMEAN",
    0x014C: "TINV",
    0x014E: "MOVIE.COMMAND",
    0x014F: "GET.MOVIE",
    0x0150: "CONCATENATE",
    0x0151: "POWER",
    0x0152: "PIVOT.ADD.DATA",
    0x0153: "GET.PIVOT.TABLE",
    0x0154: "GET.PIVOT.FIELD",
    0x0155: "GET.PIVOT.ITEM",
    0x0156: "RADIANS",
    0x0157: "DEGREES",
    0x0158: "SUBTOTAL",
    0x0159: "SUMIF",
    0x015A: "COUNTIF",
    0x015B: "COUNTBLANK",
    0x015C: "SCENARIO.GET",
    0x015D: "OPTIONS.LISTS.GET",
    0x015E: "ISPMT",
    0x015F: "DATEDIF",
    0x0160: "DATESTRING",
    0x0161: "NUMBERSTRING",
    0x0162: "ROMAN",
    0x0163: "OPEN.DIALOG",
    0x0164: "SAVE.DIALOG",
    0x0165: "VIEW.GET",
    0x0166: "GETPIVOTDATA",
    0x0167: "HYPERLINK",
    0x0168: "PHONETIC",
    0x0169: "AVERAGEA",
    0x016A: "MAXA",
    0x016B: "MINA",
    0x016C: "STDEVPA",
    0x016D: "VARPA",
    0x016E: "STDEVA",
    0x016F: "VARA",
    0x0170: "BAHTTEXT",
    0x0171: "THAIDAYOFWEEK",
    0x0172: "THAIDIGIT",
    0x0173: "THAIMONTHOFYEAR",
    0x0174: "THAINUMSOUND",
    0x0175: "THAINUMSTRING",
    0x0176: "THAISTRINGLENGTH",
    0x0177: "ISTHAIDIGIT",
    0x0178: "ROUNDBAHTDOWN",
    0x0179: "ROUNDBAHTUP",
    0x017A: "THAIYEAR",
    0x017B: "RTD",
}


const SHEET_STATE_MAP = {
    VISIBLE: 0x00,
    HIDDEN: 0x01,
    VERY_HIDDEN: 0x02
}
const SHEET_STATE = {
    0x00: 'Visible',
    0x01: 'Hidden',
    0x02: 'Very Hidden'
}
function get_sheet_state_descriptive(state) {
    return SHEET_STATE[state] || "UNKNOWN"
}
const SHEET_TYPE_MAP = {
    WORK_SHEET: 0x00,
    DIALOG_SHEET: 0x00,
    MACRO_SHEET: 0x01,
    CHART_SHEET: 0x02,
    VBA_MODULE: 0x06
}
const SHEET_TYPE = {
    0x00: 'Work sheet or Dialog sheet',
    0x01: 'Macro sheet',
    0x02: 'Chart sheet',
    0x06: 'VBA module'
}
function get_sheet_type_descriptive(type) {
    return SHEET_TYPE[type] || "UNKNOWN"
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
            value: ieee754.read(data, 0, true, 0, 8)
        }

    }

}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/6e5eed10-5b77-43d6-8dd0-37345f8654ad
 * @param {Buffer} buff 
 */
function parseColRelU(buff) {
    let data = buff.readUInt16LE(0)
    return {
        col: data & 0x3FFF,
        colRelative: (data >> 14) & 0x01,
        rowRelative: data >> 15
    }
}


/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/24bbaffd-ab88-489e-983f-3c400c8a8559
 * @param {Buffer} celltable 
 * @param {number} reference 
 */
function lookupRowBlock(celltable, reference) {
    let dbcell = celltable.slice(reference)
    let dbRtrw = dbcell.readUInt32LE(0);

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

module.exports.BiffDocument = BiffDocument