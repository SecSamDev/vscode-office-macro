const { PT_TOKEN_EXCEL_MAP: TOKEN_EXCEL_MAP, F_TAB_MAP, C_TAB_MAP, F_TAB_PARAMS_MAP, PT_TOKEN_EXCEL: TOKEN_EXCEL } = require('./biff-constants')
const ieee754 = require('./ieee754')

const TOKEN_CLASS_PRIMITIVE = 0;
const TOKEN_CLASS_OPERATOR = 1;
const TOKEN_CLASS_FUNCTION = 2;
const TOKEN_CLASS_REFERENCE = 3;
const TOKEN_CLASS_ERROR = 4;
const TOKEN_CLASS_DEFINED_NAME = 5;


const PT_TOKEN_EXCEL_CONSTRUCTOR = {
    'ptgExp': (data) => { return { type: 'ptgExp' } },
    'ptgTbl': (data) => { return { type: 'ptgTbl' } },
    'ptgAdd': (data) => {
        return {
            type: 'ptgAdd',
            cls: TOKEN_CLASS_OPERATOR,
            byte_length: 0,
            value: '+',
            js_v: '+'
        }
    },
    'ptgSub': (data) => { return { type: 'ptgSub', cls: TOKEN_CLASS_OPERATOR, value: '-', js_v: '-' } },
    'ptgMul': (data) => { return { type: 'ptgMul', cls: TOKEN_CLASS_OPERATOR, value: '*', js_v: '*' } },
    'ptgDiv': (data) => { return { type: 'ptgDiv', cls: TOKEN_CLASS_OPERATOR, value: '/', js_v: '/' } },
    'ptgPower': (data) => { return { type: 'ptgPower', cls: TOKEN_CLASS_OPERATOR, value: '^', js_v: '^' } },
    'ptgConcat': (data) => {
        return {
            type: 'ptgConcat',
            cls: TOKEN_CLASS_OPERATOR,
            byte_length: 0,
            value: '&',
            js_v: '+'
        }
    },
    'ptgLT': (data) => { return { type: 'ptgLT', cls: TOKEN_CLASS_OPERATOR, value: '<', js_v: '<' } },
    'ptgLE': (data) => { return { type: 'ptgLE', cls: TOKEN_CLASS_OPERATOR, value: '<=', js_v: '<=' } },
    'ptgEQ': (data) => { return { type: 'ptgEQ', cls: TOKEN_CLASS_OPERATOR, value: '==', js_v: '==' } },
    'ptgGE': (data) => { return { type: 'ptgGE', cls: TOKEN_CLASS_OPERATOR, value: '>=', js_v: '>=' } },
    'ptgGT': (data) => { return { type: 'ptgGT', cls: TOKEN_CLASS_OPERATOR, value: '>', js_v: '>' } },
    'ptgNE': (data) => { return { type: 'ptgNE', cls: TOKEN_CLASS_OPERATOR, value: '!=', js_v: '!=' } },
    'ptgIsect': (data) => { return { type: 'ptgIsect' } },
    'ptgUnion': (data) => { return { type: 'ptgUnion' } },
    'ptgRange': (data) => { return { type: 'ptgRange' } },
    'ptgUplus': (data) => { return { type: 'ptgUplus' } },
    'ptgUminus': (data) => { return { type: 'ptgUminus' } },
    'ptgPercent': (data) => { return { type: 'ptgPercent' } },
    'ptgParen': (data) => { return { type: 'ptgParen' } },
    'ptgMissArg': (data) => {
        let val = processShortXLUnicodeString(data)
        return {
            type: 'ptgMissArg',
            value: "MISSING",
            byte_length: 0,
            cls: TOKEN_CLASS_ERROR,
            js_v: "undefined"
        }
    },
    'ptgStr': (data) => {
        let val = processShortXLUnicodeString(data)
        return {
            type: 'ptgStr',
            value: val.value,
            js_v: `'${escape_function_string(val.value)}'`,//TODO: escape chars
            byte_length: val.byte_length,
            cls: TOKEN_CLASS_PRIMITIVE
        }
    },
    'ptgExtend': (data) => { return { type: 'ptgExtend' } },
    'ptgAttr': (data) => { return { type: 'ptgAttr' } },
    'ptgSheet': (data) => { return { type: 'ptgSheet' } },
    'ptgEndSheet': (data) => { return { type: 'ptgEndSheet' } },
    'ptgErr': (data) => { return { type: 'ptgErr' } },
    'ptgBool': (data) => { return { type: 'ptgBool' } },
    'ptgInt': (data, context) => {
        return {
            type: 'ptgInt',
            value: data.readUInt16LE(),
            js_v: data.readUInt16LE(),
            byte_length: 2,
            cls: TOKEN_CLASS_PRIMITIVE
        }
    },
    'ptgNum': (data) => {
        return { type: 'ptgNum' }
    },
    'ptgArray': (data, context) => {
        //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/edd64b46-0fa0-4ef0-b95b-fe2cd41c7f73
        let ptg_extra_cols = context.rgbe.readUInt8(0) + 1
        let ptg_extra_rows = context.rgbe.readUInt16LE(1)
        let arr = new Array(ptg_extra_cols * ptg_extra_rows)
        let offset = 3
        for (let i_c = 0; i_c < ptg_extra_cols; i_c++) {
            for (let i_r = 0; i_r < ptg_extra_rows; i_r++) {
                let ar_v = parse_SerAr(context.rgbe.slice(offset))
                arr[i_c * ptg_extra_cols + i_r] = ar_v.value
                offset += ar_v.byte_length
            }
        }
        //TODO: escape strings
        let js_v = arr.map((v) => typeof v === 'string' ? `'${escape_function_string(v)}'` : v);
        return {
            type: 'ptgArray',
            value: arr,
            js_v: `${js_v.join(',')}]`,
            cls: TOKEN_CLASS_PRIMITIVE,
            byte_length: 7
        }
    },
    'ptgFunc': (data) => {
        let vl = processPtgFuncFixed(data);
        return {
            type: 'ptgFunc',
            value: vl.f_name,
            n_params: vl.n_params,
            cls: TOKEN_CLASS_FUNCTION,
            byte_length: 4
        }
    },
    'ptgFuncVar': (data) => {
        let vl = processPtgFuncVar(data);
        return {
            type: 'ptgFuncVar',
            value: vl.f_name,
            n_params: vl.n_params,
            cls: TOKEN_CLASS_FUNCTION,
            byte_length: 4
        }
    },
    'ptgName': (data, context) => {
        let vl = data.readUInt32LE()
        return {
            type: 'ptgName',
            value: vl,
            cls: TOKEN_CLASS_DEFINED_NAME,
            js_v: `CELL_LIST.name_${vl}()`,
            byte_length: 4
        }
    },
    'ptgRef': (data, context) => {
        return {
            type: 'ptgRef',
            cls: TOKEN_CLASS_REFERENCE,
            row: data.readUInt16LE(),
            col: data.readUInt16LE(2),
            value: `R${data.readUInt16LE() + 1}C${data.readUInt16LE(2) + 1}`,
            js_v: `CELL_LIST.cell_${context.sheet}_${data.readUInt16LE()}_${data.readUInt16LE(2)}()`,
            byte_length: 4
        }
    },
    'ptgArea': (data) => { return { type: 'ptgArea' } },
    'ptgMemArea': (data) => { return { type: 'ptgMemArea' } },
    'ptgMemErr': (data) => { return { type: 'ptgMemErr' } },
    'ptgMemNoMem': (data) => { return { type: 'ptgMemNoMem' } },
    'ptgMemFunc': (data) => { return { type: 'ptgMemFunc' } },
    'ptgRefErr': (data) => {
        return {
            type: 'ptgRefErr',
            value: "ERROR",
            cls: TOKEN_CLASS_ERROR,
            js_v: 'undefined',
            byte_length: 4,
        }
    },
    'ptgAreaErr': (data) => {
        return {
            type: 'ptgAreaErr',
            value: "ERROR",
            cls: TOKEN_CLASS_ERROR,
            js_v: 'undefined',
            byte_length: 8,
        }
    },
    'ptgRefN': (data) => { return { type: 'ptgRefN' } },
    'ptgAreaN': (data) => { return { type: 'ptgAreaN' } },
    'ptgMemAreaN': (data) => { return { type: 'ptgMemAreaN' } },
    'ptgMemNoMemN': (data) => { return { type: 'ptgMemNoMemN' } },
    'ptgNameX': (data) => {
        let ixti = data.readUInt16LE(0)
        let nameindex = data.readUInt32LE(2)
        return {
            type: 'ptgNameX',
            value: nameindex,
            ixti: ixti,
            cls: TOKEN_CLASS_DEFINED_NAME,
            js_v: `CELL_LIST.name_${vl}()`,
            byte_length: 6
        }
    },
    'ptgRef3d': (data, context) => {
        //Local reference or to another cell
        let row = data.readUInt16LE(2);
        let col = data.readUInt16LE(4);
        let sheet = null
        let row_r = false
        let col_r = false
        if (context.named_formula) {
            //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/2db37ba7-32f3-4395-88fe-6646034a5358
            col_r = ((col >> 14) & 0x01) > 0
            row_r = ((col >> 15)) > 0

            //Todo from Unsigned to Signed...
            if (row_r) {
                row = data.readInt16LE(2);
                //To absolute
                row = context.row + row
            }
            if (col_r) {
                col = data.readInt16LE(4) & 0x7FFF;
                //To absolute
                col = context.col + col
            } else {
                col = col & 0x7FFF
            }
        } else {
            //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/f2395c33-34a4-4b07-85a9-9bb5f07848d9
        }
        return {
            type: 'ptgRef3d',
            cls: TOKEN_CLASS_REFERENCE,
            row,
            row_r,
            col_r,
            col,
            value: `R${row + 1}C${col + 1}`,
            js_v: `CELL_LIST.cell_${context.sheet}_${row}_${col}()`,
            byte_length: 6
        }
    },
    'ptgArea3d': (data) => { return { type: 'ptgArea3d' } },
    'ptgRefErr3d': (data) => { return { type: 'ptgRefErr3d' } },
    'ptgAreaErr3d': (data) => { return { type: 'ptgAreaErr3d' } },
    'ptgArrayV': (data, rgb_extra) => {
        //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/edd64b46-0fa0-4ef0-b95b-fe2cd41c7f73
        let ptg_extra_cols = rgb_extra.readUInt8(0) + 1
        let ptg_extra_rows = rgb_extra.readUInt16LE(1)
        let arr = new Array(ptg_extra_cols * ptg_extra_rows)
        let offset = 3
        for (let i_c = 0; i_c < ptg_extra_cols; i_c++) {
            for (let i_r = 0; i_r < ptg_extra_rows; i_r++) {
                let ar_v = parse_SerAr(rgb_extra.slice(offset))
                arr[i_c * ptg_extra_cols + i_r] = ar_v.value
                offset += ar_v.byte_length
            }
        }
        let js_v = arr.map((v) => typeof v === 'string' ? `'${escape_function_string(v)}'` : v);
        return {
            type: 'ptgArrayV',
            value: arr,
            cls: TOKEN_CLASS_PRIMITIVE,
            js_v: `[${js_v.join(",")}]`,
            byte_length: 7
        }
    },
    'ptgFuncV': (data) => {
        let vl = processPtgFuncFixed(data);
        return {
            type: 'ptgFuncV',
            value: vl.f_name,
            n_params: vl.n_params,
            cls: TOKEN_CLASS_FUNCTION,
            byte_length: 4
        }
    },
    'ptgFuncVarV': (data) => {
        let vl = processPtgFuncVar(data);
        return {
            type: 'ptgFuncVarV',
            value: vl.f_name,
            n_params: vl.n_params,
            cls: TOKEN_CLASS_FUNCTION,
            byte_length: 4
        }
    },
    'ptgNameV': (data) => {
        let vl = data.readUInt32LE()
        return {
            type: 'ptgNameV',
            value: vl,
            cls: TOKEN_CLASS_DEFINED_NAME,
            byte_length: 4
        }
    },
    'ptgRefV': (data, context) => {
        return {
            type: 'ptgRefV',
            cls: TOKEN_CLASS_REFERENCE,
            row: data.readUInt16LE(),
            col: data.readUInt16LE(2),
            value: `R${data.readUInt16LE() + 1}C${data.readUInt16LE(2) + 1}`,
            js_v: `CELL_LIST.cell_${context.sheet}_${data.readUInt16LE()}_${data.readUInt16LE(2)}()`,
            byte_length: 4
        }
    },
    'ptgAreaV': (data) => { return { type: 'ptgAreaV' } },
    'ptgMemAreaV': (data) => { return { type: 'ptgMemAreaV' } },
    'ptgMemErrV': (data) => { return { type: 'ptgMemErrV' } },
    'ptgMemNoMemV': (data) => { return { type: 'ptgMemNoMemV' } },
    'ptgMemFuncV': (data) => { return { type: 'ptgMemFuncV' } },
    'ptgRefErrV': (data) => {
        return {
            type: 'ptgRefErrV',
            value: "ERROR",
            cls: TOKEN_CLASS_ERROR,
            byte_length: 4,
        }
    },
    'ptgAreaErrV': (data) => {
        return {
            type: 'ptgAreaErrV',
            value: "ERROR",
            cls: TOKEN_CLASS_ERROR,
            byte_length: 8,
        }
    },
    'ptgRefNV': (data) => { return { type: 'ptgRefNV' } },
    'ptgAreaNV': (data) => { return { type: 'ptgAreaNV' } },
    'ptgMemAreaNV': (data) => { return { type: 'ptgMemAreaNV' } },
    'ptgMemNoMemNV': (data) => { return { type: 'ptgMemNoMemNV' } },
    'ptgFuncCEV': (data) => { return { type: 'ptgFuncCEV' } },
    'ptgNameXV': (data) => {
        let ixti = data.readUInt16LE(0)
        let nameindex = data.readUInt32LE(2)
        return {
            type: 'ptgNameXV',
            value: nameindex,
            ixti: ixti,
            cls: TOKEN_CLASS_DEFINED_NAME,
            byte_length: 6
        }
    },
    'ptgRef3dV': (data) => { return { type: 'ptgRef3dV' } },
    'ptgArea3dV': (data) => { return { type: 'ptgArea3dV' } },
    'ptgRefErr3dV': (data) => { return { type: 'ptgRefErr3dV' } },
    'ptgAreaErr3dV': (data) => { return { type: 'ptgAreaErr3dV' } },
    'ptgArrayA': (data, rgb_extra) => {
        //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/edd64b46-0fa0-4ef0-b95b-fe2cd41c7f73
        let ptg_extra_cols = rgb_extra.readUInt8(0) + 1
        let ptg_extra_rows = rgb_extra.readUInt16LE(1)
        let arr = new Array(ptg_extra_cols * ptg_extra_rows)
        let offset = 3
        for (let i_c = 0; i_c < ptg_extra_cols; i_c++) {
            for (let i_r = 0; i_r < ptg_extra_rows; i_r++) {
                let ar_v = parse_SerAr(rgb_extra.slice(offset))
                arr[i_c * ptg_extra_cols + i_r] = ar_v.value
                offset += ar_v.byte_length
            }
        }
        return {
            type: 'ptgArrayA',
            value: arr,
            cls: TOKEN_CLASS_PRIMITIVE,
            byte_length: 7
        }
    },
    'ptgFuncA': (data) => {
        let vl = processPtgFuncFixed(data);
        return {
            type: 'ptgFuncA',
            value: vl.f_name,
            n_params: vl.n_params,
            cls: TOKEN_CLASS_FUNCTION,
            byte_length: 4
        }
    },
    'ptgFuncVarA': (data) => {
        let vl = processPtgFuncVar(data);
        return {
            type: 'ptgFuncVarA',
            value: vl.f_name,
            n_params: vl.n_params,
            cls: TOKEN_CLASS_FUNCTION,
            byte_length: 4
        }
    },
    'ptgNameA': (data) => {
        let vl = data.readUInt32LE()
        return {
            type: 'ptgNameA',
            value: vl,
            cls: TOKEN_CLASS_DEFINED_NAME,
            byte_length: 4
        }
    },
    'ptgRefA': (data) => { return { type: 'ptgRefA' } },
    'ptgAreaA': (data) => { return { type: 'ptgAreaA' } },
    'ptgMemAreaA': (data) => { return { type: 'ptgMemAreaA' } },
    'ptgMemErrA': (data) => { return { type: 'ptgMemErrA' } },
    'ptgMemNoMemA': (data) => { return { type: 'ptgMemNoMemA' } },
    'ptgMemFuncA': (data) => { return { type: 'ptgMemFuncA' } },
    'ptgRefErrA': (data) => {
        return {
            type: 'ptgRefErrA',
            value: "ERROR",
            cls: TOKEN_CLASS_ERROR,
            byte_length: 4,
        }
    },
    'ptgAreaErrA': (data) => {
        return {
            type: 'ptgAreaErrA',
            value: "ERROR",
            cls: TOKEN_CLASS_ERROR,
            byte_length: 8,
        }
    },
    'ptgRefNA': (data) => { return { type: 'ptgRefNA' } },
    'ptgAreaNA': (data) => { return { type: 'ptgAreaNA' } },
    'ptgMemAreaNA': (data) => { return { type: 'ptgMemAreaNA' } },
    'ptgMemNoMemNA': (data) => { return { type: 'ptgMemNoMemNA' } },
    'ptgFuncCEA': (data) => { return { type: 'ptgFuncCEA' } },
    'ptgNameXA': (data) => {
        let vl = data.readUInt32LE()
        return {
            type: 'ptgNameXA',
            value: vl,
            cls: TOKEN_CLASS_DEFINED_NAME,
            byte_length: 4
        }
    },
    'ptgRef3dA': (data) => { return { type: 'ptgRef3dA' } },
    'ptgArea3dA': (data) => { return { type: 'ptgArea3dA' } },
    'ptgRefErr3dA': (data) => { return { type: 'ptgRefErr3dA' } },
    'ptgAreaErr3dA': (data) => { return { type: 'ptgAreaErr3dA' } },
}

const SER_AR_TYPES = {
    SER_NIL: 0,
    SER_NUM: 1,
    SER_STR: 2,
    SER_BOOL: 4,
    SER_ERR: 0x10
}

function parse_SerAr(blob) {
    let type = blob.readUInt8(0)
    if (type == SER_AR_TYPES.SER_NIL) {
        return {
            value: null,
            byte_length: 9
        }
    } else if (type == SER_AR_TYPES.SER_NUM) {
        return {
            value: ieee754.read(blob.slice(1), 0, true, 0, 8),
            byte_length: 9
        }
    } else if (type == SER_AR_TYPES.SER_STR) {
        let u_string = processShortXLUnicodeString(blob.slice(1))
        return {
            value: u_string.value,
            byte_length: u_string.byte_length + 1
        }
    } else if (type == SER_AR_TYPES.SER_BOOL) {
        return {
            value: blob.readUInt8(1) > 0,
            byte_length: 9
        }
    } else if (type == SER_AR_TYPES.SER_ERR) {
        return {
            value: blob.readUInt8(1),
            byte_length: 9
        }
    } else {
        throw new Error("Invalid SerAr value")
    }
}



/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/6cdf7d38-d08c-4e56-bd2f-6c82b8da752e
 * @param {Buffer} data 
 * @param {{row:0,col:0,rgcb : [], named_formula: false}} context
 */
function process_RGCE(data, context = { row: 0, col: 0, rgcb: [], named_formula: false }) {
    let offset = 0;
    let ptg_list = []
    while (offset < data.length) {
        let ptg_id = data.readUInt8(offset)
        offset += 1
        if (ptg_id === 0x18 || ptg_id === 0x19) {
            break;
        } else {
            let token_excel = TOKEN_EXCEL[ptg_id]
            if (token_excel == undefined) {
                throw new Error("Cannot process RGCE")
            }
            let ptg_value = PT_TOKEN_EXCEL_CONSTRUCTOR[token_excel](data.slice(offset), context)
            ptg_list.push(ptg_value)
            if (ptg_value.byte_length == undefined) {
                throw new Error("Cannot process RGCE")
            }
            offset += ptg_value.byte_length

        }
    }
    return ptg_list
}


/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/6cdf7d38-d08c-4e56-bd2f-6c82b8da752e
 * @param {[{type: string,
            value: any,
            cls: number,
            byte_length: number}]} ptg_list 
 */
function structure_RGCE(ptg_list = []) {
    let ptg_i = ptg_list
    if (ptg_i.length == 0) {
        return {
            value: "",
            elements: []
        }
    }
    if (ptg_i[0].cls == TOKEN_CLASS_FUNCTION) {
        let params = []
        let offst = 1;
        for (let i_param = 0; i_param < ptg_i[0].n_params[1]; i_param++) {
            let vl = structure_RGCE(ptg_i.slice(offst))
            offst += vl.elements.length
            params.push(vl)
        }
        return {
            value: `${ptg_i[0].value}(${params.map((vl) => vl.value).join(", ")})`,
            elements: ptg_list.slice(0, offst),
        }
    } else if (ptg_i[0].cls != TOKEN_CLASS_OPERATOR) {
        let last_class = ptg_i[0].cls
        let elements = [ptg_i[0]]
        for (let i = 1; i < ptg_i.length; i++) {
            if (last_class == ptg_i[i].cls) {
                return {
                    value: elements.map((val) => val.value).join(" "),
                    elements: elements,
                }
            }
            if (ptg_i[i].cls == TOKEN_CLASS_FUNCTION) {
                let fnc = structure_RGCE(ptg_i.slice(i))
                elements.push(fnc);
            } else {
                elements.push(ptg_i[i])
            }
            last_class = ptg_i[i].cls
        }
        return {
            value: elements.map((val) => val.value).join(" "),
            elements: elements,
        }
    }
    return {
        value: "",
        elements: [],
    }
}



/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/6cdf7d38-d08c-4e56-bd2f-6c82b8da752e
 * @param {[{type: string,
    value: any,
    cls: number,
    byte_length: number}]} ptg_list 
*/
function RGCE_to_javasript(ptg_list = []) {
    let ptg_i = ptg_list
    if (ptg_i.length == 0) {
        return {
            value: "null",
            elements: []
        }
    }
    if (ptg_i[0].cls == TOKEN_CLASS_FUNCTION) {
        if(ptg_i[0].value == "User Defined Function"){
            console.log("1")
        }
        let params = []
        let offst = 1;
        for (let i_param = 0; i_param < ptg_i[0].n_params[1]; i_param++) {
            let vl = RGCE_to_javasript(ptg_i.slice(offst))
            offst += vl.elements.length
            if (typeof vl.js_v === 'object') {
                vl.js_v = vl.js_v.js_v || vl.js_v.value
            }
            params.push(vl)
        }
        return `context.functions_${escape_function_name(ptg_i[0].value)}(${params.map((vl) => vl.js_v || vl.value).join(", ")})`
    } else if (ptg_i[0].cls != TOKEN_CLASS_OPERATOR) {
        let last_class = ptg_i[0].cls
        let elements = [ptg_i[0]]
        for (let i = 1; i < ptg_i.length; i++) {
            if (last_class == ptg_i[i].cls || arePrimitives(last_class, ptg_i[i].cls)) {
                return {
                    value: elements.map((val) => val.js_v || val.value).join(" "),
                    elements: elements,
                }
            }
            if (ptg_i[i].cls == TOKEN_CLASS_FUNCTION) {
                let fnc = RGCE_to_javasript(ptg_i.slice(i))
                elements.push(fnc);
                i += fnc.elements.length
                continue
            } else {
                elements.push(ptg_i[i])
            }
            last_class = ptg_i[i].cls
        }
        return {
            value: elements.map((val) => val.js_v || val.value).join(" "),
            elements
        }
    }
    return {
        value: "null",
        elements: []
    }
}

function arePrimitives(cls1, cls2){
    const primitives = [TOKEN_CLASS_DEFINED_NAME, TOKEN_CLASS_FUNCTION, TOKEN_CLASS_PRIMITIVE, TOKEN_CLASS_REFERENCE, TOKEN_CLASS_ERROR, TOKEN_CLASS_FUNCTION]
    return primitives.includes(cls1) && primitives.includes(cls2)
}



/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/05162858-0ca9-44cb-bb07-a720928f63f8
 * @param {Buffer} data 
 */
function processShortXLUnicodeString(data) {
    let cch = data.readUInt8(0)
    let fHighByte = data.readUInt8(1) >> 7
    let rgb = data.slice(2, 2 + cch * (fHighByte + 1))
    return {
        value: fHighByte ? rgb.toString('utf-16le') : rgb.toString('utf-8'),
        byte_length: (fHighByte + 1) * cch
    }
}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/36ca6de7-be16-48bc-aa5e-3eaf4942f671
 * @param {Buffer} data 
 */
function processXLUnicodeString(data) {
    let cch = data.readUInt16LE()
    let fHighByte = data.readUInt8(2) >> 7
    let rgb = data.slice(2, 2 + cch * (fHighByte + 1))
    return {
        value: fHighByte ? rgb.toString('utf-16le') : rgb.toString('utf-8'),
        byte_length: (fHighByte + 1) * cch
    }
}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/5d105171-6b73-4f40-a7cd-6bf2aae15e83
 * @param {Buffer} data 
 */
function processPtgFuncVar(data) {
    //The PtgFuncVar structure specifies a call to a function with a variable number of parameters as defined in function-call.
    let cparams = data.readUInt8(0)
    let tab = data.readUInt16LE(1)
    let fCeFunc = tab >> 15
    tab = tab & 0x7FFF
    return {
        f_name: fCeFunc ? C_TAB_MAP[tab] : F_TAB_MAP[tab],
        n_params: [cparams, cparams]
    }
}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/87ce512d-273a-4da0-a9f8-26cf1d93508d
 * @param {Buffer} data 
 */
function processPtgFuncFixed(data) {
    //The PtgFunc structure specifies a call to a function with a fixed number of parameters, as defined in function-call.
    let iftab = data.readUInt16LE(0)
    return {
        f_name: F_TAB_MAP[iftab],
        n_params: F_TAB_PARAMS_MAP[iftab]
    }
}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/e64abeee-2f3a-4004-b9e3-3d67e29d6066
 * @param {Buffer} data 
 */
function processXLUnicodeStringNoCch(data) {
    let fHighByte = data.readUInt8(0) >> 7
    let rgb = data.slice(1)
    return {
        value: fHighByte ? rgb.toString('utf-16le') : rgb.toString('utf-8'),
        byte_length: data.byteLength
    }
}

function escape_function_string(val = "") {
    return val.replace(/\\/g, "\\\\").replace(/'/g, "\\'")
}

function escape_function_name(val = "") {
    return val.replace(/ /g, "_")
}


module.exports.PT_TOKEN_EXCEL_CONSTRUCTOR = PT_TOKEN_EXCEL_CONSTRUCTOR
module.exports.process_RGCE = process_RGCE
module.exports.structure_RGCE = structure_RGCE
module.exports.RGCE_to_javasript = RGCE_to_javasript
module.exports.RCE_to_javasript = processShortXLUnicodeString
module.exports.processXLUnicodeStringNoCch = processXLUnicodeStringNoCch
module.exports.processXLUnicodeString = processXLUnicodeString