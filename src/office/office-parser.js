const fs = require('fs')
const { nextTick, off } = require('process')

class VbaDirStream {
    /**
     * 
     * @param {{information_record : VbaDirInformationRecord, references_record : VbaDirReferencesRecord}} data 
     */
    constructor(data) {
        this.information_record = data.information_record
        this.references_record = data.references_record
    }
    /**
     * Extracts the information associated with the VBA project from the dir stream
     * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/3d07f2c3-dee0-4ae3-b91f-3e32b789c534
     * @param {Buffer} buff 
     */
    static from_buffer(buff){
        let compressed_container = new CompressedContainer(buff)
        let decompressed = compressed_container.decompress()
        let data = {
            /**
             * @type {VbaDirInformationRecord}
             */
            information_record : null,
            /**
             * @type {VbaDirReferencesRecord}
             */
            references_record : null
        }

        data.information_record = new VbaDirInformationRecord(decompressed)
        data.references_record = new VbaDirReferencesRecord(decompressed.slice(data.information_record.size))
        return new VbaDirStream(data)
    }
}

const REFERENCE_CONTROL = 0x002F
const REFERENCE_ORIGINAL = 0x0033
const REFERENCE_REGISTERED = 0x000D
const REFERENCE_PROJECT = 0x000E

class VbaDirReferencesRecord {
    /**
     * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/1cf3c0b7-71ca-41cb-83f8-6360181512e2
     * @param {Buffer} buff 
     */
    constructor(buff) {

        this.references = []
        let offset = 0
        let in_references = true;
        while (in_references) {
            //Reference Named Record
            let id = buff.readUInt16LE(offset)
            offset += 2
            let name = ""
            let name_unicode = ""
            if (id == 0x0016) {
                //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/135dd749-c217-4d73-8549-b54e52e89945
                let size = buff.readUInt32LE(offset)
                if (size > buff.length) {
                    //Not documented behaviour
                    console.error(buff.slice(offset - 2, offset + 20))
                    console.error("Size of NamedReference malformed, position: " + offset)
                    break;
                }
                offset += 4
                name = buff.slice(offset, offset + size).toString()
                offset += size
                let reserved = buff.readUInt16LE(offset)//62
                offset += 2
                //Unicode
                size = buff.readUInt32LE(offset)
                if (size > buff.length) {
                    break;
                }
                offset += 4
                name_unicode = buff.slice(offset, offset + size).toString('utf16le')
                offset += size
                id = buff.readUInt16LE(offset)
                offset += 2
            }

            switch (id) {
                case REFERENCE_REGISTERED: {
                    //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/6c39388e-96f5-4b93-b90a-ae625a063fcf
                    let reference = VbaDirReferencesRecord.readReferenceRegistered(buff.slice(offset))

                    this.references.push({
                        name: name,
                        name_unicode: name_unicode,
                        type: id,
                        libid: reference

                    })
                    offset += 14 + reference.size_libid//6  reserved
                    break;
                }
                case REFERENCE_CONTROL: {
                    //TODO:
                    let size_twiddled = buff.readUInt32LE(offset + 4)
                    let libid_twiddled = buff.slice(offset + 8, offset + 8 + size_twiddled).toString()
                    offset += 14 + size_twiddled
                    let name_record_extended = {}
                    if (buff.readUInt16LE(offset) == 0x16) {
                        name_record_extended = VbaDirReferencesRecord.readNameRecord(buff.slice(offset))
                        offset += name_record_extended.offset
                    }
                    let size_libid_ext = buff.readUInt16LE(offset + 6)
                    let ref_libid = buff.slice(offset + 10, offset + 10 + size_libid_ext).toString()
                    let libid_ext = VbaDirReferencesRecord.process_libid_string(ref_libid)
                    //let id = buff.readUInt16LE(offset)
                    offset += 36 + size_libid_ext
                    this.references.push({
                        type: id,
                        libid_ext,
                        libid_twiddled,
                        name_record_extended
                    })
                    break;
                }
                case REFERENCE_PROJECT: {
                    let ref_size = buff.readUInt32LE(offset)
                    let ref_size_libid_abs = buff.readUInt32LE(offset + 4)
                    let ref_libid_abs = buff.slice(offset + 8, offset + 8 + ref_size_libid_abs).toString()
                    let ref_size_libid_rel = buff.readUInt32LE(offset + 8 + ref_size_libid_abs)
                    offset += 12 + ref_size_libid_abs
                    let ref_libid_rel = buff.slice(offset, offset + ref_size_libid_rel).toString()
                    offset += ref_size_libid_rel
                    let major_version = buff.readUInt32LE(offset)
                    let minor_version = buff.readUInt16LE(offset + 4)
                    offset += 6
                    this.references.push({
                        type: id,
                        libid: {
                            major_version,
                            minor_version,
                            path_abs: ref_libid_abs,
                            path_rel: ref_libid_rel
                        }
                    })
                    break;
                }
                case REFERENCE_ORIGINAL: {
                    let ref_size_libid = buff.readUInt32LE(offset)
                    let ref_libid = buff.slice(offset + 4, offset + 4 + ref_size_libid).toString()
                    let process_libid = VbaDirReferencesRecord.process_libid_string(ref_libid)
                    process_libid.size_libid = ref_size_libid
                    this.references.push({
                        type: id,
                        libid: process_libid
                    })
                    offset += 4 + ref_size_libid
                    break;
                }
                default: {
                    in_references = false;
                    break;
                }
            }
        }
        offset -= 2
        //PROJECTMODULES
        let id = buff.readUInt16LE(offset)
        checkMalformed(id, 0x0F)
        /**
         * @type {{
            name: string,
            stream: string,
            source_offset: number
        }[]}
         */
        this.modules = []

        offset += 6
        let count = buff.readUInt16LE(offset)
        offset += 10
        for (let i = 0; i < count; i++) {
            try {
                let id = buff.readUInt16LE(offset);//0x19
                checkMalformed(id, 0x19)
                let size = buff.readUInt32LE(offset + 2);
                let mod_name = buff.slice(offset + 6, offset + 6 + size).toString()

                offset += 8 + size
                //MODULENAMEUNICODE Record
                //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/b5bd9112-9e13-40c5-8f71-05dab0fefb22
                id = buff.readUInt16LE(offset - 2)
                checkMalformed(id, 0x0047)
                size = buff.readUInt32LE(offset);

                let mod_name_unicode = buff.slice(offset + 4, offset + 4 + size).toString('utf16le')
                offset += 4 + size

                //MODULESTREAMNAME Record
                id = buff.readUInt16LE(offset)
                checkMalformed(id, 0x001A)
                size = buff.readUInt32LE(offset + 2)

                let stream_name = buff.slice(offset + 6, offset + 6 + size).toString()
                offset += 6 + size
                let reserved = buff.readUInt16LE(offset)
                checkMalformed(reserved, 0x32)
                size = buff.readUInt32LE(offset + 2)

                let stream_name_unicode = buff.slice(offset + 6, offset + 6 + size).toString('utf16le')
                offset += 6 + size

                //DOCSTRING Record
                id = buff.readUInt16LE(offset)
                checkMalformed(id, 0x001C)
                size = buff.readUInt32LE(offset + 2)

                let doc_string = buff.slice(offset + 6, offset + 6 + size).toString()
                offset += 6 + size
                reserved = buff.readUInt16LE(offset)
                checkMalformed(reserved, 0x0048)
                size = buff.readUInt32LE(offset + 2)

                let doc_string_unicode = buff.slice(offset + 6, offset + 6 + size).toString()
                offset += 6 + size

                //ModuleOffset Record
                id = buff.readUInt16LE(offset)
                checkMalformed(id, 0x0031)
                let source_offset = buff.readUInt32LE(offset + 6)
                this.modules.push({
                    name: mod_name,
                    stream: stream_name,
                    source_offset: source_offset
                })
                offset = buff.indexOf(Uint8Array.from([0x002B]), offset + 6) + 6

            } catch (err) {
                console.error(err)
            }

        }
    }
    static readNameRecord(buff) {
        let offset = 0
        let size = buff.readUInt32LE(offset + 2)
        offset += 6
        let name = buff.slice(offset, offset + size).toString()
        offset += size
        let reserved = buff.readUInt16LE(offset)//62
        offset += 2
        //Unicode
        size = buff.readUInt32LE(offset)
        offset += 4
        let name_unicode = buff.slice(offset, offset + size).toString('utf16le')
        offset += size
        return {
            name: name,
            name_unicode: name_unicode,
            offset: offset
        }
    }
    static readReferenceRegistered(buff) {
        let ref_size_libid = buff.readUInt32LE(4)

        //let size = buff.readUInt32LE(0)
        let libid = buff.slice(8, 6 + ref_size_libid).toString()
        let toReturn = VbaDirReferencesRecord.process_libid_string(libid)
        toReturn.size_libid = ref_size_libid
        return toReturn

    }
    static process_libid_string(libid) {
        let regexLibidReference = /\*\\(?<libid_reference_kind>[GH])\{(?<libid_guid>[^\}]+)\}#(?<libid_major>[^\.]+)\.(?<libid_minor>[^#]+)#(?<libid_lcid>[^#]+)#(?<libid_path>[^#]+)#(?<libid_reg_name>.+)/.exec(libid)
        if (regexLibidReference && regexLibidReference.groups["libid_guid"]) {
            return {
                reference_kind: regexLibidReference.groups["libid_reference_kind"],
                guid: regexLibidReference.groups["libid_guid"],
                major_version: regexLibidReference.groups["libid_major"],
                minor_version: regexLibidReference.groups["libid_minor"],
                lcid: regexLibidReference.groups["libid_lcid"],
                path: regexLibidReference.groups["libid_path"],
                reg_name: regexLibidReference.groups["libid_reg_name"],
                libid: libid
            }
        } else {
            return {
                libid: libid
            }
        }
    }

}
function checkMalformed(code, expected) {
    if (code != expected) {
        throw new Error("Malformed code. Expected: " + expected + " get " + code)
    }
}

class VbaDirInformationRecord {
    /**
     * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/5abef063-3661-46dd-ba80-8cb507afdb1d
     * @param {Buffer} buff 
     */
    constructor(buff) {
        this.size = 0
        this.buffer = buff
        this.sys_kind_record = {
            id: buff.readUInt16LE(0),
            size: buff.readUInt32LE(2),
            sysKind: buff.readUInt16LE(6),
        }
        if (this.sys_kind_record.id != 0x0001)
            console.error('PROJECTSYSKIND ID must be 0x0001')
        if (this.sys_kind_record.size != 0x00000004)
            console.error('PROJECTSYSKIND size must be 0x00000004')
        if (this.sys_kind_record.sysKind > 3)
            console.error('PROJECTSYSKIND SysKind must be 0x00000000, 0x00000001, 0x00000002, 0x00000003')
        buff = buff.slice(10)
        this.lcid_record = {
            id: buff.readUInt16LE(0),
            size: buff.readUInt32LE(2),
            lcid: buff.readUInt32LE(6),
        }
        if (this.lcid_record.id != 0x0002)
            console.error('PROJECTLCID ID must be 0x0002 <> ' + this.lcid_record.id)
        if (this.lcid_record.size != 0x00000004)
            console.error('PROJECTLCID size must be 0x00000004 <> ' + this.lcid_record.size)
        if (this.lcid_record.lcid != 0x00000409)
            console.error('PROJECTLCID lcid must be 0x00000409 <> ' + this.lcid_record.lcid)

        buff = buff.slice(10)
        this.lcid_invoke_record = {
            id: buff.readUInt16LE(0),// 0X0014 = 20
            size: buff.readUInt32LE(2),// 0X00000004
            lcid_invoke: buff.readUInt32LE(6), //0X00000409
        }
        if (this.lcid_invoke_record.id != 0x0014)
            console.error('PROJECTLCID ID must be 0x0014 <> ' + this.lcid_invoke_record.id)
        if (this.lcid_invoke_record.size != 0x0004)
            console.error('PROJECTLCID size must be 0x0004 <> ' + this.lcid_invoke_record.size)
        if (this.lcid_invoke_record.lcid_invoke != 0x00000409)
            console.error('PROJECTLCID LcidInvoke must be 0x00000409 <> ' + this.lcid_invoke_record.lcid_invoke)

        buff = buff.slice(10)
        this.project_code_page_record = {
            id: buff.readUInt16LE(0),// 0x0003
            size: buff.readUInt16LE(2) >> 8,//0x00000002
            code_page: buff.readUInt16LE(6),
        }

        buff = buff.slice(8)// ???
        this.project_name_record = {
            id: buff.readUInt16LE(0),//0x0004
            size: buff.readUInt32LE(2),// 1 <= X <= 128
            name: "",
        }
        this.project_name_record.name = buff.slice(6, this.project_name_record.size + 6).toString()
        //this.size += 44 + this.project_name_record.size
        //PROJECTDOCSTRING Record
        buff = buff.slice(this.project_name_record.size + 6)
        this.project_doc_string = {
            id: buff.readUInt16LE(0),// 0X0005
            size: buff.readUInt32LE(2),// <= 2000
            doc_string: "",
            size_unicode: 0,
            doc_string_unicode: ""
        }
        let offset = 8 + this.project_doc_string.size // 2Bytes reserved
        //TODO: MCBS string
        this.project_doc_string.doc_string = this.project_doc_string.size > 0 ? buff.slice(6, 6 + this.project_doc_string.size).toString() : ""
        this.project_doc_string.size_unicode = buff.readUInt32LE(offset)
        offset += 4
        this.project_doc_string.doc_string_unicode = this.project_doc_string.size_unicode > 0 ? buff.slice(offset, offset + this.project_doc_string.size_unicode).toString() : ""
        offset += this.project_doc_string.size_unicode

        //this.size += 12 + this.project_doc_string.size + this.project_doc_string.size_unicode

        //PROJECTHELPFILEPATH Record
        buff = buff.slice(offset)
        this.project_help_file_path = {
            id: buff.readUInt16LE(0),//0x0006
            size_1: buff.readUInt32LE(2),
            file_1: "",
            size_2: 0,
            file_2: ""
        }
        offset = 6
        this.project_help_file_path.file_1 = this.project_help_file_path.size_1 > 0 ? buff.slice(offset, offset + this.project_help_file_path.size_1).toString() : ""
        offset += this.project_help_file_path.size_1 + 2//2 reserved
        this.project_help_file_path.size_2 = buff.readUInt32LE(offset)
        offset += 4
        this.project_help_file_path.file_2 = this.project_help_file_path.size_2 > 0 ? buff.slice(offset, offset + this.project_help_file_path.size_2).toString() : ""
        offset += this.project_help_file_path.size_2
        //this.size += 10 + this.project_help_file_path.size_2 + this.project_help_file_path.size_1
        //PROJECTHELPCONTEXT Record
        buff = buff.slice(offset)
        this.project_help_context = {
            id: buff.readUInt16LE(0),//0x0007
            size: buff.readUInt32LE(2),//0x00000004
            context: buff.readUInt32LE(6)
        }

        //PROJECTLIBFLAGS Record
        buff = buff.slice(10)
        this.project_lib_flags = {
            id: buff.readUInt16LE(0),//0x0008
            size: buff.readUInt32LE(2),//0x00000004
            flags: buff.readUInt32LE(6)//0x00000000
        }

        //PROJECTVERSION Record
        buff = buff.slice(10)
        this.project_lib_flags = {
            id: buff.readUInt16LE(0),//0x0009
            reserved: buff.readUInt32LE(2),//0x00000004
            major: buff.readUInt32LE(6),//Major version VBA
            minor: buff.readUInt16LE(10)//Minor version VBA
        }
        //PROJECTCONSTANTS Record
        buff = buff.slice(12)
        this.project_constants = {
            id: buff.readUInt16LE(0),//0x000C
            size: buff.readUInt32LE(2),//<=1015
            constants: [],
            size_unicode: 0,
            constants_unicode: []
        }

        this.constants = this.project_constants.size > 0 ? buff.slice(6, 6 + this.project_constants.size) : []
        offset = 8 + this.project_constants.size//2 reserved
        this.project_constants.size_unicode = buff.readUInt32LE(offset)
        offset += 4
        this.project_constants.constants_unicode = this.project_constants.size_unicode > 0 ? buff.slice(offset, offset + this.project_constants.size_unicode) : []
        this.size = this.buffer.length - (buff.length - (offset + this.project_constants.size_unicode))
        delete this.buffer
        //this.size += 44 + this.project_constants.size + this.project_constants.size_unicode
    }
}

class CompressedContainer {
    /**
     * Extracts a CompressedContainer from a buffer
     * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/4742b896-b32b-4eb0-8372-fbf01e3c65fd
     * @param {Buffer} buff 
     */
    constructor(buff) {
        this.buffer = buff
        /**
         * @type {CompressedChunk[]}
         */
        this.compressed_chunks = []
        if (buff[0] != 0x01)
            console.error('CompressedContainer must start with 0x01')
        this.state_vars = {
            compressed_record_end: buff.byteLength,
            compressed_current: 1,
            compressed_chunk_start: 1,
            decompressed_current: 0,
            decompressed_buffer_end: 0,
            decompressed_chunk_start: 0
        }
        this._extractCompressedChunks()

    }

    _extractCompressedChunks() {
        let decompresed = []
        while (this.state_vars.compressed_current < this.state_vars.compressed_record_end) {
            try {
                this.state_vars.compressed_chunk_start = this.state_vars.compressed_current
                let compressed_chunk = new CompressedChunk(this.buffer, this.state_vars)
                decompresed.push(compressed_chunk.decompress())

            } catch (err) {
                break;
            }

        }
        this.decompressed_data = Buffer.concat(decompresed)
    }
    decompress() {
        return this.decompressed_data
    }
}
const COMPRESS_CHUNK_FLAG_LESS_4096 = 1
const COMPRESS_CHUNK_FLAG_MUST_4096 = 0
//001110110011
//10110101 10110011
class CompressedChunk {
    /**
     * Extracts a CompressedChunk from a CompressedContainer
     * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ec1bc788-27de-47d9-8db4-6be2b5cff52b
     * @param {Buffer} buff 
     */
    constructor(buff, state_vars) {
        this.data = buff
        this.state_vars = state_vars
        this.chunk_size = (buff.readUInt16LE(this.state_vars.compressed_chunk_start) & 0x0FFF) + 3
        this.chunk_flag = buff[this.state_vars.compressed_chunk_start + 1] >> 4

        if ((this.chunk_flag & 0x07) != 3) {
            console.error('CompressedChunkSignature must be 0b011: ' + (this.chunk_flag & 0x07))
        }
        if (this.chunk_flag === COMPRESS_CHUNK_FLAG_LESS_4096 && this.chunk_size > 4096) {
            console.error('Size does not follow FLAG: must be less or equals than 4096')
        }
        if (this.chunk_flag === COMPRESS_CHUNK_FLAG_MUST_4096 && this.chunk_size != 4096) {
            console.error('Size does not follow FLAG: must be 4096')
        }
        this.compressed = ((this.chunk_flag >>> 3) == 1) ? true : false
        if (buff.length < (this.chunk_size - 1)) {
            throw new Error('No valid CompressedChunk. Malformed size: ' + this.chunk_size + " -> " + buff.length)
        }
        this.state_vars.decompressed_chunk_start = this.state_vars.decompressed_current
        this.state_vars.compressed_end = Math.min(this.state_vars.compressed_record_end, this.state_vars.compressed_chunk_start + this.chunk_size)
        this.state_vars.compressed_current = this.state_vars.compressed_chunk_start + 2

    }

    /**
     * Decompress this.data if its compress.
     * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/de5c80af-ac75-42ab-a0d2-a21173f49c9a
     */
    decompress() {
        //Decompressing a CompressedChunk
        this.decompressed_data = Buffer.alloc(0)
        if (this.compressed) {
            //Decompressing a token sequence
            while (this.state_vars.compressed_current < this.state_vars.compressed_end) {
                //Array of TokenSequence (0<= tokens < 8)
                let flag_byte = this.data[this.state_vars.compressed_current];
                this.state_vars.compressed_current += 1
                for (let i = 0; i < 8; i++) {
                    if (this.state_vars.compressed_current >= this.state_vars.compressed_end) {
                        return this.decompressed_data
                    }
                    if (((flag_byte >> i) & 0x1) == 0) {
                        //LiteralToken
                        this.decompressed_data = Buffer.concat([this.decompressed_data, this.data.slice(this.state_vars.compressed_current, this.state_vars.compressed_current + 1)])
                        this.state_vars.decompressed_current += 1
                        this.state_vars.compressed_current += 1
                    } else {
                        //CopyToken https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/c821c6f8-ec48-40a9-88fa-8c2b89589aec
                        let token = this.data.readUInt16LE(this.state_vars.compressed_current);
                        //27136,
                        let copy_token_masks = CopyTokenHelp(this.state_vars.decompressed_current, this.state_vars.decompressed_chunk_start)
                        let length_token = (token & copy_token_masks.length_mask) + 3
                        let temp1 = token & copy_token_masks.offset_mask
                        let temp2 = 16 - copy_token_masks.bit_count
                        let offset = (temp1 >> temp2) + 1
                        let copy_source = this.state_vars.decompressed_current - offset
                        for (let pos = copy_source; pos < (copy_source + length_token); pos++) {
                            this.decompressed_data = Buffer.concat([this.decompressed_data, this.decompressed_data.slice(pos, pos + 1)])
                        }
                        this.state_vars.decompressed_current += length_token
                        this.state_vars.compressed_current += 2
                    }
                }
            }

        } else {
            this.decompressed_data = this.data.slice(0, 4096)
            this.state_vars.compressed_current += 4096
        }
        return this.decompressed_data
    }
}

function CopyTokenHelp(decompressed_current, decompressed_chunk_start) {
    let difference = decompressed_current - decompressed_chunk_start
    let bit_count = Math.ceil(Math.log2(difference));
    bit_count = Math.max(bit_count, 4)
    let length_mask = 0xFFFF >> bit_count
    let offset_mask = (~length_mask)
    let maximum_length = length_mask + 3
    return {
        length_mask,
        offset_mask,
        bit_count,
        maximum_length
    }

}

const GUID_MAP_REFERENCES = {
    "00020430-0000-0000-C000-000000000046" : "",
}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ef7087ac-3974-4452-aab2-7dba2214d239
 */
class Vba_VBAProjectStream {
    
    /**
     * 
     * @param {{performance_cache : Buffer, vba_version : number, little_endian : boolean}} obj 
     */
    constructor(obj) {
        this.performance_cache = obj.performance_cache
        this.vba_version = obj.vba_version
        this.little_endian = obj.little_endian
        this.modules = []
    }
    /**
     * Reads the _VBA_PROJECT stream
     * @param {Buffer} buff 
     */
    static from_buffer(buff){
        //Reserved1
        if (buff.readUInt16LE(0) != 0x61CC)
            console.error("Must start with CC61")
        let vba_version = buff.readUInt16LE(2)
        //Reserved2
        if (buff.readUInt8(4) != 0x0)
            console.error("Reserved2 must be 0x0")
        //Reserved3 = 2 bytes undefined -> Sets the endianess
        //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/43531f35-2801-4cac-b6da-88dc975056da
        let little_endian = buff.readUInt16LE(5) != 0x00;
        let performance_cache = buff.slice(7)
        //Can be interesting for analysis
        return new Vba_VBAProjectStream({
            vba_version,
            performance_cache,
            little_endian
        })
    }

    process_performance_cache(){
        //PCODE
        if(this.little_endian){
            this._process_performance_cache_le();
        }else{

        }
        
    }
    /**
     * Process PerformanceCache for LittleEndian
     * @private
     */
    _process_performance_cache_le(){
        let prf_c = this.performance_cache
        let numRefs = prf_c.readUInt16LE(23)
        let offset = 27
        let hasUnicodeRef = (this.vba_version >= 0x5B) && (![0x60, 0x62, 0x63].includes(this.vba_version)) || (this.vba_version == 0x4E)
        let hasUnicodeName = (this.vba_version >= 0x59) && (![0x60, 0x62, 0x63].includes(this.vba_version)) || (this.vba_version == 0x4E)
        let hasNonUnicodeName = (this.vba_version <= 0x59) && (this.vba_version != 0x4E) || (this.vba_version > 0x6B && this.vba_version < 0x5F)

        //PROJECTLCID lcid must be 0x00000409
        //https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/1136037b-5e9e-4e2d-81f8-615ace60be9d
        let references = []
        //Definicion de modulos
        for(let i_ref = 0; i_ref < numRefs; i_ref++){
            let refLength = prf_c.readUInt16LE(offset)
            if(refLength < 40){
                let refString = prf_c.slice(offset + 2, offset + 2 + refLength).toString('utf16le')
                if(refString === '*\\CNormal'){
                    offset += refLength*2 + 16
                    i_ref--;
                }else{
                    offset += refLength + 2
                }
                
                continue;
            }
            if(hasUnicodeRef){
                //Similar as https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/3737ef6e-d819-4186-a5f2-6e258ddf66a5
                let refString =  prf_c.slice(offset + 2, offset + 2 + refLength).toString('utf16le')
                offset += 2 + refLength
                if(prf_c.readUInt32LE(offset + 6) == 0x00){
                    offset += 12
                }else{
                    i_ref--;
                }
                references.push(refString)
                continue;
            }
        }
        let pos_mod = prf_c.indexOf(Uint8Array.from([0xFF,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00]),offset)
        if(pos_mod < 0){
            throw new Error("Cannot parse pcode")
        }
        offset = pos_mod + 23
        
        let numModules = prf_c.readUInt16LE(offset);
        offset += 2
        for(let i_mod = 0; i_mod < numModules; i_mod++){
            let size_name = prf_c.readUInt16LE(offset);
            let mod_name = prf_c.slice(offset + 2,offset + 2 + size_name).toString('utf16le');
            offset += 2 + size_name
            let ref_size = prf_c.readUInt16LE(offset);
            let ref_name = prf_c.slice(offset + 2,offset + 2 + ref_size).toString('utf16le');
            offset += 6 + ref_size
            size_name = prf_c.readUInt16LE(offset);
            let name2 = prf_c.slice(offset + 2,offset + 2 + size_name).toString('utf16le');
            offset += 23 + size_name
            console.log(mod_name)
            this.modules.push({
                name: mod_name,
                name2,
                ref : ref_name
            })
        }
        //Skip the module descriptors
        //Module
        //Module size
        //module string
        //26 bytes basura
        //module size 2
        //module string

    }
}

class VbaProjectStream {
    /**
     * Create a Project from data
     * @param {{id,document,package,base_class,modules : [],exe_name32, name, help_context_id,version_compatible_32,cmg,dpb,GC,host_extender_info,workspace}} data 
     */
    constructor(data) {
        this.id = data.id;
        this.document = data.document;
        this.package = data.package;
        this.base_class = data.base_class;
        this.modules = data.modules;
        this.exe_name32 =data.exe_name32;
        this.name = data.name;
        this.help_context_id = data.help_context_id;
        this.version_compatible_32 = data.version_compatible_32;
        this.cmg = data.cmg;
        this.dpb = data.dpb;
        this.GC = data.GC;
        this.host_extender_info = data.host_extender_info;
        this.workspace = data.workspace;

    }
    static from_buffer(data){
        if (data instanceof Buffer) {
            data = data.toString()
        }
        if (typeof data != 'string') {
            throw new Error("VbaProjectStream needs a string")
        }
        let retData = {}

        let id_match = /^ID="{([A-Za-z0-9\\-]+)}"/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.id = id_match[1].trim()
        }
        id_match = /^Document=([^/]+)/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.document = id_match[1].trim()
        }
        id_match = /^Package={([A-Za-z0-9\\-]+)}/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.package = id_match[1].trim()
        }
        id_match = /^BaseClass=([^\n]+)/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.base_class = id_match[1].trim()
        }
        let data_searched = data;
        retData.modules = []
        let moduleRegex = /^Module=([^\n]+)/m
        while ((id_match = moduleRegex.exec(data_searched))) {
            retData.modules.push(id_match[1].trim())
            data_searched = data_searched.slice(id_match.index + 1)
        }
        id_match = /^ExeName32="([^"]+)"/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.exe_name32 = id_match[1].trim()
        }
        id_match = /^Name="([^"]+)"/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.name = id_match[1].trim()
        }
        id_match = /^HelpContextID="([^"]+)"/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.help_context_id = id_match[1].trim()
        }
        id_match = /^VersionCompatible32="([^"]+)"/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.version_compatible_32 = id_match[1].trim()
        }
        id_match = /^CMG="([^"]+)"/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.cmg = id_match[1].trim()
        }
        id_match = /^DPB="([^"]+)"/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.dpb = id_match[1].trim()
        }
        id_match = /^GC="([^"]+)"/m.exec(data)
        if (id_match && id_match.length === 2) {
            retData.GC = id_match[1].trim()
        }
        id_match = /^\[Host Extender Info\]/m.exec(data)

        if (id_match) {
            retData.host_extender_info = {}
            data_searched = data.slice(id_match.index + id_match[0].length + 2).split("\n")
            for (let i = 0; i < data_searched.length; i++) {
                if (data_searched[i].length < 5) {
                    break;
                }
                id_match = /^([^=]+)=([^\n]+)/.exec(data_searched[i])
                if (!id_match)
                    break;
                retData.host_extender_info[id_match[1]] = id_match[2]
            }
        }
        id_match = /^\[Workspace\]/m.exec(data)

        if (id_match) {
            retData.workspace = {}
            data_searched = data.slice(id_match.index + id_match[0].length + 2).split("\n")
            for (let i = 0; i < data_searched.length; i++) {
                if (data_searched[i].length < 5) {
                    break;
                }
                id_match = /^([^=]+)=([^\n]+)/.exec(data_searched[i])
                if (!id_match)
                    break;
                retData.workspace[id_match[1]] = id_match[2]
            }
        }
        return new VbaProjectStream(retData)
    }
}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/c66b58a6-f8ba-4141-9382-0612abce9926
 */
class VbaModule {
    /**
     * 
     * @param {source_code : string, performance_cache : Buffer} data 
     */
    constructor(data) {
        this.source_code = data.source_code;
        this.performance_cache = data.performance_cache;

    }
    static from_buffer(buffer, module_offset = 0){
        let compressed_container = new CompressedContainer(buffer.slice(module_offset))
        let source_code = compressed_container.decompress().toString()
        let data = {
            performance_cache : buffer.slice(0,module_offset),
            source_code : source_code
        }
        return new VbaModule(data)
    }

    process_performance_cache(){
        //z
        let magic = this.performance_cache.readUInt16LE(0)
        checkMalformed(magic,0x61CC)
        let version = this.performance_cache.readUInt16LE(2)
        let unicodeRef = (version >= 0x5B) && ([0x60,0x62,0x63].includes(version)) || version == 0x4E
    }

}
/**
 * 
 * @param {Buffer} buff 
 */
function localizeCompressContainer(buff, pos = 0) {
    let lastPos = pos - 1
    while (pos >= 0) {
        pos = buff.indexOf(1, lastPos + 1)
        if (((buff[pos + 2] >> 4) & 0x07) == 3) {
            return pos
        }
        if (pos < 0)
            break;
        lastPos = pos
    }
    throw new Error('CompressContainer not found')


}

/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/f7983581-d107-4a1f-b5f7-f3650e777c04
 */
class ObjectPool {
    /**
     * Object inside ObjectPool storage
     * @param {Buffer} ocxname 
     * @param {Buffer} content 
     */
    constructor(ocxname, content) {
        this.name = ocxname.slice(0, ocxname.length - 4).toString('utf16le').normalize().trim()
        this.content = ""
        this.raw_content = content
        let id = content.readUInt8(4)
        if (id == 0x01) {
            this.content = content.slice(0x1C, (content.readUInt32LE(0x10) & 0x00FFFFFF) + 0x1C).toString('utf-8')
        } else if (id == 0x2C) {
            this.content = content.slice(0x10, (content.readUInt32LE(0x0C) & 0x00FFFFFF) + 0x10).toString('utf-8')
        } else {
            console.error("Could parse content ID: " + id)
            this.content = ""
        }
    }
}
/**
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/fe59ffd9-79e9-44fb-a1c9-2466057b05c6
 */
class DocumentSummaryInformation {

    /**
     * 
     * @param {{
            version : number,
            system_identifier : number,
            clsid : string,
            num_property_set : number,
            fmtid0 : string,
            fmtid1 : string,
            property_set_0 : PropertySet,
            property_set_1 : PropertySet
        }} obj 
     */
    constructor(obj) {
        this.version = obj.version
        this.system_identifier = obj.system_identifier
        this.clsid = obj.clsid//16 bytes
        this.num_property_set = 0x0001
        this.fmtid0 = obj.fmtid0
        //https://docs.microsoft.com/en-us/openspecs/windows_protocols/MS-OLEPS/6e65d6fa-6044-4e23-ae71-d65d1e3b1249
        this.property_set_0 = obj.property_set_0
        if (obj.fmtid1) {
            this.num_property_set = 0x02
            this.fmtid1 = obj.fmtid1
            this.property_set_1 = obj.property_set_1
        }


    }

    osMajor(){
        return this.system_identifier & 0xFF
    }
    osMinor(){
        return (this.system_identifier >> 8) & 0xFF
    }
    osType(){
        return (this.system_identifier >> 16) & 0xFFFF
    }

    /**
     * Serialize the object in a binary structure
     */
    to_buffer() {
        let totalSize = 48 + this.property_set_0.properties.length// 2 byte_order + 2 version + 4 system_identifier + 16 clsid + 4 num_property_set + 16 fmtid0 + 4 offset0
        if (this.fmtid1) {
            totalSize += 24 + this.property_set_1.properties.length// 16 fmtid1 + 4 offset1 + 4 size1
        }
        let buff = Buffer.alloc(totalSize)
        buff.writeUInt16LE(0xFFFE)//Byte_order
        buff.writeUInt16LE(this.version, 2)
        buff.writeUInt32LE(this.system_identifier, 4)
        let clsid = from_GUID_to_buffer(this.clsid)
        clsid.copy(buff, 8, 0, 16)
        buff.writeUInt16LE(this.num_property_set, 24)
        let fmtid0 = from_GUID_to_buffer(this.fmtid0)
        fmtid0.copy(buff, 28, 0, 16)
        let offset0 = 48
        if (this.num_property_set === 2) {
            let fmtid1 = from_GUID_to_buffer(this.fmtid1)
            fmtid1.copy(buff, 48, 0, 16)
            let offset1 = 72 + this.property_set_0.properties.length //68 + 4 size + X property_set
            offset0 = 68
            buff.writeUInt32LE(offset1, 64)
            //TODO serialization of property Sets
            //this.property_set_1.copy(buff, offset1, 0)
        }
        buff.writeUInt32LE(offset0, 44)
        //this.property_set_0.copy(buff, offset0)
        return buff
    }
    /**
     * Generate a DocumentSummaryInformation from a Buffer
     * @param {Buffer} buff 
     */
    static from_buffer(buff) {
        let byte_order = buff.readUInt16LE(0)
        checkMalformed(byte_order, 0xFFFE)
        let version = buff.readUInt16LE(2)
        let system_identifier = buff.readUInt32LE(4)
        let clsid = buff.slice(8, 24)
        let num_property_set = buff.readUInt32LE(24)
        let fmtid0 = buff.slice(28, 44)
        let offset0 = buff.readUInt32LE(44)
        let prop_0_size = buff.readUInt32LE(offset0)
        //PropertySet includes 4 bytes for size
        let property_set_0 = PropertySet.from_buffer(buff.slice(offset0, offset0 + prop_0_size))

        if (num_property_set === 1)
            return new DocumentSummaryInformation({
                version,
                system_identifier,
                clsid : from_buffer_to_GUID(clsid),
                num_property_set,
                fmtid0 : from_buffer_to_GUID(fmtid0),
                property_set_0
            })

        let fmtid1 = buff.slice(48, 64)
        let offset1 = buff.readUInt32LE(64)
        let prop_1_size = buff.readUInt32LE(offset1)
        let property_set_1 = PropertySet.from_buffer(buff.slice(offset1, offset1 + prop_1_size))
        return new DocumentSummaryInformation({
            version,
            system_identifier,
            clsid : from_buffer_to_GUID(clsid),
            num_property_set,
            fmtid0 : from_buffer_to_GUID(fmtid0),
            fmtid1 : from_buffer_to_GUID(fmtid1),
            property_set_0,
            property_set_1
        })

    }
}

class Dictionary {
    /**
     * 
     * @param {Buffer} buff 
     */
    static from_buffer(buff){
        let numEntries = buff.readUInt32LE(0)
        let entries = new Map()
        let offset = 4
        for(let i = 0; i < numEntries; i++){
            let id = buff.readUInt32LE(offset)
            let size = buff.readUInt32LE(offset + 4)
            let name = buff.slice(offset + 8, offset + 8 + size).toString()
            offset += 8 + size
            entries.set(id,name)
        }
        return entries
    }
    /**
     * Serialize as Buffer
     * @param {Map} map 
     */
    static to_buffer(map){
        let total_size = 4;
        map.forEach((val,key)=>{
            total_size += 4 + val.length
        })
        //We wont be letting holes
        let buf = Buffer.allocUnsafe(total_size)
        buf.writeUInt32LE(map.size)
        let offset = 4;
        map.forEach((val,key)=>{
            buf.writeUInt32LE(key,offset)
            offset += 4
            buf.writeUInt32LE(val.length,offset)
            offset += 4
            Buffer.from(val).copy(buf,offset,0,val.length)
            offset += val.length
        })
        return buf
    }
}

class PropertySet {
    constructor({properties}){
        this.properties = properties
    }
    /**
     * 
     * @param {Buffer} buff 
     */
    static from_buffer(buff){
        let size = buff.readUInt32LE(0)
        let num_props = buff.readUInt32LE(4)
        let properties = []
        for(let i = 0; i < num_props; i++){
            let identifier = buff.readUInt32LE(8 + i*8)
            let offset = buff.readUInt32LE(12 + i*8)
            if(identifier == 0x0){
                // Dictionary Property
                //https://docs.microsoft.com/en-us/openspecs/windows_protocols/MS-OLEPS/4177a4bc-5547-49fe-a4d9-4767350fd9cf
                properties.push(Dictionary.from_buffer(buff.slice(offset)))
            }else{
                let prop = extractPropertyField(identifier,offset,buff)
                properties.push(prop)
            }
        }
        return new PropertySet({properties})
    }

    to_buffer(){

        let total_size = 8;
        let buf = []
        let initial = Buffer.allocUnsafe(4)
        let serialized_props = 0
        for(let i = 0; i < this.properties.length; i++){
            try{
                let serialized = serialize_property(this.properties[i])
                buf.push(serialized)
                total_size += serialized.length
                serialized_props++
            }catch(err){
                console.error(err)
            }
        }
        initial.writeUInt32LE(total_size)
        initial.writeUInt32LE(serialized_props)
        return Buffer.concat([initial,...buff])
    }
}


class SummaryInformation {
    /**
     * 
     * @param {{
        version : number,
        system_identifier : number,
        clsid : string,
        fmtid : string,
        properties : any,
        num_property_set : number
    }} obj 
 */
    constructor(obj){
        this.version = obj.version
        this.clsid = obj.clsid
        this.system_identifier = obj.system_identifier
        this.fmtid = obj.fmtid
        this.properties = obj.properties
        this.num_property_set = obj.num_property_set
    }

    osMajor(){
        return this.system_identifier & 0xFF
    }
    osMinor(){
        return (this.system_identifier >> 8) & 0xFF
    }
    osType(){
        return (this.system_identifier >> 16) & 0xFFFF
    }

    to_buffer(){
        let totalSize = 56//Incresed by 4 + X for each property
        let buff = Buffer.alloc(totalSize)
        buff.writeUInt16LE(0xFFFE)//Byte_order
        buff.writeUInt16LE(this.version, 2)
        buff.writeUInt32LE(this.system_identifier, 4)
        let clsid = from_GUID_to_buffer(this.clsid)
        clsid.copy(buff, 8, 0, 16)
        buff.writeUInt16LE(this.num_property_set, 24)
        let fmtid = from_GUID_to_buffer(this.fmtid)
        fmtid.copy(buff, 28, 0, 16)
        buff.writeUInt32LE(48,44)//PropertySet start in 48
        let offset = 48
        let props = Object.keys(this.properties)
        //Write size at the end of the for
        let size = 8
        let offset_prop = props.length * 8 + 8
        buff.writeUInt32LE(props.length, offset + 4)
        let properties_offsets = []
        let properties_types = []
        for(let i = 0;i < props.length; i++){
            let buff_prop = serialize_property(this.properties[props[i]])
            let buff_offset = Buffer.alloc(8)
            buff_offset.writeUInt32LE(property_ID_from_String(props[i]),0)
            buff_offset.writeUInt32LE(offset_prop,4)
            properties_offsets.push(buff_offset)
            properties_types.push(buff_prop)
            size += 8 + buff_prop.length
            offset_prop += buff_prop.length
        }
        buff.writeUInt32LE(size, offset)
        buff = Buffer.concat([buff, ...properties_offsets, ...properties_types])
        return buff
    }

    /**
     * 
     * @param {Buffer} buff 
     */
    static from_buffer(buff){
        let byte_order = buff.readUInt16LE(0)
        checkMalformed(byte_order, 0xFFFE)
        let version = buff.readUInt16LE(2)
        let system_identifier = buff.readUInt32LE(4)
        let clsid = buff.slice(8, 24)
        let fmtid = buff.slice(28, 44)
        let num_property_set = buff.readUInt32LE(24)
        let offset = buff.readUInt32LE(44)
        let prop_size = buff.readUInt32LE(offset)
        let start_prop_set = offset
        let num_properties = buff.readUInt32LE(offset + 4)
        offset += 8
        let properties = {}
        for(let i = 0; i < num_properties; i++){
            let id = buff.readUInt32LE(offset + i*8)
            let offsetProp = buff.readUInt32LE(offset + i*8 + 4)
            try{
                let prop = extractPropertyField(id,start_prop_set + offsetProp,buff)
                properties[property_ID_to_String(id)] = prop
            }catch(err){
                console.error(err)
            }
            
        }
        return new SummaryInformation({
            version,
            system_identifier,
            clsid : from_buffer_to_GUID(clsid),
            fmtid : from_buffer_to_GUID(fmtid),
            num_property_set : num_property_set,
            properties
        })

    }
}

const PROPERTY_CODE_PAGE = 0x0001
const PROPERTY_CODE_PAGE_NAME = "CODEPAGE"
const PROPERTY_PIDSI_TITLE = 0x00000002; 
const PROPERTY_PIDSI_TITLE_NAME = "TITLE"
const PROPERTY_PIDSI_SUBJECT = 0x00000003; 
const PROPERTY_PIDSI_SUBJECT_NAME = "SUBJECT"
const PROPERTY_PIDSI_AUTHOR = 0x00000004; 
const PROPERTY_PIDSI_AUTHOR_NAME = "AUTHOR"
const PROPERTY_PIDSI_KEYWORDS = 0x00000005; 
const PROPERTY_PIDSI_KEYWORDS_NAME = "KEYWORDS"
const PROPERTY_PIDSI_COMMENTS = 0x00000006; 
const PROPERTY_PIDSI_COMMENTS_NAME = "COMMENTS"
const PROPERTY_PIDSI_TEMPLATE = 0x00000007; 
const PROPERTY_PIDSI_TEMPLATE_NAME = "TEMPLATE"
const PROPERTY_PIDSI_LASTAUTHOR = 0x00000008; 
const PROPERTY_PIDSI_LASTAUTHOR_NAME = "LASTAUTHOR"
const PROPERTY_PIDSI_REVNUMBER = 0x00000009; 
const PROPERTY_PIDSI_REVNUMBER_NAME = "REVNUMBER"
const PROPERTY_PIDSI_EDITTIME = 0x0000000A; 
const PROPERTY_PIDSI_EDITTIME_NAME = "EDITTIME"
const PROPERTY_PIDSI_LASTPRINTED = 0x0000000B; 
const PROPERTY_PIDSI_LASTPRINTED_NAME = "LASTPRINTED"
const PROPERTY_PIDSI_CREATE_DTM = 0x0000000C; 
const PROPERTY_PIDSI_CREATE_DTM_NAME = "CREATE_DTM"
const PROPERTY_PIDSI_LASTSAVE_DTM = 0x0000000D; 
const PROPERTY_PIDSI_LASTSAVE_DTM_NAME = "LASTSAVE_DTM"
const PROPERTY_PIDSI_PAGECOUNT = 0x0000000E; 
const PROPERTY_PIDSI_PAGECOUNT_NAME = "PAGECOUNT"
const PROPERTY_PIDSI_WORDCOUNT = 0x0000000F; 
const PROPERTY_PIDSI_WORDCOUNT_NAME = "WORDCOUNT"
const PROPERTY_PIDSI_CHARCOUNT = 0x00000010; 
const PROPERTY_PIDSI_CHARCOUNT_NAME = "CHARCOUNT"
const PROPERTY_PIDSI_THUMBNAIL = 0x00000011; 
const PROPERTY_PIDSI_THUMBNAIL_NAME = "THUMBNAIL"
const PROPERTY_PIDSI_APPNAME = 0x00000012; 
const PROPERTY_PIDSI_APPNAME_NAME = "APPNAME"
const PROPERTY_PIDSI_DOC_SECURITY = 0x00000013; 
const PROPERTY_PIDSI_DOC_SECURITY_NAME = "DOC_SECURITY"

const PROPERTIES_VALUES = [PROPERTY_CODE_PAGE, PROPERTY_PIDSI_TITLE, PROPERTY_PIDSI_SUBJECT, PROPERTY_PIDSI_AUTHOR, PROPERTY_PIDSI_KEYWORDS, PROPERTY_PIDSI_COMMENTS, PROPERTY_PIDSI_TEMPLATE, PROPERTY_PIDSI_LASTAUTHOR, PROPERTY_PIDSI_REVNUMBER, PROPERTY_PIDSI_EDITTIME, PROPERTY_PIDSI_LASTPRINTED, PROPERTY_PIDSI_CREATE_DTM, PROPERTY_PIDSI_LASTSAVE_DTM, PROPERTY_PIDSI_PAGECOUNT, PROPERTY_PIDSI_WORDCOUNT, PROPERTY_PIDSI_CHARCOUNT, PROPERTY_PIDSI_THUMBNAIL, PROPERTY_PIDSI_APPNAME, PROPERTY_PIDSI_DOC_SECURITY]
const PROPERTIES_NAMES = [PROPERTY_CODE_PAGE_NAME, PROPERTY_PIDSI_TITLE_NAME, PROPERTY_PIDSI_SUBJECT_NAME, PROPERTY_PIDSI_AUTHOR_NAME, PROPERTY_PIDSI_KEYWORDS_NAME, PROPERTY_PIDSI_COMMENTS_NAME, PROPERTY_PIDSI_TEMPLATE_NAME, PROPERTY_PIDSI_LASTAUTHOR_NAME, PROPERTY_PIDSI_REVNUMBER_NAME, PROPERTY_PIDSI_EDITTIME_NAME, PROPERTY_PIDSI_LASTPRINTED_NAME, PROPERTY_PIDSI_CREATE_DTM_NAME, PROPERTY_PIDSI_LASTSAVE_DTM_NAME, PROPERTY_PIDSI_PAGECOUNT_NAME, PROPERTY_PIDSI_WORDCOUNT_NAME, PROPERTY_PIDSI_CHARCOUNT_NAME, PROPERTY_PIDSI_THUMBNAIL_NAME, PROPERTY_PIDSI_APPNAME_NAME, PROPERTY_PIDSI_DOC_SECURITY_NAME]

function property_ID_to_String(id){
    let prop= PROPERTIES_VALUES.indexOf(id)
    if(prop >= 0){
        return PROPERTIES_NAMES[prop]
    }
    return "" + id
}
function property_ID_from_String(id){
    let prop= PROPERTIES_NAMES.indexOf(id)
    if(prop >= 0){
        return PROPERTIES_VALUES[prop]
    }
    return parseInt(id)
}

const VT_EMPTY = 0x0000
const VT_NULL = 0x0001
const VT_I2 = 0x0002
const VT_I4 = 0x0003
const VT_R4 = 0x0004
const VT_R8 = 0x0005
const VT_CY = 0x0006
const VT_DATE = 0x0007
const VT_BSTR = 0x0008
const VT_ERROR = 0x000A
const VT_BOOL = 0x000B
const VT_VARIANT = 0x000C
const VT_DECIMAL = 0x000E
const VT_I1 = 0x0010
const VT_UI1 = 0x0011
const VT_UI2 = 0x0012
const VT_UI4 = 0x0013
const VT_I8 = 0x0014
const VT_UI8 = 0x0015
const VT_INT = 0x0016
const VT_UINT = 0x0017
const VT_LPSTR = 0x001E
const VT_LPWSTR = 0x001F
const VT_FILETIME = 0x0040
const VT_BLOB = 0x0041
const VT_STREAM = 0x0042
const VT_STORAGE = 0x0043
const VT_STREAMED_OBJECT = 0x0044
const VT_STORED_OBJECT = 0x0045
const VT_BLOB_OBJECT = 0x0046
const VT_CF = 0x0047
const VT_CLSID = 0x0048
const VT_VERSIONED_STREAM = 0x0049
const VT_VECTOR_VT_I2 = 0x1002
const VT_VECTOR_VT_I4 = 0x1003
const VT_VECTOR_VT_R4 = 0x1004
const VT_VECTOR_VT_R8 = 0x1005
const VT_VECTOR_VT_CY = 0x1006
const VT_VECTOR_VT_DATE = 0x1007
const VT_VECTOR_VT_BSTR = 0x1008
const VT_VECTOR_VT_ERROR = 0x100A
const VT_VECTOR_VT_BOOL = 0x100B
const VT_VECTOR_VT_VARIANT = 0x100C
const VT_VECTOR_VT_I1 = 0x1010
const VT_VECTOR_VT_UI1 = 0x1011
const VT_VECTOR_VT_UI2 = 0x1012
const VT_VECTOR_VT_UI4 = 0x1013
const VT_VECTOR_VT_I8 = 0x1014
const VT_VECTOR_VT_UI8 = 0x1015
const VT_VECTOR_VT_LPSTR = 0x101E
const VT_VECTOR_VT_LPWSTR = 0x101F
const VT_VECTOR_VT_FILETIME = 0x1040
const VT_VECTOR_VT_CF = 0x1047
const VT_VECTOR_VT_CLSID = 0x1048
const VT_ARRAY_VT_I2 = 0x2002
const VT_ARRAY_VT_I4 = 0x2003
const VT_ARRAY_VT_R4 = 0x2004
const VT_ARRAY_VT_R8 = 0x2005
const VT_ARRAY_VT_CY = 0x2006
const VT_ARRAY_VT_DATE = 0x2007
const VT_ARRAY_VT_BSTR = 0x2008
const VT_ARRAY_VT_ERROR = 0x200A
const VT_ARRAY_VT_BOOL = 0x200B
const VT_ARRAY_VT_VARIANT = 0x200C
const VT_ARRAY_VT_DECIMAL = 0x200E
const VT_ARRAY_VT_I1 = 0x2010
const VT_ARRAY_VT_UI1 = 0x2011
const VT_ARRAY_VT_UI2 = 0x2012
const VT_ARRAY_VT_UI4 = 0x2013
const VT_ARRAY_VT_INT = 0x2016
const VT_ARRAY_VT_UINT = 0x2017

/**
 * https://docs.microsoft.com/en-us/openspecs/windows_protocols/MS-OLEPS/f122b9d7-e5cf-4484-8466-83f6fd94b3cc
 * @param {number} id 
 * @param {number} offset 
 * @param {Buffer} buffer 
 */
function extractPropertyField(id, offset, buffer, codePage = 'utf8'){
    //TypedPropertyValue
    let type = buffer.readUInt16LE(offset)
    let padding = buffer.readUInt16LE(offset + 2)//0x00 always
    switch(type){
        case VT_EMPTY: return {id,type,value : null};
        case VT_NULL: return {id,type,value : null};
        case VT_I2: return {id,type,value : buffer.readInt16LE(offset + 4), totalSize : 6};
        case VT_I4: return {id,type,value : buffer.readInt32LE(offset + 4),totalSize : 8};
        case VT_R4: return {id,type,value : buffer.readFloatLE(offset + 4),totalSize : 8};
        case VT_R8: return {id,type,value : buffer.readDoubleLE(offset + 4),totalSize : 12};
        case VT_CY: return {id,type,value : buffer.readBigInt64LE(offset + 4),totalSize : 12};
        case VT_DATE: return {id,type,value : buffer.readBigInt64LE(offset + 4),totalSize : 12};
        case VT_BSTR: {//CodePageString https://docs.microsoft.com/en-us/openspecs/windows_protocols/MS-OLEPS/a4c32611-5b79-4965-8f50-50639c138e16
            let size = buffer.readInt32LE(offset + 4)
            return {id,type,value : buffer.slice(offset + 8, offset + 8 + size).toString(codePage), totalSize : 8 + size};
        };
        case VT_ERROR: return {id,type,value : buffer.readUInt32LE(offset + 4),totalSize : 8};
        //https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/7b39eb24-9d39-498a-bcd8-75c38e5823d0
        case VT_BOOL: return {id,type,value : buffer.readUInt16LE(offset + 4) === 0xFFFF ? true : false, totalSize : 6};
        case VT_DECIMAL : throw new Error("VT_DECIMAL Not implemented")
        case VT_VARIANT : {
            throw new Error("VT_VARIANT Not implemented")
        }
        case VT_LPSTR: {
            let size = buffer.readInt32LE(offset + 4)
            return {id,type,value : buffer.slice(offset + 8, offset + 8 + size).toString(codePage), totalSize : 8 + size};
        };
        case VT_LPWSTR: {
            let size = buffer.readInt32LE(offset + 4)
            return {id,type,value : buffer.slice(offset + 8, offset + 8 + size).toString('utf16le'), totalSize : 8 + size};
        }
        case VT_FILETIME : {
            let dwLowDateTime = buffer.readInt32LE(offset + 4)
            let dwHighDateTime = buffer.readInt32LE(offset + 8)
            return {id,type,dwLowDateTime,dwHighDateTime, totalSize : 12}
        }
        case VT_VECTOR_VT_LPSTR : {
            let v_head_lngth = buffer.readInt32LE(offset + 4)
            let subofset = offset + 8
            let str_array = new Array(v_head_lngth)
            for(let i = 0; i < v_head_lngth; i++){
                let size = buffer.readInt32LE(subofset)
                str_array[i] = buffer.slice(subofset + 4, subofset + 4 + size - 1).toString(codePage);//Remove trailing 0
                subofset += 4 + size
            }
            return {id,type, value : str_array, totalSize : 8 + subofset - offset}
        }
        case VT_VECTOR_VT_VARIANT : {
            let v_head_lngth = buffer.readInt32LE(offset + 4)
            let subofset = offset + 8
            let str_array = new Array(v_head_lngth)
            for(let i = 0; i < v_head_lngth; i++){
                let vt_variant = extractPropertyField("",subofset,buffer)
                subofset += vt_variant.totalSize
                str_array[i] = vt_variant
            }
            return {id,type, value : str_array, totalSize : 8 + subofset - offset}
        }
        default: throw new Error(type + " Not implemented")


    }
}

function serialize_property(prop){
    let type = prop.type
    let value = prop.value
    switch(type){
        case VT_LPSTR: {
            let buff = Buffer.alloc(8 + value.length)
            buff.writeUInt32LE(type, 0)
            buff.writeUInt32LE(value.length, 4)
            for(let i = 0; i < value.length; i++){
                buff[8 + i] = value.charCodeAt(i)
            }
            return buff
        };
        case VT_I2: {
            let buff = Buffer.alloc(8)
            buff.writeUInt32LE(type, 0)
            buff.writeUInt16LE(value, 4)
            return buff
        };
        case VT_I4: {
            let buff = Buffer.alloc(8)
            buff.writeUInt32LE(type, 0)
            buff.writeUInt32LE(value, 4)
            return buff
        };
        case VT_FILETIME : {
            let buff = Buffer.alloc(12)
            buff.writeUInt32LE(type, 0)
            buff.writeUInt32LE(prop.dwLowDateTime, 4)
            buff.writeUInt32LE(prop.dwHighDateTime, 8)
            return buff
        }
        default: throw new Error(type + " Not implemented")
    }
}
/**
 * 
 * @param {Buffer} buff 
 */
function from_buffer_to_GUID(buff){
    if(buff.length < 16){
        throw new Error('Not valid buffer. Length must be 16 bytes')
    }
    return (buff.slice(0,4).toString('hex') + "-" + buff.slice(4,6).toString('hex') + "-" + buff.slice(6,8).toString('hex') + "-" + buff.slice(8,10).toString('hex') + "-" + buff.slice(10,16).toString('hex')).toUpperCase()
}

/**
 * 
 * @param {string} guid 
 */
function from_GUID_to_buffer(guid){
    let buff = Buffer.alloc(16)
    guid = guid.replace(/\-/g,"")
    for(let i = 0; i < 16; i++){
        buff.writeUInt8(parseInt(guid.slice(i*2,i*2+2),16),i)
    }
    return buff
}


exports.ObjectPool = ObjectPool
exports.VbaDirInformationRecord = VbaDirInformationRecord
exports.VbaDirStream = VbaDirStream
exports.VbaProjectStream = VbaProjectStream
exports.VbaModule = VbaModule
exports.DocumentSummaryInformation = DocumentSummaryInformation
exports.SummaryInformation = SummaryInformation
exports.localizeCompressContainer = localizeCompressContainer
exports.Vba_VBAProjectStream = Vba_VBAProjectStream