const zlib = require('zlib')
const util = require('util');
const {OleCompoundDoc} = require('./ole-doc')
const unzip_async = util.promisify(zlib.unzip)

class ActiveMime {
    /**
     * 
     * @param {Buffer} buff 
     */
    constructor(buff){
        let is_ole_doc = buff.readUInt32LE(0) == 0xE011CFD0
        if(is_ole_doc){
            this.ole_doc = new OleCompoundDoc(buff)
        }
        
    }
    /**
     * 
     * @param {Buffer} buffer 
     */
    static async from_buffer(buffer){
        let header = buffer.slice(0,12)
        let unknown = buffer.readUInt16LE(12)
        checkMalformed(unknown, 0xf001)
        let field_size = buffer.readUInt32LE(14)
        if (field_size === 4){
            unknown = buffer.readUInt32LE(18)
            checkMalformed(unknown, 0xffffffff)
        }else if(field_size === 8){
            unknown = buffer.readUInt64LE(18)
            checkMalformed(unknown, 0xffffffffffffffff)
        }else {
            throw new Error('Malformed field_size')
        }
        unknown = buffer.readUInt32LE(18 + field_size)
        checkMalformed(unknown & 0xF0f0ff0f, -268435456)
        let compressed_size = buffer.readUInt32LE(22 + field_size)
        let field_size_d = buffer.readUInt32LE(26 + field_size)
        let field_size_e = buffer.readUInt32LE(30 + field_size)
        unknown = field_size_d === 4 ? buffer.readUInt32LE(34 + field_size) : buffer.readUInt64LE(34 + field_size);
        checkMalformed(unknown,field_size_d === 4 ? 0x00 : 0x0000000000000001)
        let vba_tail_type = field_size_e === 4 ? buffer.readUInt32LE(34 + field_size_d + field_size) : buffer.readUInt64LE(34 + field_size_d + field_size);
        let size = buffer.readUInt32LE(34 + field_size + field_size_d + field_size_e)
        let compressed_data = buffer.slice(38 + field_size + field_size_d + field_size_e)
        const data = await unzip_async(compressed_data)
        let toRet = new ActiveMime(data)
        await toRet.ole_doc.read()
        return toRet
    }
}
function checkMalformed(code, expected) {
    if (code != expected) {
        throw new Error("Malformed code. Expected: " + expected + " get " + code)
    }
}

exports.ActiveMime = ActiveMime