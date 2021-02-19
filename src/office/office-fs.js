const { OleCompoundDoc } = require("./ole-doc")
const { ActiveMime } = require('./active-mime')
const unzipper = require('unzipper')
const { basename, dirname } = require('path')
/**
 * Abstracción de clases ZIP y OLE
 */
class OfficeFileFS {

    constructor(data) {
        if (data.active_mime) {
            this.active_mime = data.active_mime
        } else if (data.ole_doc) {
            this.ole_doc = data.ole_doc
        } else if (data.zip) {
            this.zip = data.zip
        } else {
            throw new Error("File cannot be detected")
        }
    }
    /**
     * 
     * @param {Buffer} buff 
     */
    static async from_buffer(buff) {
        if (buff.slice(0, ACTIVE_MIME_MAGIC.length).compare(ACTIVE_MIME_MAGIC) === 0) {
            let active_mime = await ActiveMime.from_buffer(buff)
            return new OfficeFileFS({ active_mime })
        } else if (buff.slice(0, OLE_MAGIC.length).compare(OLE_MAGIC) === 0) {
            let ole_doc = new OleCompoundDoc(buff)
            await ole_doc.read()
            return new OfficeFileFS({ ole_doc })
        } else if (is_zip(buff)) {
            let zip = await unzipper.Open.buffer(buff);
            return new OfficeFileFS({ zip })
        } else {
            throw new Error("File cannot be detected")
        }
    }

    /**
     * 
     * @param {string} dir_name 
     */
    ls_dir(dir_name) {

        let pth = dir_name.split("/")

        if (this.ole_doc) {
            let storage = this.ole_doc
            for (let i = 0; i < pth.length; i++) {
                if (pth[i].trim() === "") {
                    continue
                }
                if (storage) {
                    storage = storage.storage(pth[i])
                } else {
                    return null
                }
            }
            return [...storage.storageList(), ...storage.streamList()]

        } else if (this.zip) {
            dir_name = "/" + dir_name
            if (!dir_name.endsWith("/")) {
                dir_name = dir_name + "/"
            }
            let toRet = new Set()
            for (let i = 0; i < this.zip.files.length; i++) {
                let child = "/" + this.zip.files[i].path
                if (child === dir_name) continue
                if (!child.startsWith(dir_name)) continue

                let subchild = child.replace(dir_name, "");
                let subpos = subchild.indexOf("/")
                toRet.add(subchild.slice(0, subpos < 0 ? subchild.length : subpos))
            }
            return [...toRet]
        } else if (this.active_mime) {
            let storage = this.active_mime
            for (let i = 0; i < pth.length; i++) {
                if (pth[i].trim() === "") {
                    continue
                }
                if (storage) {
                    storage = storage.storage(pth[i])
                } else {
                    return null
                }
            }
            return [...storage.storageList(), ...storage.streamList()]
        } else {
            throw new Error('Not valid')
        }

    }
    /**
     * 
     * @param {string} dir_name 
     */
    ls_dir2(dir_name) {

        let pth = dir_name.split("/")

        if (this.ole_doc) {
            let storage = this.ole_doc
            for (let i = 0; i < pth.length; i++) {
                if (pth[i].trim() === "") {
                    continue
                }
                if (storage) {
                    storage = storage.storage(pth[i])
                } else {
                    return null
                }
            }
            return { storage: storage.storageList(), streams: storage.streamList() }

        } else if (this.zip) {
            dir_name = "/" + dir_name
            if (!dir_name.endsWith("/")) {
                dir_name = dir_name + "/"
            }
            let streams = new Set()
            let storage = new Set()
            for (let i = 0; i < this.zip.files.length; i++) {
                let child = "/" + this.zip.files[i].path
                if (child === dir_name) continue
                if (!child.startsWith(dir_name)) continue

                let subchild = child.replace(dir_name, "");
                let subpos = subchild.indexOf("/")
                if(subpos >= 0){
                    storage.add(subchild.slice(0, subpos))
                }else{
                    streams.add(subchild.slice(0, subchild.length))
                }
                
            }
            return { storage : [...storage], streams : [...streams] }
        } else if (this.active_mime) {
            let storage = this.active_mime
            for (let i = 0; i < pth.length; i++) {
                if (pth[i].trim() === "") {
                    continue
                }
                if (storage) {
                    storage = storage.storage(pth[i])
                } else {
                    return null
                }
            }
            return { storage: storage.storageList(), streams: storage.streamList() }
        } else {
            throw new Error('Not valid')
        }

    }
    async read_file(dir_name) {
        let pth = dir_name.split("/")
        if (this.ole_doc) {
            let stream = this.ole_doc
            for (let i = 0; i < pth.length - 1; i++) {
                if (pth[i].trim() === "") {
                    continue
                }
                if (stream) {
                    stream = await stream.storage(pth[i])
                } else {
                    return null
                }
            }
            try{
                return await stream.stream(pth[pth.length - 1])
            }catch(e){
                return Buffer.from(e.toString())
            }
            

        } else if (this.zip) {
            if (dir_name.startsWith("/")) {
                dir_name = dir_name.slice(1)
            }
            try {
                for (let i = 0; i < this.zip.files.length; i++) {
                    if (this.zip.files[i].path == dir_name) {
                        return await this.zip.files[i].buffer();
                    }
                }
            }
            catch (e) {
                return Buffer.from(e.toString())
            }
            return Buffer.from("Error")

        } else if (this.active_mime) {
            let stream = this.active_mime
            for (let i = 0; i < pth.length - 1; i++) {
                if (pth[i].trim() === "") {
                    continue
                }
                if (stream) {
                    stream = stream.storage(pth[i])
                } else {
                    return null
                }
            }
            return stream.stream(pth[pth.length - 1])
        } else {
            throw new Error('Not valid')
        }
    }
}
const ACTIVE_MIME_MAGIC = Buffer.from('ActiveMime', 'ascii')
const OLE_MAGIC = Buffer.from("D0CF11E0A1B11AE1", 'hex')
const ZIP_MAGIC = Buffer.from('504B0304', 'hex')
/**
 * 
 * @param {Buffer} buffer 
 */
function applyParserByMagic(buffer) {
    if (buffer.slice(0, ACTIVE_MIME_MAGIC.length).compare(ACTIVE_MIME_MAGIC) === 0) {
        throw new Error('Cannot parse MSO/ActiveMime')
    } else if (buffer.slice(0, OLE_MAGIC.length).compare(OLE_MAGIC) === 0) {
        return new OleCompoundDoc(buffer)
    } else if (is_zip(buffer)) {
        //Its a zip file
        return unzipper.Open.buffer(buffer);
    } else {
        throw new Error("File cannot be detected")
    }

}
/**
 * 
 * @param {Buffer} buffer 
 */
function is_zip(buffer) {
    //0x50, 0x4b, 0x03, 0x04
    return buffer.slice(0, ZIP_MAGIC.length).compare(ZIP_MAGIC) === 0
}

const isChildOf = (child, parent) => {
    if (child === parent) return false
    if (!child.startsWith(parent)) return false
    const parentTokens = parent.split('/').filter(i => i.length)
    const childTokens = child.split('/').filter(i => i.length)
    if (childTokens.length != (parentTokens.length + 1)) {
        return false
    }

    return parentTokens.every((t, i) => childTokens[i] === t)
}

exports.OfficeFileFS = OfficeFileFS
exports.is_zip = is_zip