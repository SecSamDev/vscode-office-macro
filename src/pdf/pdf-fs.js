class PdfFS {

    constructor(data) {
        this.objects = data.objects
        this.scripts = data.scripts
    }
    /**
     * 
     * @param {Buffer} buff 
     */
    static from_buffer(buff) {
        if (buff.indexOf("%PDF") != 0) {
            throw new Error("This is not a valid PDF")
        }
        let objects = {}
        let scriptsId = []
        let lastNewLinePos = 0
        while (true) {
            let nextLinePos = buff.indexOf("\n", lastNewLinePos)
            if (nextLinePos < 0) {
                break
            }
            let line = buff.slice(lastNewLinePos, nextLinePos)
            if (line.length == 0 || line[0] == "%") {
                lastNewLinePos = nextLinePos + 1
                continue
            }
            let strLine = line.toString().trim()
            if (strLine.length == 0 || strLine.startsWith("%")) {
                lastNewLinePos = nextLinePos + 1
                continue
            }
            if (strLine.endsWith("obj")) {
                let obj = PdfObject.from_buffer(buff.slice(lastNewLinePos))
                if (obj) {
                    objects[obj.obj.id] = obj.obj
                    if(obj.obj.metadata && obj.obj.metadata.JS){
                        scriptsId.push(obj.obj.metadata.JS.id)
                    }
                    lastNewLinePos += obj.lastPos
                    continue
                }
            }
            lastNewLinePos = nextLinePos + 1
        }
        let scripts = []
        for(let id of scriptsId){
            scripts.push(objects[id].content.toString().replace(/\r/g,"\n"))
        }
        return new PdfFS({
            scripts,
            objects,
            metadata: {
                //TODO: extract metadata
            }
        })
        
    }
}
module.exports.PdfFS = PdfFS

class PdfObject {
    /**
     * @param {number} id 
     * @param {string} metadata 
     * @param {string} content 
     */
    constructor(id, metadata, content) {
        this.id = id
        this.metadata = metadata
        this.content = content
    }

    /**
     * 
     * @param {Buffer} buff 
     * @returns {{obj : PdfObject, lastPos : number}}
     */
    static from_buffer(buff) {
        let nextLinePos = buff.indexOf("\n")
        let lastLinePos = nextLinePos
        let strLine = buff.slice(0, nextLinePos).toString().trim()
        if (!strLine.endsWith("obj")) {
            return null
        }
        let objctSplit = strLine.split(" ")
        let objId = objctSplit[0]

        //---------------------------------------------------------- METADATA
        nextLinePos = buff.indexOf("\n", lastLinePos + 1)
        strLine = buff.slice(lastLinePos + 1, nextLinePos).toString().trim()
        let metadata_obj = {}
        if (strLine.startsWith("<<")) {
            //Extract metadata
            let {metadata, next_line} = extractMetadataInfo(buff.slice(nextLinePos + 1))
            metadata_obj = metadata
            lastLinePos = nextLinePos + 1 + next_line
        }
        let content = null
        nextLinePos = buff.indexOf("\n", lastLinePos)
        strLine = buff.slice(lastLinePos, nextLinePos).toString().trim()
        if (strLine == "stream") {
            let subbuf = buff.slice(nextLinePos + 1)
            let endstream = findEndStrem(subbuf)
            //Extract metadata
            content = subbuf.slice(0, endstream)
            lastLinePos = nextLinePos + endstream + 12
        }else if (strLine == "endobj") {
            return {
                obj : new PdfObject(objId,metadata_obj,content),
                lastPos : nextLinePos + 1
            }
        }else{
            //Extract plain content
            let subbuf = buff.slice(nextLinePos + 1)
            let endObj = subbuf.indexOf("endobj")
            if(endObj < 0){
                throw new Error("Could not found endobj")
            }
            content = subbuf.slice(0,endObj).toString()
            lastLinePos = nextLinePos + 1 + endObj
        }
        nextLinePos = buff.indexOf("\n", lastLinePos)
        strLine = buff.slice(lastLinePos, nextLinePos).toString().trim()
        if (strLine == "endobj") {
            return {
                obj : new PdfObject(objId,metadata_obj,content),
                lastPos : nextLinePos + 1
            }
        }
        throw new Error("Format error")


    }
}

/**
 * 
 * @param {Buffer} buff 
 */
function extractMetadataInfo(buff) {
    let lastLinePos = 0
    let metadata = {}
    let arrayPos = 0
    while (true) {
        let nextLinePos = buff.indexOf("\n", lastLinePos)
        if (nextLinePos < 0) {
            throw new Error("Object not finished in pos: " + (buff.byteOffset + lastLinePos))
        }
        let line = buff.slice(lastLinePos, nextLinePos).toString().trim()
        if (line == ">>" || line == "]") {//Metadata end
            lastLinePos = nextLinePos + 1
            break
        }
        if (line.startsWith("/")) {
            let splited = cleanSplited(line.split(" "))
            if (splited.length == 4) {
                //Contains ObjID
                metadata[splited[0].slice(1)] = {
                    type: splited[0].slice(1),
                    id: Number(splited[1]),
                    sid: Number(splited[2]),
                    chr: splited[3],
                }
            } else if (splited.length == 2) {
                if(splited[1] == "<<"){
                    let submetadata = extractMetadataInfo(buff.slice(nextLinePos+1))
                    metadata[splited[0].slice(1)] = {
                        type: splited[0].slice(1),
                        content: submetadata.metadata
                    }
                    lastLinePos = nextLinePos+1 +submetadata.next_line
                    continue
                }else if(splited[1] == "["){
                    let submetadata = extractMetadataInfo(buff.slice(nextLinePos+1))
                    metadata[splited[0].slice(1)] = {
                        type: splited[0].slice(1),
                        content: submetadata.metadata
                    }
                    lastLinePos = nextLinePos+1 +submetadata.next_line
                    continue
                }else{
                    metadata[splited[0].slice(1)] = {
                        type: splited[0].slice(1),
                        content: splited[1]
                    }
                }
            }else if(splited.length == 1){
                metadata[splited[0].slice(1)] = {
                    type: splited[0].slice(1)
                }
            }
        }else{
            metadata[arrayPos] = line
            arrayPos += 1
        }
        lastLinePos = nextLinePos + 1
    }
    return {metadata, next_line : lastLinePos}
}

/**
 * 
 * @param {Buffer} buff 
 */
function findEndStrem(buff) {
    let endstream = buff.indexOf("\nendstream")
    if(endstream < 0){
        throw new Error("Object finished withouth finding endstream")
    }
    let endObj = buff.slice(endstream).indexOf("\nendobj")
    if(endObj < 0){
        throw new Error("Object finished withouth finding endobj")
    }
    let strcontent = buff.slice(endstream + 10, endstream + endObj).toString().trim()
    if(strcontent == ""){
        return endstream
    }else{
        return findEndStrem(buff.slice(endstream + 10))
    }
}


/**
 * 
 * @param {string[]} splt 
 */
function cleanSplited(splt) {
    let ret = []
    for(let i = 0; i < splt.length; i++){
        let clean = splt[i].trim()
        if(clean != ""){
            ret.push(clean)
        }
    }
    return ret
}

class PdfBuffer {
    /**
     * 
     * @param {Buffer} buff 
     */
    constructor(buff) {
        this.buff = buff
        this.position = 0
        this.lastToken = null
    }
    nextContent(){

    }
}