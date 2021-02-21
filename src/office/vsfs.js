const { basename } = require('path')
const { OfficeFileFS } = require('./office-fs')
const vscode = require('vscode')

const { BiffDocument } = require('./office-biff')
const { OleCompoundDoc } = require('./ole-doc')
const { DocumentSummaryInformation, SummaryInformation, VbaDirInformationRecord, VbaModule, VbaProjectStream, VbaDirStream, ObjectPool } = require('./office-parser')

//VSCode FS
class VSCodeOfficeFS {
    /**
     * 
     * @param {OfficeFileFS} doc 
     */
    constructor(doc) {
        this.doc = doc
        this.elements = {}
        this.special_files = {}
        this.tree_structure = {
            "" :  { type: vscode.FileType.Directory, ctime:0, mtime:0, size: 0 }
        }
        this.dir_stream = null
    }
    stat(uri = "") {
        if (uri.startsWith("/")) {
            uri = uri.slice(1)
        }
        if (this.tree_structure[uri]) {
            return this.tree_structure[uri]
        }
        if (this.special_files[uri]) {
            return { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 };
        }
        return { type: vscode.FileType.Unknown, ctime:0, mtime:0, size: 0 };
    }

    async readFile(uri = "") {
        if (uri.startsWith("/")) {
            uri = uri.slice(1)
        }
        if (uri.includes(".bin$/")) {
            //Sub FS inside a .bin
            let binPath = uri.slice(0, uri.indexOf(".bin$") + 4);
            let contentPath = uri.slice(uri.indexOf(".bin$") + 6);
            return this.elements[binPath].readFile(contentPath)
        }
        if (uri in this.special_files) {
            if(!this.special_files[uri]){
                if(uri.endsWith(".vba") || uri.endsWith(".pcode")){
                    let fileName = uri.replace(".vba","").replace(".pcode","")
                    let project_file = await this.doc.read_file(fileName)
                    let module_offset = this.dir_stream.references_record.modules.find(val => fileName.includes(val.name));
                    let md = VbaModule.from_buffer(project_file,module_offset ? module_offset.source_offset : 0)
                    this.special_files[fileName + ".vba"] = md.source_code
                    this.special_files[fileName + ".pcode"] = md.performance_cache
                    return Buffer.from(this.special_files[uri])
                }
                return Buffer.from("<ERROR PROCESSING>")
            }
            return Buffer.from(this.special_files[uri])
        }
        return this.doc.read_file(uri)
    }

    async readDirectory(uri = "") {
        if (uri.startsWith("/")) {
            uri = uri.slice(1)
        }
        const ctime = 0;
        const mtime = 0;

        let pths = this.doc.ls_dir2(uri);
        if (uri.includes(".bin$/")) {
            //Sub FS inside a .bin
            let binPath = uri.slice(0, uri.indexOf(".bin$") + 4);
            let contentPath = uri.slice(uri.indexOf(".bin$") + 6);
            pths = this.elements[binPath].ls_dir2(contentPath)
        }
        let paths = []
        for (let i = 0; i < pths.storage.length; i++) {
            let pth = pths.storage[i].toString()
            let filePath = uri== "" ? pth :  uri + "/" + pth
            paths.push([pth, vscode.FileType.Directory])
            this.tree_structure[filePath] = { type: vscode.FileType.Directory, ctime:0, mtime:0, size: 0 }
            if(uri == "ObjectPool"){
                let content = await this.doc.read_file(filePath + "/contents")
                let ocxname = await this.doc.read_file(filePath + "/OCXNAME")
                let obj = new ObjectPool(ocxname, content)
                this.special_files[filePath + "/content.json"] = obj.toJSON()
                this.tree_structure[filePath + "/content.json"] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
            }
        }
        for (let i = 0; i < pths.streams.length; i++) {
            let pth = pths.streams[i].toString()
            paths.push([pth, vscode.FileType.File])

            let filePath = uri== "" ? pth :  uri + "/" + pth
            this.tree_structure[filePath] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
            if (pth.endsWith(".bin")) {
                paths.push([filePath + "$", vscode.FileType.Directory])
                this.processBinaryFile(filePath)
                this.tree_structure[filePath + "$"] = { type: vscode.FileType.Directory, ctime:0, mtime:0, size: 0 }
            } else if (isVbaDirStream(filePath)) {
                paths.push([filePath + ".json", vscode.FileType.Directory])
                let project_file = await this.doc.read_file(filePath)
                this.dir_stream = VbaDirStream.from_buffer(project_file);
                this.special_files[filePath + ".json"] = JSON.stringify(this.dir_stream,null,"\t")
                this.tree_structure[filePath + ".json"] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
            } else if (isVbaModule(filePath)) {
                paths.push([filePath + ".vba", vscode.FileType.File])
                paths.push([filePath + ".pcode", vscode.FileType.File])
                if(this.dir_stream){
                    let project_file = await this.doc.read_file(filePath)
                    let module_offset = this.dir_stream.references_record.modules.find(val => val.name == pth);
                    let md = VbaModule.from_buffer(project_file,module_offset ? module_offset.source_offset : 0)
                    this.special_files[filePath + ".vba"] = md.source_code
                    this.special_files[filePath + ".pcode"] = md.performance_cache
                }else{
                    this.special_files[filePath + ".vba"] = null
                    this.special_files[filePath + ".pcode"] = null
                }
                this.tree_structure[filePath + ".vba"] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
                this.tree_structure[filePath + ".pcode"] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
            } else if (isDocumentSummaryInformation(filePath)) {
                paths.push([filePath + ".json", vscode.FileType.File])
                let project_file = await this.doc.read_file(filePath)
                this.special_files[filePath + ".json"] = JSON.stringify(DocumentSummaryInformation.from_buffer(project_file), null, "\t")
                this.tree_structure[filePath + ".json"] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
            } else if (isSummaryInformation(filePath)) {
                paths.push([filePath + ".json", vscode.FileType.File])
                let project_file = await this.doc.read_file(filePath)
                this.special_files[filePath + ".json"] = JSON.stringify(SummaryInformation.from_buffer(project_file), null, "\t")
                this.tree_structure[filePath + ".json"] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
            } else if (isBiffDocument(filePath)) {
                paths.push([filePath + ".json", vscode.FileType.File])
                let project_file = await this.doc.read_file(filePath)
                this.special_files[filePath + ".json"] = JSON.stringify(BiffDocument.from_buffer(project_file), null, "\t")
                this.tree_structure[filePath + ".json"] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
            } else if (isVbaProjectStream(filePath)) {
                paths.push([filePath + ".json", vscode.FileType.File])
                let project_file = await this.doc.read_file(filePath)
                this.special_files[filePath + ".json"] = JSON.stringify(VbaProjectStream.from_buffer(project_file), null, "\t")
                this.tree_structure[filePath + ".json"] = { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 }
            }else if(uri.startsWith("ObjectPool/")){
                paths.push([uri + "/content.json", vscode.FileType.File])
            }
        }
        return paths
    }

    async processBinaryFile(uri = "") {
        let bin_prj = this.doc.read_file(uri)
        this.elements[uri + "$"] = await OfficeFileFS.from_buffer(bin_prj);
    }
}

function isVbaDirStream(uri = "") {
    if (uri.endsWith("/dir")) {
        return true
    }
    return false
}
function isVbaProjectStream(uri = "") {
    if (uri.includes("/_VBA_PROJECT")) {
        return true
    }
    return false
}

function isVbaModule(uri = "") {
    if (uri.startsWith("Macros/VBA/")) {
        if (uri.endsWith("/dir")) return false
        if (uri.includes("/_VBA_PROJECT")) return false
        if (uri.includes("/__SRP_")) return false

        return true
    }
    if (uri.includes("/VBA/")) {
        if (uri.endsWith("/dir")) return false
        if (uri.includes("/_VBA_PROJECT")) return false
        if (uri.includes("/__SRP_")) return false

        return true
    }
    return false
}
function isDocumentSummaryInformation(uri = "") {
    if (uri.includes("DocumentSummaryInformation")) {
        return true
    }
    return false
}
function isSummaryInformation(uri = "") {
    if (uri.includes("SummaryInformation")) {
        return true
    }
    return false
}
function isBiffDocument(uri = "") {
    if (uri == "Workbook") {
        return true
    }
    return false
}

exports.VSCodeOfficeFS = VSCodeOfficeFS