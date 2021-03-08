const { basename } = require('path')
const vscode = require('vscode')
const {PdfFS} = require('./pdf-fs')

//VSCode FS
class VSCodePdfFS {
    /**
     * 
     * @param {PdfFS} doc 
     */
    constructor(doc) {
        this.doc = doc
    }
    
    stat(uri = "") {
        if(uri.includes(".vscode")){
            return null
        }
        if (uri == "/") {
            return { type: vscode.FileType.Directory, ctime:0, mtime:0, size: 0 };
        }
        if (uri == "/scripts") {
            return { type: vscode.FileType.Directory, ctime:0, mtime:0, size: 0 };
        }
        if (uri == "/objects") {
            return { type: vscode.FileType.Directory, ctime:0, mtime:0, size: 0 };
        }
        if (uri.startsWith("/scripts/")) {
            return { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 };
        }
        if (uri.startsWith("/objects/")) {
            return { type: vscode.FileType.File, ctime:0, mtime:0, size: 0 };
        }
        return { type: vscode.FileType.Unknown, ctime:0, mtime:0, size: 0 };
    }

    async readFile(uri = "") {
        if(uri.startsWith("/scripts/")){
            let id = Number(uri.slice(9).split(".")[0])
            return Buffer.from(this.doc.scripts[id])
        }
        if(uri.startsWith("/objects/")){
            let id = uri.slice(9).split(".")[0]
            return Buffer.from(JSON.stringify(this.doc.objects[id],null,"\t"))
        }
        return null
    }

    async readDirectory(uri = "") {
        if (uri == "/") {
            return [["/objects", vscode.FileType.Directory],["/scripts", vscode.FileType.Directory]]
        }
        if (uri == "/scripts") {
            let toRet = []
            for(let i = 0; i < this.doc.scripts.length; i++){
                toRet.push([i + ".js", vscode.FileType.File])
            }
            return toRet
        }
        if (uri == "/objects") {
            return Object.keys(this.doc.objects).map(val => [val + ".json", vscode.FileType.File])
        }
        return []
    }
    createDirectory(_uri) {
        const error = 'createDirectory should not be called';
        vscode.window.showErrorMessage(error);
        throw new Error(error);
    }

    writeFile(_uri, _content, _options) {
        const error = 'writeFile should not be called';
        vscode.window.showErrorMessage(error);
        throw new Error(error);
    }

    delete(_uri, _options) {
        const error = 'writeFile should not be called';
        vscode.window.showErrorMessage(error);
        throw new Error(error);
    }

    rename(_oldUri, _newUri, _options) {
        const error = 'rename should not be called';
        vscode.window.showErrorMessage(error);
        throw new Error(error);
    }
}

exports.VSCodePdfFS = VSCodePdfFS