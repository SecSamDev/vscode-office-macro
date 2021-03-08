const vscode = require('vscode');
const fs = require('fs');
const path = require('path')
const { VSCodeOfficeFS } = require("./office/vsfs")
const {OfficeFileFS} = require('./office/office-fs')
const cache = {};
class OfficeProvider {
    constructor(...args) {
        this.emitter = new vscode.EventEmitter();
        this.onDidChangeFile = this.emitter.event;

    }

    watch(_uri_options) {
        return new vscode.Disposable(() => { });
    }
    stat(uri) {
        if(uri.path.includes(".vscode")){
            return null
        }
        const fileUri = vscode.Uri.parse(uri.query);
        const office = cache[fileUri.toString()];
        const name = path.basename(fileUri.path);
        if (!office) {
            const error = `${name} was not found in cache! ${Object.keys(cache)}`;
            vscode.window.showErrorMessage(error);
            throw new Error(error);
        }
        return office.stat(uri.path)
    }

    async load_cache(uri) {
        try {
            let doc_stream = fs.readFileSync(uri.path)
            if(doc_stream.indexOf("{\\rtf") < 3) {
                vscode.window.showInformationMessage(`RTF document not supported by this extension`);
                return Buffer.from("RTF document not supported by this extension").toString("utf-8")
            }
            let office_fs = await OfficeFileFS.from_buffer(doc_stream);
            let vs_fs = new VSCodeOfficeFS(office_fs);
            //  Update content in "/"
            await vs_fs.readDirectory("/")
            cache[uri.toString()] = vs_fs
            return Buffer.from("Analisis of file in process...")

        } catch (e) {
            vscode.window.showInformationMessage(`Failed to parse ${path.basename(uri.path)}!`);
            return Buffer.from("No valid Office document").toString("utf-8")
        }
    }

    createDirectory(_uri) {
        const error = 'createDirectory should not be called';
        vscode.window.showErrorMessage(error);
        throw new Error(error);
    }
    async readFile(uri) {
        if(uri.path.includes(".vscode")){
            return null
        }
        const officeUri = vscode.Uri.parse(uri.query);
        let office = cache[officeUri.toString()];
        const name = path.basename(officeUri.path);
        if (!office) {
            this.load_cache(officeUri)
            office = cache[officeUri.toString()];
            if (!office) {
                const error = `${name} was not found in cache! ${Object.keys(cache)}`;
                vscode.window.showErrorMessage(error);
                throw new Error(error);
            }
        }
        return await office.readFile(uri.path)
    }

    async readDirectory(uri) {
        let toAdd = []
        const officeUri = vscode.Uri.parse(uri.query);
        const office = cache[officeUri.toString()];
        if (!office) {
            const error = `${officeUri.toString()} was not found in cache! ${Object.keys(cache)}`;
            vscode.window.showErrorMessage(error);
            throw new Error(error);
        }
        return await office.readDirectory(uri.path)
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



async function tryPreviewOfficeDocument(document) {
    let name = path.basename(document.uri.fsPath);
    let extension = path.extname(document.uri.fsPath).substr(1).toLowerCase();

    if (!extension.includes("doc") && !extension.includes("xls") && !extension.includes("ppt")) {
        return
    }
    if (extension.includes("doc")) {
        extension = "doc"
    }
    if (extension.includes("xls")) {
        extension = "xls"
    }
    if (extension.includes("ppt")) {
        extension = "ppt"
    }
    let html = await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: `Parsing ${name}` }, async () => {
        try {
            let doc_stream = fs.readFileSync(document.uri.fsPath)
            let office_fs = await OfficeFileFS.from_buffer(doc_stream);
            let vs_fs = new VSCodeOfficeFS(office_fs);
            //  Update content in "/"
            await vs_fs.readDirectory("/")
            cache[document.uri.toString()] = vs_fs
            return Buffer.from("Analisis of file in process...")


        } catch (e) {
            vscode.window.showInformationMessage(`Failed to parse ${name}!`);
            return Buffer.from("No valid Office document").toString("utf-8")
        }

    });

    const documentUri = vscode.Uri.parse(`${extension}:/?${document.uri}`);
    if (vscode.workspace.getWorkspaceFolder(documentUri) === undefined) {
        vscode.workspace.updateWorkspaceFolders(vscode.workspace.workspaceFolders?.length || 0, 0, { uri: documentUri, name });
    }



    const webviewPanel = vscode.window.createWebviewPanel("Analysis Resume", name, vscode.ViewColumn.Active);
    webviewPanel.webview.html = html;
}

module.exports.OfficeProvider = OfficeProvider
module.exports.tryPreviewOfficeDocument = tryPreviewOfficeDocument