const vscode = require('vscode')
const {PdfFS} = require('./pdf/pdf-fs')
const {VSCodePdfFS} = require('./pdf/vsfs')
const path = require('path')
const fs = require('fs')

const cache = {};

//VSCode FS
class PdfProvider {
    constructor(...args) {
        this.emitter = new vscode.EventEmitter();
        this.onDidChangeFile = this.emitter.event;

    }
    
    stat(uri = "") {
        if(uri.fsPath.includes(".vscode")){
            return null
        }
        const fileUri = vscode.Uri.parse(uri.query);
        const pdf = cache[fileUri.toString()];
        const name = path.basename(fileUri.fsPath);
        if (!pdf) {
            const error = `${name} was not found in cache! ${Object.keys(cache)}`;
            vscode.window.showErrorMessage(error);
            throw new Error(error);
        }
        return pdf.stat(uri.fsPath)
    }

    static load_cache(uri) {
        try {
            let buff = fs.readFileSync(uri.fsPath)
            let pdfdoc = PdfFS.from_buffer(buff)
            cache[uri.toString()] = new VSCodePdfFS(pdfdoc)
            return Buffer.from("Loading PDF...")
        } catch (e) {
            vscode.window.showInformationMessage(`Failed to parse ${path.basename(uri.path)}!`);
            return null
        }
    }

    async readFile(uri = "") {
        if(uri.path.includes(".vscode")){
            return null
        }
        const pdfUri = vscode.Uri.parse(uri.query);
        let pdf = cache[pdfUri.toString()];
        const name = path.basename(pdfUri.path);
        if (!pdf) {
            PdfProvider.load_cache(pdfUri)
            pdf = cache[pdfUri.toString()];
            if (!pdf) {
                const error = `${name} was not found in cache! ${Object.keys(cache)}`;
                vscode.window.showErrorMessage(error);
                throw new Error(error);
            }
        }
        return await pdf.readFile(uri.path)
    }

    async readDirectory(uri = "") {
        let toAdd = []
        const pdfUri = vscode.Uri.parse(uri.query);
        const pdf = cache[pdfUri.toString()];
        if (!pdf) {
            const error = `${pdfUri.toString()} was not found in cache! ${Object.keys(cache)}`;
            vscode.window.showErrorMessage(error);
            throw new Error(error);
        }
        return await pdf.readDirectory(uri.path)
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


async function tryPreviewPdfDocument(document) {
    let name = path.basename(document.uri.fsPath);
    let extension = path.extname(document.uri.fsPath).substr(1).toLowerCase();

    if (!extension.includes("pdf")) {
        return
    }
    extension = "pdf"
    let html = await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: `Parsing ${name}` }, async () => {
        try {
            PdfProvider.load_cache(document.uri)
            return Buffer.from("Analisis of file in process...")


        } catch (e) {
            vscode.window.showInformationMessage(`Failed to parse ${name}!`);
            return Buffer.from("No valid PDF document. Parser incomplete").toString("utf-8")
        }

    });

    const documentUri = vscode.Uri.parse(`${extension}:/?${document.uri}`);
    if (vscode.workspace.getWorkspaceFolder(documentUri) === undefined) {
        vscode.workspace.updateWorkspaceFolders(vscode.workspace.workspaceFolders?.length || 0, 0, { uri: documentUri, name });
    }



    const webviewPanel = vscode.window.createWebviewPanel("Analysis Resume", name, vscode.ViewColumn.Active);
    webviewPanel.webview.html = html;
}

module.exports.PdfProvider = PdfProvider
module.exports.tryPreviewPdfDocument = tryPreviewPdfDocument