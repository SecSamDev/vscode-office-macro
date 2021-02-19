const vscode = require('vscode');
const fs = require('fs');
const path = require('path')
const { MultiFileAnalyzer } = require('./analyzer')
const { OfficeFileFS } = require('./office/office-fs')
const { OfficeAnalysisResults } = require('./office/office-results')
const { VbaProjectStream, VbaDirStream, VbaModule } = require('./office/office-parser')

const { OleCompoundDoc } = require('./office/ole-doc')
const cache = {};
let macroLab = vscode.window.createOutputChannel("MacroLab");
async function activate(context) {
    const fileSystemProvider = new MacroLabProvider();
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('doc', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('xls', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('ppt', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));
    /*
    context.subscriptions.push(vscode.workspace.onDidOpenTextDocument(document => tryPreviewDocument(document)));
    if (vscode.window.activeTextEditor !== undefined) {
        await tryPreviewDocument(vscode.window.activeTextEditor.document);
    }*/
    context.subscriptions.push(vscode.commands.registerCommand('macrolab.openOfficeDocument', async (resource) => {
        await tryPreviewDocument({ uri: resource })
    }));
}

module.exports.activate = activate

async function tryPreviewDocument(document) {
    macroLab.appendLine("Preview document: " + document.uri.toString())
    let name = path.basename(document.uri.path);
    let extension = path.extname(document.uri.path).substr(1).toLowerCase();

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
       //Create output channel
       

       //Write to output.
       orange.appendLine("I am a banana.");
    let html = await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: `Parsing ${name}` }, async () => {
        try {
            macroLab.appendLine("Pre-reading: " + document.uri.path)
            let doc_stream = fs.readFileSync(document.uri.path)
            macroLab.appendLine("Reading OK")
            let analyzer = await MultiFileAnalyzer.from_buffer(doc_stream)
            await analyzer.analyze()
            cache[document.uri.toString()] = analyzer.analyzer;
            let toRet = analyzer.analyzer.report.static.macros.reduce((pv, val) => { pv += "'" + val.name + "\n" + val.code; return pv }, "");
            return Buffer.from(toRet)

        } catch (e) {
            macroLab.appendLine(e)
            vscode.window.showInformationMessage(`Failed to parse ${name}!`);
            return Buffer.from("No valid Office document").toString("utf-8")
        }

    });

    const documentUri = vscode.Uri.parse(`${extension}:/?${document.uri}`);
    macroLab.appendLine(documentUri.toString())
    if (vscode.workspace.getWorkspaceFolder(documentUri) === undefined) {
        vscode.workspace.updateWorkspaceFolders(vscode.workspace.workspaceFolders?.length || 0, 0, { uri: documentUri, name });
    }



    const webviewPanel = vscode.window.createWebviewPanel(name, name, vscode.ViewColumn.Active);
    webviewPanel.webview.html = html;
}


class MacroLabProvider {
    constructor(...args) {
        this.emitter = new vscode.EventEmitter();
        this.onDidChangeFile = this.emitter.event;

    }

    watch(_uri_options) {
        return new vscode.Disposable(() => { });
    }
    stat(uri) {
        const fileUri = vscode.Uri.parse(uri.query);
        const office = cache[fileUri.toString()];
        const name = path.basename(fileUri.path);
        if (!office) {
            const error = `${name} was not found in cache! ${Object.keys(cache)}`;
            vscode.window.showErrorMessage(error);
            throw new Error(error);
        }

        const { ctime, mtime } = office;

        if (uri.path === '/') {
            return { type: vscode.FileType.Directory, ctime, mtime, size: 0 };
        }
        if (/\..{1,4}$/.test(uri.path)) {
            if (uri.path.endsWith('.bin')) {
                return { type: vscode.FileType.File, ctime, mtime, size: 0 };
            }
            return { type: vscode.FileType.File, ctime, mtime, size: 0 };
        }
        if (uri.path == "/1Table" || uri.path == "/Data" || uri.path.includes("/PROJECT") || uri.path == "/WordDocument") {
            return { type: vscode.FileType.File, ctime, mtime, size: 0 };
        }
        if (uri.path.startsWith("/Macros/VBA/") || uri.path.startsWith("/")) {
            return { type: vscode.FileType.File, ctime, mtime, size: 0 };
        }
        if (uri.path.includes('/contents') || uri.path.includes('OCXNAME') || uri.path.includes('ObjInfo') || uri.path.includes('PRINT') || uri.path.includes('__SRP_') || uri.path.endsWith('VBFrame') || uri.path.endsWith('CompObj') || uri.path.endsWith('/f') || uri.path.endsWith('/o') || uri.path.includes('/_VBA_PROJECT') || uri.path.endsWith('/dir') || uri.path.endsWith('/Workbook')) {
            return { type: vscode.FileType.File, ctime, mtime, size: 0 };
        }
        if (uri.path.endsWith('.bin')) {
            return { type: vscode.FileType.Directory, ctime, mtime, size: 0 };
        }
        return { type: vscode.FileType.Directory, ctime, mtime, size: 0 };

    }

    async load_cache(uri){
        try {
            let doc_stream = fs.readFileSync(uri.path)
            let analyzer = await MultiFileAnalyzer.from_buffer(doc_stream)
            await analyzer.analyze()
            cache[uri.toString()] = analyzer.analyzer;
            let toRet = analyzer.analyzer.report.static.macros.reduce((pv, val) => { pv += "'" + val.name + "\n" + val.code; return pv }, "");
            return Buffer.from(toRet)

        } catch (e) {
            macroLab.appendLine(e)
            vscode.window.showInformationMessage(`Failed to parse ${name}!`);
            return Buffer.from("No valid Office document").toString("utf-8")
        }
    }

    createDirectory(_uri) {
        const error = 'createDirectory should not be called';
        vscode.window.showErrorMessage(error);
        throw new Error(error);
    }
    async readFile(uri) {
        const officeUri = vscode.Uri.parse(uri.query);
        let office = cache[officeUri.toString()];
        const name = path.basename(officeUri.path);
        if (!office) {
            const error = `${name} was not found in cache! ${Object.keys(cache)}`;
            vscode.window.showErrorMessage(error);
            this.load_cache(officeUri)
            office = cache[officeUri.toString()];
            throw new Error(error);
        }
        if (uri.path === '/Macros/VBA/dir.json') {
            return Buffer.from(JSON.stringify(office.project_file, null, "\t"))
        }
        if (uri.path === '/DocumentSummaryInformation') {
            if (office.doc_summaryinfo) {
                return Buffer.from(JSON.stringify(office.doc_summaryinfo, null, "\t"))
            }
            return Buffer.from("")
        }
        if (uri.path === '/SummaryInformation') {
            if (office.summaryinfo) {
                return Buffer.from(JSON.stringify(office.summaryinfo, null, "\t"))
            }
            return Buffer.from("")
        }
        if (uri.path === '/Workbook') {
            if (office.workbook) { 
                return Buffer.from(JSON.stringify(office.workbook, null, "\t"))
            }
        }
        if (uri.path === '/Workbook.js') {
            if (office.workbookjs) {
                return Buffer.from(office.workbookjs)
            }
        }
        if (uri.path.includes(".bin")) {
            if (uri.path.includes(".bin/")) {
                return Buffer.from(office.workbookjs)
            } else {

            }
        }
        if (uri.path.startsWith('/Macros/VBA/')) {
            if (uri.path.endsWith(".vba")) {
                let macro_name = path.basename(uri.path, ".vba");
                let macro = office.report.static.macros.find((val) => val.name == macro_name);
                if (macro) {
                    return Buffer.from(macro.code)
                }

            } else if (uri.path.endsWith(".pcode")) {
                let macro_name = path.basename(uri.path, ".pcode");
                let macro = office.report.static.macros.find((val) => val.name == macro_name);
                if (macro) {
                    return Buffer.from(macro.pcode)
                }
            }
        }
        let content = await office.doc.read_file(uri.path)
        return Buffer.from(content ? content : "")
    }

    readDirectory(uri) {
        let toAdd = []
        const officeUri = vscode.Uri.parse(uri.query);
        const office = cache[officeUri.toString()];
        if (!office) {
            const error = `${officeUri.toString()} was not found in cache! ${Object.keys(cache)}`;
            vscode.window.showErrorMessage(error);
            throw new Error(error);
        }
        let pth = uri.path.startsWith("/") ? uri.path.slice(1) : uri.path
        /*
        if (uri.path === '/' && office.report.static.macros.length > 0) {
            return [["Macros", vscode.FileType.Directory]]
        }
        if (uri.path === '/Macros') {
            return [["VBA", vscode.FileType.Directory]]
        }
        if (uri.path === '/Macros/VBA') {
            return [["dir.json", vscode.FileType.File], ...office.report.static.macros.map(val => [[val.name + ".vba", vscode.FileType.File], [val.name + ".pcode", vscode.FileType.File]]).reduce((pv, val) => { pv.push(val[0]); pv.push(val[1]); return pv }, [])]
        }
        if (uri.path.includes(".bin")) {
            //TODO:
            return [...office.report.static.macros.map(val => [[val.name + ".vba", vscode.FileType.File], [val.name + ".pcode", vscode.FileType.File]]).reduce((pv, val) => { pv.push(val[0]); pv.push(val[1]); return pv }, [])]
        }*/
        let toRet = office.doc.ls_dir2(pth)
        let storage = toRet.storage.map(val => [val.toString(), vscode.FileType.Directory])
        let streams = toRet.streams.map(val => [val.toString(), vscode.FileType.File])

        for (let i = 0; i < streams.length; i++) {
            if (streams[i][0] == "Workbook") {
                toAdd.push(["Workbook.js", vscode.FileType.File])
            }
        }
        if (uri.path === '/Macros/VBA') {
            toAdd = [["dir.json", vscode.FileType.File], ...office.report.static.macros.map(val => [[val.name + ".vba", vscode.FileType.File], [val.name + ".pcode", vscode.FileType.File]]).reduce((pv, val) => { pv.push(val[0]); pv.push(val[1]); return pv }, [])]
        }
        return [...storage, ...streams, ...toAdd]

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
