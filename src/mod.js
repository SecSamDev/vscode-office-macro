const vscode = require('vscode');
const path = require('path')
const {OfficeProvider, tryPreviewOfficeDocument} = require('./office-provider')
//This only works if the text is well formatted. TODO: improve parser
const {PdfProvider, tryPreviewPdfDocument} = require('./pdf-provider')

    
async function activate(context) {
    const fileSystemProvider = new OfficeProvider();
    const pdfFS = new PdfProvider();
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('doc', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('xls', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('ppt', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('pdf', pdfFS, { isCaseSensitive: true, isReadonly: true }));

    context.subscriptions.push(vscode.commands.registerCommand('macrolab.openOfficeDocument', async (resource) => {
        let extension = path.extname(resource.fsPath).substr(1).toLowerCase();
        if (extension.includes("doc") || extension.includes("xls") || extension.includes("ppt")) {
            try{
                await tryPreviewOfficeDocument({ uri: resource })
            }catch(e){}
        }
    }));
    context.subscriptions.push(vscode.commands.registerCommand('macrolab.openPdfDocument', async (resource) => {
        let extension = path.extname(resource.fsPath).substr(1).toLowerCase();
        if (extension.includes("pdf")) {
            try{
                await tryPreviewPdfDocument({ uri: resource })
            }catch(e){
                console.log(e)
            }
        }
    }));
}

module.exports.activate = activate
