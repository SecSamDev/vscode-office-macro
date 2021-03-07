const vscode = require('vscode');
const path = require('path')
const {OfficeProvider, tryPreviewOfficeDocument} = require('./office-provider')

async function activate(context) {
    const fileSystemProvider = new OfficeProvider();
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('doc', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('xls', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));
    context.subscriptions.push(vscode.workspace.registerFileSystemProvider('ppt', fileSystemProvider, { isCaseSensitive: true, isReadonly: true }));

    context.subscriptions.push(vscode.commands.registerCommand('macrolab.openOfficeDocument', async (resource) => {
        let extension = path.extname(resource.fsPath).substr(1).toLowerCase();
        if (extension.includes("doc") || extension.includes("xls") || extension.includes("ppt")) {
            try{
                await tryPreviewOfficeDocument({ uri: resource })
            }catch(e){}
        }
        
        
    }));
}

module.exports.activate = activate
