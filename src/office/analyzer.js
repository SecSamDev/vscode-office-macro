const { OfficeFileFS } = require('./office-fs')
const { OfficeAnalysisResults } = require('./office-results')
const { VbaProjectStream, VbaDirStream, VbaModule, DocumentSummaryInformation, SummaryInformation } = require('./office-parser')
const { OleCompoundDoc } = require('./ole-doc')
const { BiffDocument } = require('./office-biff')
const OFFICE_TYPE_UNKNOWN = 0
const OFFICE_TYPE_WORD = 1
const OFFICE_TYPE_EXCEL = 2
const OFFICE_TYPE_POWERPOINT = 3
const OFFICE_TYPE_OUTLOOK = 4
const OFFICE_TYPE_PUBLISHER = 5


/**
 * Analizador de documentos de office
 */
class OfficeAnalyzer {
    /**
     * 
     * @param {OfficeFileFS} doc 
     */
    constructor(doc) {
        this.doc = doc
        this.office_type = null
        this.report = new OfficeAnalysisResults()
        this.project_file = null
        this.sub_ole = {}
        this.doc_summaryinfo = null
    }

    async analyze() {
        //First detect
        this.getOfficeType()
        await this._extract_macros()
        if (this.report.static.macros.length > 0) {
            //Has macros -> Execute subanalyzers for macros
        }

    }
    async _extract_macros() {
        let root_dir = this.doc.ls_dir('')
        if (root_dir.includes('Macros')) {
            //Extract macros source code
            let files_macros = this.doc.ls_dir('Macros')
            if (files_macros.length > 0) {
                let project_file = await this.doc.read_file('Macros/VBA/dir')
                this.project_file = VbaDirStream.from_buffer(project_file)
                for (let i = 0; i < this.project_file.references_record.modules.length; i++) {
                    let module = this.project_file.references_record.modules[i]
                    try {
                        let vba_module = await this.doc.read_file('Macros/VBA/' + module.stream)
                        let vba_mod = VbaModule.from_buffer(vba_module, module.source_offset)
                        this.report.static.macros.push({
                            code: vba_mod.source_code,
                            name: module.name,
                            pcode: vba_mod.performance_cache
                        })
                    } catch (err) {
                        this.report.errors.push(`Error extracting VBA module ${i}: ${module.name}`)
                    }
                }

            }
        } else {
            //Search for .bin project
            let binProjects = root_dir.filter((val) => {
                return (val.indexOf(".bin") > 0)
            })
            if (binProjects.length > 0) {
                for (let i = 0; i < binProjects.length; i++) {
                    try {
                        let bin_prj = await this.doc.read_file(binProjects[i])
                        let ole_doc = new OleCompoundDoc(bin_prj)
                        await ole_doc.read();
                        this.sub_ole[binProjects[i]] = ole_doc
                    } catch (err) {
                        this.report.errors.push(`Error reading .bin file that could contain MACROS: ${binProjects[i]}`)
                    }
                }
            } else {
                //No macros
                
                
            }
        }
        if(root_dir.includes("DocumentSummaryInformation")){
            let doc_summaryinfo = await this.doc.read_file('DocumentSummaryInformation')
            this.doc_summaryinfo = DocumentSummaryInformation.from_buffer(doc_summaryinfo)
        }
        if(root_dir.includes("SummaryInformation")){
            let doc_summaryinfo = await this.doc.read_file('SummaryInformation')
            this.summaryinfo = SummaryInformation.from_buffer(doc_summaryinfo)
        }
        if(root_dir.includes("Workbook")){
            let workbook_stream = await this.doc.read_file('Workbook')
            this.workbook = BiffDocument.from_buffer(workbook_stream)
            this.workbookjs = this.workbook.generate_javascript();
        }
    }

    async _analyze_word() {

    }

    getReport() {
        return this.report;
    }

    getOfficeType() {
        if (this.office_type) {
            return this.office_type
        } else {
            let files_root = this.doc.ls_dir('')
            if (files_root.includes('WordDocument') || files_root.includes('wl')) {
                this.office_type = OFFICE_TYPE_WORD
            } else if (files_root.includes('xl')) {
                this.office_type = OFFICE_TYPE_EXCEL
            } else if (files_root.includes('ppt')) {
                this.office_type = OFFICE_TYPE_POWERPOINT
            }
            this.office_type = OFFICE_TYPE_UNKNOWN
        }
        return this.office_type
    }

}


exports.OfficeAnalyzer = OfficeAnalyzer