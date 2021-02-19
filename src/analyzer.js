const { OfficeAnalyzer } = require("./office/analyzer");
const { OfficeFileFS } = require('./office/office-fs')
const TYPE_OFFICE = 0
const TYPE_PDF = 1
const TYPE_EML = 2

class MultiFileAnalyzer {

    static async from_buffer(buffer) {
        this.type = TYPE_OFFICE;
        switch (this.type) {
            case TYPE_OFFICE:
                let of_fs = await OfficeFileFS.from_buffer(buffer);
                return new MultiFileAnalyzer({
                    analyzer: new OfficeAnalyzer(of_fs)
                })

            default:
                break;
        }
    }
    /**
     * 
     * @param {{analyzer : OfficeAnalyzer}} data 
     */
    constructor(data) {
        //DetectFile type
        this.analyzer = data.analyzer

    }
    async analyze() {
        return await this.analyzer.analyze()
    }

    getReport() {
        return this.analyzer.getReport()
    }

}

exports.MultiFileAnalyzer = MultiFileAnalyzer