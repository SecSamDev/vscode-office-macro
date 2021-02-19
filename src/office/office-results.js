

/**
 * {
                uuid : "1234567-123456",
                doc_name : this.doc_name,
                hash : {
                    md5 : "12312312",
                    sha1 : "asdada",
                },
                score : 80,
                techniques : [
                    {
                        type : "VBA Stomping",
                        value : "Founded in Macro1"
                    },
                    {
                        type : "Code Obfuscation",
                        value : "Founded in Macro1"
                    }
                ],
                static : {
                    macros : [
                        {
                            name : "Macro1",
                            code : "Codigo \n macro \n VBA",
                            pcode : "Codigo \n macro\n pcode"
                        }
                    ],
                    commands : [
                        {
                            type : "powershell",
                            value : "powershel -h -en asdadaaff7as9f7asf9"
                        }
                    ],
                    binaries : [],
                    iocs : [
                        {
                            type : "IP",
                            value : "192.1.168.100"
                        }
                    ],
                    //Typical functions used by malware
                    functions : [

                    ]
                },
                dynamic : {
                    ole_objects : [],
                    vars : [],
                    //Typical functions used by malware
                    functions : [],
                    offuscation_calls : [],
                    statistics : {
                        vars : {
                            used : 3,
                            declared : 100,
                            assigned : 20
                        }
                    },
                    commands : [],
                    iocs : [
                        {
                            type : "URL",
                            value : "http://192.1.168.100/malware.php"
                        }
                    ]
                }
 */

class OfficeAnalysisResults{
    constructor(){
        this.uuid = Math.random() + "-" + Date.now()
        this.hash = {
            md5 : "",
            sha1 : "",
        }
        this.score = 0
        /**
         * Techniques used by malware
         * @type {[{type : String,value : String}]}
         */
        this.techniques = []
        this.static = {
            /**
             * Macro data
             * @type {[{name : String, code : String, pcode : String}]}
             */
            macros : [
            ],
            /**
             * Commands extracted from macro
             * @type {[{type : String,value : String}]}
             */
            commands : [
            ],
            /**
             * Binary files stored in document
             * @type {[{name : String,value : String}]}
             */
            binaries : [],
            /**
             * IOCs extracted from macro
             * @type {[{type : String,value : String}]}
             */
            iocs : [
            ],
            /**
             * Functions used by macros
             * @type {[{name : String,value : String}]}
             */
            functions : [
            ],
            /**
             * Vars injected by macro
             * @type {[{name : String,value : String}]}
             */
            vars : [],
        }

        this.dynamic = {
            /**
             * Objects used by macro and not stored inside a macro (FormUser)
             * @type {[{name : String,value : String}]}
             */
            ole_objects : [],
            /**
             * Vars injected by macro
             * @type {[{name : String,value : String}]}
             */
            vars : [],
            /**
             * Functions used by the macro
             * @type {[{name : String,value : String}]}
             */
            functions : [],
            /**
             * Functions used by the macro to obfuscate the calls
             * @type {[{name : String,value : String, args : String[]}]}
             */
            offuscation_calls : [],
            statistics : {
                vars : {
                    writed : 0,
                    readed : 0,
                    not_used : 0
                },
                comments : 0
            },
            /**
             * Commands executed by the macro
             * @type {[{type : String,value : String}]}
             */
            commands : [],
            /**
             * IOCs extracted from the execution of the macro
             * @type {[{type : String,value : String}]}
             */
            iocs : [
            ]
        },
        /**
         * Errors processing the file
         * @type {string[]}
         */
        this.errors = []
    }
    /**
     * 
     * @param {{type : String,value : String}} tech Techniques detected in the document 
     */
    setTechnique(tech){
        if(!!tech.type && !!tech.value){
            tech.severity = !!tech.severity ? tech.severity : 'LOW';
            this.techniques.push(tech)
        }
        
    }
    getTechnique(name){
        return this.techniques.filter((val)=>val.type === name)
    }
}

module.exports.OfficeAnalysisResults = OfficeAnalysisResults