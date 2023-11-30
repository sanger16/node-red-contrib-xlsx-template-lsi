module.exports = function(RED) {
    const fs = require('fs');
    var xlsxTemplate = require('xlsx-template');

    function xlsxTemplateLsi(config) {
        RED.nodes.createNode(this,config);
        const node = this;
        node.on('input', function(msg) {
            let reportData = msg.reportdata;
            //var readXlsx = msg.payload;
            let objectData = JSON.parse(msg.payload);
            let pathTemplate = msg.template;
            let pathReport = msg.report;
            /* Reporta errores */
            if (!objectData)
                node.error('Data in msg.payload is undefined!', msg.payload);
            if (!pathTemplate)
                node.error('Please define path to the template in msg.template', msg.template);
            if (!pathReport)
                node.error('Please provide the path to save the report in msg.report', msg.report);

            /*  Load an XLSX file into memory  */
            fs.readFile(pathTemplate,  (err, data) => {

                /*  Create a template */
                let template = new xlsxTemplate(data);

                /*  Replacements take place on first sheet */
                let sheetNumber = 1;           

                /*  Perform substitution */
                template.substitute(sheetNumber, objectData);

                /*  Get binary data */
                let genData = template.generate();

                /* Save file */
                fs.writeFileSync(pathReport, genData, "binary");

            });
            node.send(msg);
        });
    }
    RED.nodes.registerType("xlsx-template-lsi",xlsxTemplateLsi);
}