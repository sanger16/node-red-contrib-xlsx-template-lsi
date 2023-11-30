# Node-red xlsx Template

This node take an xlsx file as template and generate a report based on this file.

This node uses the [xlsx-template](https://www.npmjs.com/package/xlsx-template) package by optilude.

## Inputs

The node receives an object stringified (i.e. `JSON.stringify(object)`) in `msg.payload` which contain all data to fill the sheet 1 of the xlsx file template to generate the report.

The object has to contain all data that is configured in the tamplate file as placeholder as decribed in the documentation of the original module xlsx-template https://www.npmjs.com/package/xlsx-template?activeTab=readme.

 It also receives the full path to the xlsx template (including the name of the file .xlsx) in `msg.template`.

```/home/user/templates/template_name.xlsx```

and the full path of the report where (including the report name .xlsx) `msg.report`

```/home/user/reports/report_name.xlsx```