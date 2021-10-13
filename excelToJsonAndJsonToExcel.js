import { LightningElement, track } from 'lwc';
import PARSER from '@salesforce/resourceUrl/ExcelParser';
import CODEMIRROR from '@salesforce/resourceUrl/Codemirror';
import { loadScript, loadStyle } from 'lightning/platformResourceLoader';
import getContactLists from '@salesforce/apex/ExcelController.getContactLists';
import getAccountLists from '@salesforce/apex/ExcelController.getAccountLists';

const CONTACT_COLUMNS = [
    { label: 'Name', fieldName: 'Name' },
    { label: 'Email', fieldName: 'Email' },
    { label: 'Phone', fieldName: 'Phone' },
]

const ACCOUNT_COLUMNS = [
    { label: 'Name', fieldName: 'Name' },
    { label: 'Type', fieldName: 'Type' },
    { label: 'Industry', fieldName: 'Industry' },
]
export default class ExcelToJsonAndJsonToExcel extends LightningElement {

    @track placeholderJson = 'Output Here';
    @track accountData = []; // used only for storing account table
    @track contactData = []; // used only for storing contact table
    @track contactDataTableColumns = CONTACT_COLUMNS;
    @track accountDataTableColumns = ACCOUNT_COLUMNS;

    editor;
    isFirstTime = true;
    xlsHeader = []; // store all the headers of the the tables
    workSheetNameList = []; // store all the sheets name of the the tables
    xlsData = []; // store all tables data
    filename = "sample_demo.xlsx"; // Name of the file


    connectedCallback() {
        this.siteURL = '/apex/jsonToExcelParserPage';

        //apex call for bringing the contact data  
        getContactLists()
            .then(result => {
                let contacts = [];
                result.forEach((currentItem, index) => {
                    let element = Object.assign({}, currentItem);
                    contacts.push(element);
                });
                this.contactHeader = Object.keys(contacts[0]);
                this.contactData = [...this.contactData, ...contacts];
                this.xlsFormatter(contacts, "Contacts");
            })
            .catch(error => {
                console.error(error);
            });

        //apex call for bringing the account data  
        getAccountLists()
            .then(result => {
                this.accountHeader = Object.keys(result[0]);
                this.accountData = [...this.accountData, ...result];
                this.xlsFormatter(result, "Accounts");
            })
            .catch(error => {
                console.error(error);
            });
    }

    // formating the data to send as input to  xlsxMain component
    xlsFormatter(data, sheetName) {
        let Header = Object.keys(data[0]);
        this.xlsHeader.push(Header);
        this.workSheetNameList.push(sheetName);
        this.xlsData.push(data);
        console.log('this.xlsData  :-  ', this.xlsData);
    }

    renderedCallback() {
        if (this.isFirstTime) {
            this.isFirstTime = false;
            Promise.all([
                loadScript(this, CODEMIRROR + '/codemirror/lib/codemirror.js'),
                loadScript(this, CODEMIRROR + '/codemirror/mode/htmlmixed/htmlmixed.js'),
                loadScript(this, PARSER + '/exceltest/xlsx.full.min.js'),
                loadStyle(this, CODEMIRROR + '/codemirror/lib/codemirror.css')
            ]).then(() => {
                console.log('result');
            })
                .catch(error => {
                    console.log('error1  :-  ' + error);
                });
            setTimeout(function () {
                this.editor = CodeMirror.fromTextArea(this.template.querySelector('.editor'), {
                    lineNumbers: true,
                    mode: "htmlmixed",
                    // readOnly: 'nocursor'
                });
                this.editor.getDoc().setValue('Output Here...');
            }.bind(this), 3000)
        }
    }

    // This code convert excel in json
    ExcelToJSON(file) {
        var reader = new FileReader();
        let test = {};

        reader.onload = event => {
            var data = event.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });

            workbook.SheetNames.forEach(function (sheetName) {
                let sheet = workbook.Sheets[sheetName];
                var range = XLSX.utils.decode_range(sheet['!ref']);
                var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                var data = JSON.stringify(XL_row_object);
                test[sheetName] = JSON.parse(data);

            })
        };
        reader.onerror = function (ex) {
            this.error = ex;
            this.dispatchEvent(
                new ShowToastEvent({
                    title: 'Error while reding the file',
                    message: ex.message,
                    variant: 'error',
                }),
            );
        };
        setTimeout(function () {
            this.editor.getDoc().setValue(JSON.stringify(test, null, 2));

        }.bind(this), 3000)
        reader.readAsBinaryString(file);
    }

    // This code for none formatted excel download
    download() {
        const XLSX = window.XLSX;
        let xlsData = this.xlsData;
        let xlsHeader = this.xlsHeader;
        let ws_name = this.workSheetNameList;
        let createXLSLFormatObj = Array(xlsData.length).fill([]);

        /* form header list */
        xlsHeader.forEach((item, index) => {
            createXLSLFormatObj[index] = [item]
        })

        /* form data key list */
        xlsData.forEach((item, selectedRowIndex) => {
            let xlsRowKey = Object.keys(item[0]);
            item.forEach((value, index) => {
                var innerRowData = [];
                xlsRowKey.forEach(item => {
                    innerRowData.push(value[item]);
                })
                createXLSLFormatObj[selectedRowIndex].push(innerRowData);
            })
        });

        /* creating new Excel */
        var wb = XLSX.utils.book_new();

        /* creating new worksheet */
        var ws = Array(createXLSLFormatObj.length).fill([]);
        console.log('ws  :-  ', JSON.parse(JSON.stringify(ws)))
        for (let i = 0; i < ws.length; i++) {
            /* converting data to excel format and puhing to worksheet */
            let data = XLSX.utils.aoa_to_sheet(createXLSLFormatObj[i]);
            ws[i] = [...ws[i], data];
            /* Add worksheet to Excel */
            XLSX.utils.book_append_sheet(wb, ws[i][0], ws_name[i]);
        }
        XLSX.writeFile(wb, this.filename);
    }

    handleFileChange(event) {
        this.ExcelToJSON(event.target.files[0]);
    }

    //this code for formatted excel download
    downloadFormattedExcel() {
        let data = [
            [Object.keys(this.contactData[0])]
        ]
        this.contactData.forEach(element => {
            data[0].push(Object.values(element));
        });

        data.push([Object.keys(this.accountData[0])])
        this.accountData.forEach(element => {
            data[1].push(Object.values(element));
        });

        let sheetDetails = {
            sheetNames: ['Contacts', 'Accounts'],
            sheets: data
        }

        let baseUrl = window.location.origin
        baseUrl = baseUrl.split('.');
        let vfPageUrl = `${baseUrl[0]}--c.visualforce.com`;
        this.template.querySelector("iframe").contentWindow.postMessage(sheetDetails, vfPageUrl);
    }
}