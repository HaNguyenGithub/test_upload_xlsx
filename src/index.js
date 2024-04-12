const { join } = require('path');
var fs = require('fs');
const path = require('path');
var XlsxTemplate = require('xlsx-template');

function exportInvoice() {

    // Load an XLSX file into memory
    fs.readFile(join(process.cwd(), 'src','invoice-export.xlsx'), function(err, data) {

        // Create a template
        var template = new XlsxTemplate(data);
        console.log("template....:", template)
        // Replacements take place on first sheet
        var sheetNumber = 1;

        // Set up some placeholder values matching the placeholders in the template
        var values = {
            invoiceDate: 110424,
            amount: 5000,
            legalName: 'Test',
            vatAddress: 'Hanoi',
            taxCode: 'HN001',
            description: 'test',
            };

        // Perform substitution
        template.substitute(sheetNumber, values);
        // Get binary data
         const filledExcelData = template.generate()

         // Đường dẫn tới tệp Excel mới
         const currentDate = new Date();
        const formattedDate = currentDate.toISOString().slice(0, 19).replace(/[-T:]/g, '_');
         const newExcelFileName = `invoice-export-${formattedDate}.xlsx`;
        const newExcelPath = path.join(process.cwd(), 'src', newExcelFileName); 
        template.substitute(sheetNumber, {});

    // Ghi dữ liệu nhị phân vào tệp Excel mới
    fs.writeFile(newExcelPath, filledExcelData, 'binary', function(err) {
        if (err) {
            console.error('Không thể ghi tệp Excel mới:', err);
            return;
        }
        console.log('Đã tạo thành công tệp Excel mới:', newExcelPath);
    });

        if(err){console.log(err) }
    });
  }
  exportInvoice()