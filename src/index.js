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
//   async exportDraftBilling(billingId: string): Promise<string> {
//     try {
//       const billing = await this.repositories.billing.findUnique({
//         where: { id: billingId },
//         include: { billingDetails: true },
//       })

//       // Get tax value from eclaim
//       const taxFromEclaim = await this.eclaimService.listTaxTypes('INVOICE')
//       const tax = taxFromEclaim.find((e) => e.type === billing.taxType)

//       // Read template file synchronously
//       const data = fs.readFileSync(
//         join(process.cwd(), 'template/invoice', 'draft-invoice-export.xlsx')
//       )

//       // Create a template
//       const template = new XlsxTemplate(data)
//       // Replacements take place on first sheet
//       const sheetNumber = 1

//       // Set up some placeholder values matching the placeholders in the template
//       const values = {
//         legalName: billing.name,
//         vatAddress: billing.address,
//         unitPrice: billing.billingDetails[0].unitPrice,
//         taxCode: billing.taxCode,
//         description: billing.billingDetails[0].description,
//         amount: this.billAmount(
//           billing.billingDetails[0].unitPrice,
//           billing.billingDetails[0].quantity
//         ),
//         tax: tax.value,
//         vatTax: this.calculateVat(
//           this.billAmount(billing.billingDetails[0].unitPrice, billing.billingDetails[0].quantity),
//           tax.value
//         ),
//         totalPayable: this.calculateTotalAmount(
//           billing.billingDetails[0].unitPrice,
//           billing.billingDetails[0].unitPrice,
//           tax.value
//         ),
//         dateInvoice: formatDateWithSlashes(new Date()),
//       }

//       // Perform substitution
//       template.substitute(sheetNumber, values)
//       // Get binary data
//       const filledExcelData = template.generate({ type: 'nodebuffer' })

//       // Đường dẫn tới tệp Excel mới
//       const currentDate = new Date()
//       const formattedDate = currentDate.toISOString().slice(0, 19).replace(/[-T:]/g, '_')
//       const fileName = `draft-invoice-export-${formattedDate}.xlsx`
//       // const newExcelPath = path.join(process.cwd(), 'template/invoice', newExcelFileName)

//       // Write filled Excel data to new file
//       // fs.writeFileSync(newExcelPath, filledExcelData, 'binary')

//       // console.log('Đã tạo thành công tệp Excel mới:', newExcelPath)

//       const url = (
//         await this.fileStorageService.uploadFileFromMail(
//           [
//             {
//               fileBuffer: filledExcelData,
//               contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
//               filename: fileName,
//             },
//           ],
//           false
//         )
//       ).toString()
//       console.log('url:', url)
//       return url
//     } catch (err) {
//       console.log('err.....:', err)
//     }
//   }