const { Connection, Statement, } = require('idb-pconnector');
const Excel = require('exceljs');

async function generateExcel() {
    const connection = new Connection({ url: '*LOCAL' });
    const statement = new Statement(connection);
    const sql = 'SELECT CUSNUM, LSTNAM, BALDUE, CDTLMT FROM QIWS.QCUSTCDT'
    const results = await statement.exec(sql);

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Customers');

    // Define columns in the worksheet, these columns are identified using a key.
    worksheet.columns = [
        { header: 'Id', key: 'CUSNUM', width: 10 },
        { header: 'Last Name', key: 'LSTNAM', width: 10 },
        { header: 'Balance Due', key: 'BALDUE', width: 11 },
        { header: 'Credit Limit', key: 'CDTLMT', width: 10 }
    ];

    // Add rows from database to worksheet 
    for (const row of results) {
        worksheet.addRow(row);
    }

    // Add autofilter on each column
    worksheet.autoFilter = 'A1:D1';

    // Process each row for calculations and beautification 
    worksheet.eachRow((row, rowNumber) => {

        row.eachCell((cell, colNumber) => {
            if (rowNumber == 1) {
                // First set the background of header row
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'f5b914' }
                };
            };
            // Set border of each cell 
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
        //Commit the changed row to the stream
        row.commit();
    });

    await workbook.xlsx.writeFile('SimpleFormatCust.xlsx');
}
// Call the generateExcel function
generateExcel().catch((error) => {
    console.error(error);
});