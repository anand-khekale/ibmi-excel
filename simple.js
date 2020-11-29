const { Connection, Statement, } = require('idb-pconnector');
const Excel = require('exceljs');

async function generateExcel() {
    // Create connection with DB2
    const connection = new Connection({ url: '*LOCAL' });
    const statement = new Statement(connection);
    const sql = 'SELECT CUSNUM, LSTNAM, BALDUE, CDTLMT FROM QIWS.QCUSTCDT'

    // Execute the statement to fetch data in results
    const results = await statement.exec(sql);

    // Create Excel workbook and worksheet
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

    // Finally save the worksheet into the folder from where we are running the code. 
    await workbook.xlsx.writeFile('SimpleCust.xlsx');
}

generateExcel().catch((error) => {
    console.error(error);
});