const express = require('express');
const bodyParser = require('body-parser');
const excelJs = require('exceljs');

const app = express();
app.use(express.json());
app.use(bodyParser.json());

const port = 3000;

const server = app.listen(port, () => console.log("Server is listen to port ",server.address().port));

app.post("/sheet", async(req,res) =>{
    try{
        let data = req.body
        if(!data) return res.send({satus: 400, success: false, response: "Please enter data"});
        let str = data.parent + '<' + data.sub + '<' + data.child;
        const wb = new excelJs.Workbook();
        const ws = wb.addWorksheet("category data");
        ws.columns = [{
            header: 'id', key: 'id', width: 20
        },
        {
            header: 'data', key: 'data', width: 50
        }
    ];
    const generated_id = Date.now().toString(36) + Math.random().toString(36);
    const req_data = { id: generated_id, data: str }
    ws.addRow(req_data)
    ws.getRow(1).eachCell((cell) =>{
        cell.font = {bold: true}
    });

        const result = await wb.xlsx.writeFile("ddSheet.xlsx");
        res.send({status: 200, success: true, response: "Your data is inserted to excel sheet"});
    }
    catch(err){
        res.send({status: 500, success: false, response: err});
    }
});
