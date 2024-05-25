const express = require('express');
const bodyParser = require('body-parser');
const fetch = require('node-fetch');
const app = express();

app.use(bodyParser.json());

const clientId = 'YOUR_CLIENT_ID';
const clientSecret = 'YOUR_CLIENT_SECRET';
const tenantId = 'YOUR_TENANT_ID';
const accessToken = 'YOUR_ACCESS_TOKEN'; // You need to implement OAuth2 to get this token

const excelFileId = 'YOUR_EXCEL_FILE_ID';
const excelSheetName = 'Sheet1';

app.post('/submit_form', async (req, res) => {
    const { name, email, message } = req.body;

    const data = [
        [name, email, message]
    ];

    try {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${excelFileId}/workbook/worksheets/${excelSheetName}/range(address='A1:C1')/insert`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                values: data
            })
        });

        if (response.ok) {
            res.json({ success: true });
        } else {
            res.json({ success: false });
        }
    } catch (error) {
        console.error('Error:', error);
        res.json({ success: false });
    }
});

app.listen(3000, () => {
    console.log('Server is running on port 3000');
});
