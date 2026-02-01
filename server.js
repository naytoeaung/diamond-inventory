const express = require('express');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const PORT = 8080;

app.use(express.static(path.join(__dirname)));

const CHANGE_UNIT = 1.1; // Constant for sell calculation

app.get('/api/data', (req, res) => {
    try {
        const workbook = XLSX.readFile(path.join(__dirname, 'Sheet.xlsx'));
        const jsonData = [];
        
        // Read Purchase sheet
        if (workbook.Sheets['Purchase']) {
            const purchaseData = XLSX.utils.sheet_to_json(workbook.Sheets['Purchase'], {header: 1});
            
            for (let i = 2; i < purchaseData.length; i++) {
                const row = purchaseData[i];
                if (!row || !row[1]) continue;
                
                const date = row[0];
                const shape = row[1];
                const size = row[2];
                const perCt = row[3];
                const quantity = parseFloat(row[4]) || 0;
                const rate = parseFloat(row[5]) || 0;
                const discountRate = parseFloat(row[6]) || 0;
                
                const originalPrice = quantity * rate;
                const discountPrice = originalPrice * discountRate;
                const finalPrice = originalPrice - discountPrice;
                
                jsonData.push({
                    date,
                    shape,
                    size: size !== null ? String(size) : null,
                    perCt: perCt !== null ? String(perCt) : null,
                    quantity,
                    rate,
                    discountRate,
                    type: 'Purchase',
                    originalPrice: Math.round(originalPrice),
                    discountPrice: Math.round(discountPrice),
                    finalPrice: Math.round(finalPrice),
                    changeUnit: null // Only for sells
                });
            }
        }
        
        // Read Sell sheet
        if (workbook.Sheets['Sell']) {
            const sellData = XLSX.utils.sheet_to_json(workbook.Sheets['Sell'], {header: 1});
            
            for (let i = 2; i < sellData.length; i++) {
                const row = sellData[i];
                if (!row || !row[1]) continue;
                
                const date = row[0];
                const shape = row[1];
                const size = row[2];
                const perCt = row[3];
                const quantity = parseFloat(row[4]) || 0;
                const originalRate = parseFloat(row[5]) || 0;
                
                // Calculate sell price: (per ct * 1.1) * original rate * quantity
                const adjustedPerCt = parseFloat(perCt) * CHANGE_UNIT;
                const originalPrice = quantity * originalRate;
                const finalPrice = quantity * adjustedPerCt * originalRate;
                
                jsonData.push({
                    date,
                    shape,
                    size: size !== null ? String(size) : null,
                    perCt: perCt !== null ? String(perCt) : null,
                    quantity,
                    rate: originalRate,
                    discountRate: 0,
                    type: 'Sell',
                    originalPrice: Math.round(originalPrice),
                    discountPrice: 0,
                    finalPrice: Math.round(finalPrice),
                    changeUnit: CHANGE_UNIT,
                    adjustedPerCt: adjustedPerCt
                });
            }
        }
        
        res.json(jsonData);
    } catch (error) {
        console.error('Error reading Excel:', error);
        res.status(500).json({ error: 'Failed to read Excel file' });
    }
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    console.log('Change unit for sells:', CHANGE_UNIT);
});