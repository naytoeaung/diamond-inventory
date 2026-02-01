#!/bin/bash

echo "🔄 Updating Diamond Inventory from Excel..."
cd ~/Desktop/Sheet\ App

# Generate data.json from Excel
node << 'EOF'
const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('Sheet.xlsx');
const jsonData = [];

// Read Purchase sheet
if (workbook.Sheets['Purchase']) {
    const data = XLSX.utils.sheet_to_json(workbook.Sheets['Purchase'], {header: 1});
    for (let i = 2; i < data.length; i++) {
        const row = data[i];
        if (!row || !row[1]) continue;
        
        const qty = parseFloat(row[4]) || 0;
        const rate = parseFloat(row[5]) || 0;
        const disc = parseFloat(row[6]) || 0;
        const orig = qty * rate;
        
        jsonData.push({
            date: row[0], 
            shape: row[1],
            size: row[2] !== null ? String(row[2]) : null,
            perCt: row[3] !== null ? String(row[3]) : null,
            quantity: qty, 
            rate, 
            discountRate: disc, 
            type: 'Purchase',
            originalPrice: Math.round(orig),
            discountPrice: Math.round(orig * disc),
            finalPrice: Math.round(orig - (orig * disc))
        });
    }
}

// Read Sell sheet
if (workbook.Sheets['Sell']) {
    const data = XLSX.utils.sheet_to_json(workbook.Sheets['Sell'], {header: 1});
    for (let i = 2; i < data.length; i++) {
        const row = data[i];
        if (!row || !row[1]) continue;
        
        const perCt = parseFloat(row[3]) || 0;
        const qty = parseFloat(row[4]) || 0;
        const rate = parseFloat(row[5]) || 0;
        const adj = perCt * 1.1;
        
        jsonData.push({
            date: row[0], 
            shape: row[1],
            size: row[2] !== null ? String(row[2]) : null,
            perCt: row[3] !== null ? String(row[3]) : null,
            quantity: qty, 
            rate, 
            discountRate: 0, 
            type: 'Sell',
            originalPrice: Math.round(qty * rate),
            discountPrice: 0,
            finalPrice: Math.round(qty * adj * rate),
            changeUnit: 1.1, 
            adjustedPerCt: adj
        });
    }
}

fs.writeFileSync('data.json', JSON.stringify(jsonData, null, 2));
console.log('✓ Generated data.json with', jsonData.length, 'items');
EOF

# Commit and push
echo "📤 Pushing to GitHub..."
git add data.json
git commit -m "Update inventory data - $(date '+%Y-%m-%d %H:%M')"
git push

echo ""
echo "✅ Done! Your site will update in 1-2 minutes at:"
echo "🌐 https://gsjulian81-bot.github.io/diamond-inventory/"