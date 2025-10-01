document.addEventListener('DOMContentLoaded', () => {
    // --- CONFIGURATION ---
    const PURCHASE_HISTORY_LIMIT = 5; 
    // -------------------

    // --- Element References ---
    const psmFileInput = document.getElementById('psmFile');
    const nukFileInput = document.getElementById('nukFile');
    const psmFileName = document.getElementById('psmFileName');
    const nukFileName = document.getElementById('nukFileName');
    const mergeBtn = document.getElementById('mergeBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const loader = document.getElementById('loader');

    let mergedDataForDownload = [];

    // --- Event Listeners ---
    psmFileInput.addEventListener('change', () => updateFileName(psmFileInput, psmFileName));
    nukFileInput.addEventListener('change', () => updateFileName(nukFileInput, nukFileName));
    mergeBtn.addEventListener('click', handleMerge);
    downloadBtn.addEventListener('click', handleDownload);

    function updateFileName(input, nameElement) {
        nameElement.textContent = input.files.length > 0 ? input.files[0].name : 'No file chosen';
    }

    async function handleMerge() {
        const psmFile = psmFileInput.files[0];
        const nukFile = nukFileInput.files[0];
        if (!psmFile || !nukFile) {
            alert('Please select both PSM and NUK .xlsx files.');
            return;
        }
        loader.style.display = 'block';
        downloadBtn.style.display = 'none';
        mergeBtn.disabled = true;
        try {
            let psmData = await readFile(psmFile);
            let nukData = await readFile(nukFile);
            psmData = cleanData(psmData);
            nukData = cleanData(nukData);
            const merged = mergeRawData(psmData, nukData);
            mergedDataForDownload = processMergedData(merged);
            alert('Files cleaned, merged, and processed successfully! Ready for download.');
            downloadBtn.style.display = 'flex';
        } catch (error) {
            console.error('Error merging files:', error);
            alert('An error occurred. Please check the console for details.');
        } finally {
            loader.style.display = 'none';
            mergeBtn.disabled = false;
        }
    }
    
    function readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const workbook = XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    // Use cellNF:false to get raw numbers for EANs
                    const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: true, defval:'' });
                    resolve(json);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * UPDATED: No longer modifies EAN, only cleans other fields.
     */
    function cleanData(data) {
        return data.map(row => {
            const newRow = {};
            for (const key in row) {
                if (Object.prototype.hasOwnProperty.call(row, key)) {
                    let value = row[key];
                    if (key !== 'EAN' && typeof value === 'string') {
                        // Remove currency symbols from non-EAN string values
                        value = value.replace(/£/g, '');
                    }
                    newRow[key] = value;
                }
            }
            return newRow;
        });
    }

    function mergeRawData(psmData, nukData) {
        const merged = {};
        psmData.forEach(row => {
            const ean = row.EAN;
            if (!ean) return;
            merged[ean] = {};
            for (const key in row) {
                merged[ean][`${key} PSM`] = row[key];
            }
        });
        nukData.forEach(row => {
            const ean = row.EAN;
            if (!ean) return;
            if (!merged[ean]) {
                merged[ean] = { 'EAN NUK': ean };
            }
            for (const key in row) {
                 merged[ean][`${key} NUK`] = row[key];
            }
        });
        return Object.values(merged);
    }
    
    function parseDate(dateStr) {
        if (typeof dateStr !== 'string' || !dateStr.trim()) {
            return null;
        }
        const parts = dateStr.match(/(\d+)/g);
        if (parts && parts.length === 3) {
            let [d, m, y] = parts.map(Number);
            if (y < 100) y += 2000;
            if (y > 1900 && m > 0 && m <= 12 && d > 0 && d <= 31) {
                return new Date(y, m - 1, d);
            }
        }
        return null;
    }

    function processMergedData(data) {
        const today = new Date();
        const thirtyDaysAgo = new Date(new Date().setDate(today.getDate() - 30));

        return data.map(item => {
            const psmStock = parseInt(item['Stock PSM'], 10) || 0;
            const nukStock = parseInt(item['Stock NUK'], 10) || 0;

            const getPriceAndAgingStats = (suffix) => {
                let highestPrice = 0, totalValue = 0, totalQty = 0;
                let agedStockQty = 0, agedStockValue = 0;
                for (let i = 1; i <= PURCHASE_HISTORY_LIMIT; i++) {
                    const qtyVal = item[`Qty ${i} ${suffix}`];
                    const priceVal = item[`Price ${i} ${suffix}`];
                    const dateVal = item[`Purchase Date ${i} ${suffix}`];
                    if (qtyVal && priceVal) {
                        const qtyNum = parseInt(qtyVal, 10);
                        const priceNum = parseFloat(priceVal);
                        if (!isNaN(qtyNum) && !isNaN(priceNum) && qtyNum > 0) {
                            totalValue += qtyNum * priceNum;
                            totalQty += qtyNum;
                            if (priceNum > highestPrice) highestPrice = priceNum;
                            const purchaseDate = parseDate(dateVal);
                            if (purchaseDate && purchaseDate < thirtyDaysAgo) {
                                agedStockQty += qtyNum;
                                agedStockValue += qtyNum * priceNum;
                            }
                        }
                    }
                }
                return { highest: highestPrice, average: totalQty > 0 ? totalValue / totalQty : 0, totalValue, totalQty, agedStockQty, agedStockValue };
            };

            const psmStats = getPriceAndAgingStats('PSM');
            const nukStats = getPriceAndAgingStats('NUK');
            const combinedTotalValue = psmStats.totalValue + nukStats.totalValue;
            const combinedTotalQty = psmStats.totalQty + nukStats.totalQty;
            const totalAgedStockQty = psmStats.agedStockQty + nukStats.agedStockQty;
            const totalAgedStockValue = psmStats.agedStockValue + nukStats.agedStockValue;
            
            const finalRow = {
                'EAN': item['EAN PSM'] || item['EAN NUK'],
                'Product Name': item['Product Name PSM'] || item['Product Name NUK'],
                'PSM Stock': psmStock,
                'Nuk Stock': nukStock,
                'Total Stock': psmStock + nukStock,
                'Highest Price PSM': psmStats.highest.toFixed(2),
                'Highest Price NUK': nukStats.highest.toFixed(2),
                'Max Price': Math.max(psmStats.highest, nukStats.highest).toFixed(2),
                'Weighted Average Price PSM': psmStats.average.toFixed(2),
                'Weighted Average Price NUK': nukStats.average.toFixed(2),
                'Weighted Average Price': (combinedTotalQty > 0 ? combinedTotalValue / combinedTotalQty : 0).toFixed(2),
            };

            for (let i = 1; i <= PURCHASE_HISTORY_LIMIT; i++) {
                finalRow[`Qty ${i} PSM`] = item[`Qty ${i} PSM`] || '';
                finalRow[`Supplier ${i} PSM`] = item[`Supplier ${i} PSM`] || '';
                finalRow[`Purchase Date ${i} PSM`] = item[`Purchase Date ${i} PSM`] || '';
                finalRow[`Invoice No.${i} PSM`] = item[`Invoice No.${i} PSM`] || '';
                finalRow[`Price ${i} PSM`] = item[`Price ${i} PSM`] || '';
                finalRow[`Qty ${i} NUK`] = item[`Qty ${i} NUK`] || '';
                finalRow[`Supplier ${i} NUK`] = item[`Supplier ${i} NUK`] || '';
                finalRow[`Purchase Date ${i} NUK`] = item[`Purchase Date ${i} NUK`] || '';
                finalRow[`Invoice No.${i} NUK`] = item[`Invoice No.${i} NUK`] || '';
                finalRow[`Price ${i} NUK`] = item[`Price ${i} NUK`] || '';
            }

            finalRow['PSM Stock > 30 Days'] = psmStats.agedStockQty;
            finalRow['NUK Stock > 30 Days'] = nukStats.agedStockQty;
            finalRow['Total Stock > 30 Days'] = totalAgedStockQty;
            finalRow['Value of Aged Stock'] = totalAgedStockValue.toFixed(2);
            finalRow['WAP of Aged Stock'] = (totalAgedStockQty > 0 ? totalAgedStockValue / totalAgedStockQty : 0).toFixed(2);

            return finalRow;
        });
    }

    function handleDownload() {
        if (mergedDataForDownload.length === 0) {
            alert('No data to download.');
            return;
        }
        const now = new Date();
        const pad = (num) => num.toString().padStart(2, '0');
        const day = pad(now.getDate()), month = pad(now.getMonth() + 1), year = now.getFullYear();
        const hours = pad(now.getHours()), minutes = pad(now.getMinutes());
        const timestamp = `${day}/${month}/${year}-${hours}:${minutes}`;
        const fileName = `Aging-PSM-NUK-Merge-${timestamp}.xlsx`;
        
        const baseHeaders = ['EAN', 'Product Name', 'PSM Stock', 'Nuk Stock', 'Total Stock', 'Highest Price PSM', 'Highest Price NUK', 'Max Price', 'Weighted Average Price PSM', 'Weighted Average Price NUK', 'Weighted Average Price'];
        const historyHeaders = [];
        for (let i = 1; i <= PURCHASE_HISTORY_LIMIT; i++) {
            historyHeaders.push(`Qty ${i} PSM`, `Supplier ${i} PSM`, `Purchase Date ${i} PSM`, `Invoice No.${i} PSM`, `Price ${i} PSM`);
            historyHeaders.push(`Qty ${i} NUK`, `Supplier ${i} NUK`, `Purchase Date ${i} NUK`, `Invoice No.${i} NUK`, `Price ${i} NUK`);
        }
        const agingHeaders = ['PSM Stock > 30 Days', 'NUK Stock > 30 Days', 'Total Stock > 30 Days', 'Value of Aged Stock', 'WAP of Aged Stock'];
        const finalHeaders = [...baseHeaders, ...historyHeaders, ...agingHeaders];
        const worksheet = XLSX.utils.json_to_sheet(mergedDataForDownload, { header: finalHeaders });

        // --- UPDATED CODE TO FIX EAN FORMATTING ---
        // Loop through all data rows to apply specific number formatting to the EAN column
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let R = range.s.r + 1; R <= range.e.r; ++R) { // Start from row 2 (index 1) to skip header
            const cell_address = { c: 0, r: R }; // Column A (index 0)
            const cell_ref = XLSX.utils.encode_cell(cell_address);
            if (worksheet[cell_ref]) {
                worksheet[cell_ref].t = 'n';    // Set cell type to 'n' for Number
                worksheet[cell_ref].z = '0';    // Set the number format to a plain integer with no decimals
            }
        }
        // --- END OF UPDATED CODE ---

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Merged Stock');
        XLSX.writeFile(workbook, fileName);
    }
});