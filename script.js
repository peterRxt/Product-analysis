// --- Global Variables ---
let uploadedData = []; // Stores all processed data from uploaded files
let currentMappedColumns = {}; // Stores the confirmed column mappings
let currentReportData = []; // Stores the data of the currently displayed report
let currentChart = null; // Stores the active Chart.js instance for destruction
let currentSection = 'upload-section'; // Track current section for back button

// --- DOM Element References ---
const excelFileInput = document.getElementById('excelFileInput');
const columnMappingArea = document.getElementById('columnMappingArea');
const descriptionColSelect = document.getElementById('descriptionCol');
const quantityColSelect = document.getElementById('quantityCol');
const salesColSelect = document.getElementById('salesCol');
const costColSelect = document.getElementById('costCol');
const confirmMappingBtn = document.getElementById('confirmMappingBtn');
const postMappingNavigation = document.getElementById('postMappingNavigation');
const navSectionButtons = document.querySelectorAll('.nav-section-btn');
const topNavItems = document.querySelectorAll('.top-nav .nav-item');
const contentSections = document.querySelectorAll('.content-section');

// Report elements
const reportDropdownBtn = document.getElementById('reportDropdownBtn');
const reportDropdownContent = document.getElementById('reportDropdownContent');
const reportInputArea = document.getElementById('reportInputArea');
const numItemsInput = document.getElementById('numItemsInput');
const numItemsLabel = document.getElementById('numItemsLabel');
const dateRangeNote = document.getElementById('dateRangeNote');
const runReportBtn = document.getElementById('runReportBtn');
const reportOutput = document.getElementById('reportOutput');
const resultsTable = document.getElementById('resultsTable');
const optionalReportLinks = document.querySelectorAll('.optional-report');

// Visuals elements
const visReportDropdownBtn = document.getElementById('visReportDropdownBtn');
const visReportDropdownContent = document.getElementById('visReportDropdownContent');
const chartTypeDropdownBtn = document.getElementById('chartTypeDropdownBtn');
const chartTypeDropdownContent = document.getElementById('chartTypeDropdownContent');
const generateVisualBtn = document.getElementById('generateVisualBtn');
const chartContainer = document.getElementById('chartContainer');
const myChartCanvas = document.getElementById('myChart');
const noVisualMsg = document.getElementById('noVisualMsg');

// Control Buttons
const exportCsvBtn = document.getElementById('exportCsvBtn');
const backBtn = document.getElementById('backBtn');
const cancelBtn = document.getElementById('cancelBtn');

// Message elements
let messageContainer;

// --- Helper Functions ---
function showMessage(msg, type, showSpinner = false) {
    if (messageContainer) {
        messageContainer.remove();
    }

    messageContainer = document.createElement('div');
    messageContainer.classList.add('message-container', `message-${type}`);
    if (showSpinner && type === 'loading') {
        messageContainer.innerHTML = `<div class="spinner"></div><span>${msg}</span>`;
    } else {
        messageContainer.textContent = msg;
    }

    const currentActiveSection = document.querySelector('.content-section.active');
    if (currentActiveSection) {
        if (currentActiveSection.id === 'upload-section') {
            currentActiveSection.insertBefore(messageContainer, columnMappingArea);
        } else {
            currentActiveSection.insertBefore(messageContainer, currentActiveSection.querySelector('h2').nextElementSibling);
        }
    } else {
        document.querySelector('.container').prepend(messageContainer);
    }

    void messageContainer.offsetWidth;
    messageContainer.classList.add('show');

    if (type !== 'loading') {
        setTimeout(() => {
            hideMessage();
        }, 5000);
    }
}

function hideMessage() {
    if (messageContainer) {
        messageContainer.classList.remove('show');
        messageContainer.addEventListener('transitionend', () => {
            if (messageContainer && !messageContainer.classList.contains('show')) {
                messageContainer.remove();
                messageContainer = null;
            }
        }, { once: true });
    }
}

function setDisabled(element, disable) {
    if (element) {
        element.disabled = disable;
        if (disable) {
            element.classList.add('disabled');
        } else {
            element.classList.remove('disabled');
        }
    }
}

function levenshteinDistance(a, b) {
    const an = a.length;
    const bn = b.length;

    if (an === 0) return bn;
    if (bn === 0) return an;

    const matrix = [];

    for (let i = 0; i <= an; i++) {
        matrix[i] = [i];
    }

    for (let j = 0; j <= bn; j++) {
        matrix[0][j] = j;
    }

    for (let i = 1; i <= an; i++) {
        for (let j = 1; j <= bn; j++) {
            const cost = a[i - 1] === b[j - 1] ? 0 : 1;
            matrix[i][j] = Math.min(
                matrix[i - 1][j] + 1,
                matrix[i][j - 1] + 1,
                matrix[i - 1][j - 1] + cost
            );
        }
    }

    return matrix[an][bn];
}

function findBestMatch(target, headers, synonyms) {
    const searchTerms = [target.toLowerCase(), ...synonyms.map(s => s.toLowerCase())];
    let bestMatch = null;
    let minDistance = Infinity;
    const fuzzinessThreshold = 2;

    for (const header of headers) {
        const lowerHeader = header.toLowerCase();
        for (const term of searchTerms) {
            const distance = levenshteinDistance(term, lowerHeader);
            if (distance < minDistance) {
                minDistance = distance;
                bestMatch = header;
            }
        }
    }
    return minDistance <= fuzzinessThreshold ? bestMatch : null;
}

function capitalizeFirstLetter(string) {
    if (!string) return '';
    return string.charAt(0).toUpperCase() + string.slice(1);
}

function formatCurrency(value) {
    if (typeof value !== 'number' || isNaN(value)) {
        return '';
    }
    return value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// --- Data Processing Functions ---
async function processExcelFiles(files) {
    showMessage('Processing files...', 'loading', true);
    uploadedData = [];

    const allHeaders = new Set();
    const parsedWorkbooks = [];

    for (const file of files) {
        try {
            const data = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve(e.target.result);
                reader.onerror = (err) => reject(err);
                reader.readAsBinaryString(file);
            });

            const workbook = XLSX.read(data, { type: 'binary' });
            for (const sheetName of workbook.SheetNames) {
                const sheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });

                if (sheetData.length === 0) continue;

                const headers = sheetData[0].map(h => String(h || '').trim());
                headers.forEach(h => {
                    if (h) allHeaders.add(h);
                });

                const rows = sheetData.slice(1);
                parsedWorkbooks.push({ headers, rows });
            }
        } catch (error) {
            showMessage(`Error reading file "${file.name}": ${error.message}`, 'error');
            setDisabled(confirmMappingBtn, true);
            return;
        }
    }

    if (allHeaders.size === 0) {
        showMessage('No valid data or headers found in uploaded files.', 'error');
        setDisabled(confirmMappingBtn, true);
        return;
    }

    populateColumnMapping(Array.from(allHeaders));
    showMessage('Files processed. Please map your columns.', 'success');
    columnMappingArea.style.display = 'block';
    setDisabled(confirmMappingBtn, false);
}

function populateColumnMapping(headers) {
    const requiredColumns = {
        description: ['description', 'item', 'product', 'product description', 'item name', 'name'],
        quantity: ['qty', 'quantity', 'quantity sold', 'number sold', 'units sold', 'sold quantity', 'sales quantity'],
        sales: ['total sales', 'net sales', 'revenue', 'revenue sales', 'sales total', 'total revenue', 'value'],
        cost: ['cost per unit', 'unit cost', 'cogs', 'cost of goods', 'cost price', 'price per unit']
    };

    const selectElements = {
        description: descriptionColSelect,
        quantity: quantityColSelect,
        sales: salesColSelect,
        cost: costColSelect
    };

    Object.values(selectElements).forEach(select => {
        select.innerHTML = '<option value="">-- Select Column --</option>';
    });

    headers.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        Object.values(selectElements).forEach(select => select.appendChild(option.cloneNode(true)));
    });

    for (const key in requiredColumns) {
        const bestMatch = findBestMatch(key, headers, requiredColumns[key]);
        if (bestMatch) {
            selectElements[key].value = bestMatch;
        }
    }
}

function aggregateData() {
    showMessage('Aggregating data...', 'loading', true);
    const aggregatedMap = new Map();

    const descCol = currentMappedColumns.description;
    const qtyCol = currentMappedColumns.quantity;
    const salesCol = currentMappedColumns.sales;
    const costCol = currentMappedColumns.cost;

    if (!Array.isArray(uploadedData) || uploadedData.length === 0) {
        showMessage('No data available for aggregation. Please upload and map files first.', 'error');
        return null;
    }

    uploadedData.forEach(row => {
        const description = String(row[descCol] || '').trim().toLowerCase();
        const quantity = parseFloat(row[qtyCol]) || 0;
        const sales = parseFloat(row[salesCol]) || 0;
        const cost = costCol ? (parseFloat(row[costCol]) || 0) : 0;

        if (aggregatedMap.has(description)) {
            const existing = aggregatedMap.get(description);
            existing.Quantity += quantity;
            existing.TotalSales += sales;
            existing.TotalCost += cost * quantity;
            aggregatedMap.set(description, existing);
        } else {
            aggregatedMap.set(description, {
                Description: capitalizeFirstLetter(description),
                Quantity: quantity,
                TotalSales: sales,
                TotalCost: cost * quantity
            });
        }
    });

    const finalAggregatedData = Array.from(aggregatedMap.values());
    showMessage('Data aggregation complete!', 'success');
    return finalAggregatedData;
}

// --- UI Management Functions ---
function switchSection(targetSectionId) {
    currentSection = targetSectionId;
    contentSections.forEach(section => {
        section.classList.remove('active');
    });
    document.getElementById(targetSectionId).classList.add('active');

    topNavItems.forEach(item => {
        item.classList.remove('active');
        if (item.dataset.target === targetSectionId) {
            item.classList.add('active');
        }
    });

    if (targetSectionId === 'upload-section') {
        setControlButtonsVisibility(false, false, false);
    } else if (targetSectionId === 'analysis-section' || targetSectionId === 'visuals-section') {
        setControlButtonsVisibility(false, true, true);
        hideReportOutput();
        hideChart();
        hideMessage();
    }
}

function setControlButtonsVisibility(showExport, showBack, showCancel) {
    exportCsvBtn.style.display = showExport ? 'inline-block' : 'none';
    backBtn.style.display = showBack ? 'inline-block' : 'none';
    cancelBtn.style.display = showCancel ? 'inline-block' : 'none';
}

function hideReportOutput() {
    reportOutput.style.display = 'none';
    resultsTable.innerHTML = '';
    currentReportData = [];
    setControlButtonsVisibility(false, true, true);
}

function hideChart() {
    if (currentChart) {
        currentChart.destroy();
        currentChart = null;
    }
    chartContainer.style.display = 'none';
    myChartCanvas.style.display = 'none';
    noVisualMsg.style.display = 'block';
    setControlButtonsVisibility(false, true, true);
}

// --- Report Generation Logic ---
function displayReport(data, headers) {
    resultsTable.innerHTML = '';

    const thead = resultsTable.createTHead();
    const headerRow = thead.insertRow();
    headers.forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });

    const tbody = resultsTable.createTBody();
    data.forEach(item => {
        const row = tbody.insertRow();
        headers.forEach(headerText => {
            const cell = row.insertCell();
            let value = item[headerText];
            if ((headerText.includes('Sales') || headerText.includes('Cost') || headerText.includes('Profit')) && typeof value === 'number') {
                value = formatCurrency(value);
            } else if (typeof value === 'number' && headerText.includes('%')) {
                value = value.toFixed(2) + '%';
            } else if (typeof value === 'number' && (headerText.includes('Quantity') || headerText.includes('Qty'))) {
                 value = value.toLocaleString();
            } else if (value === null || value === undefined) {
                value = '';
            }
            cell.textContent = value;
        });
    });

    currentReportData = data;
    reportOutput.style.display = 'block';
    setControlButtonsVisibility(true, true, true);
}

function runSelectedReport(reportType) {
    hideMessage();
    hideReportOutput();
    hideChart();

    if (!uploadedData || uploadedData.length === 0) {
        showMessage('Please upload and map your data first.', 'error');
        return;
    }
    if (!currentMappedColumns.description || !currentMappedColumns.quantity || !currentMappedColumns.sales) {
        showMessage('Essential columns (Description, Quantity, Sales) are not mapped. Please map them.', 'error');
        return;
    }

    showMessage(`Generating ${capitalizeFirstLetter(reportType.replace('-', ' '))} report...`, 'loading', true);

    let reportData = [];
    let reportHeaders = ['Description', 'Quantity', 'Total Sales'];
    let numItems = parseInt(numItemsInput.value);

    const aggregated = aggregateData();
    if (!aggregated) {
        hideMessage();
        return;
    }

    try {
        switch (reportType) {
            case 'fast-moving':
                if (isNaN(numItems) {
                    numItems = 10; // Default value if not specified
                }
                if (numItems <= 0 || numItems > aggregated.length) {
                    showMessage(`Please enter a valid number of items (1 to ${aggregated.length}) for Fast Moving report.`, 'error');
                    return;
                }
                reportData = aggregated.sort((a, b) => b.Quantity - a.Quantity).slice(0, numItems);
                break;

            case 'slow-moving':
                if (isNaN(numItems)) {
                    numItems = 10; // Default value if not specified
                }
                if (numItems <= 0 || numItems > aggregated.length) {
                    showMessage(`Please enter a valid number of items (1 to ${aggregated.length}) for Slow Moving report.`, 'error');
                    return;
                }
                reportData = aggregated.sort((a, b) => a.Quantity - b.Quantity).slice(0, numItems);
                break;

            case 'contribution':
                const totalSales = aggregated.reduce((sum, item) => sum + item.TotalSales, 0);
                reportData = aggregated.map(item => ({
                    ...item,
                    'Contribution %': totalSales > 0 ? (item.TotalSales / totalSales) * 100 : 0
                }));
                reportData.sort((a, b) => b['Contribution %'] - a['Contribution %']);
                reportHeaders.push('Contribution %');
                break;

            case 'growth-rate':
                showMessage('Growth Rate analysis requires specific date columns and will be implemented in a future update.', 'info');
                hideMessage();
                return;

            case 'profitability':
                if (!currentMappedColumns.cost) {
                    showMessage('Cost per unit column is not mapped. Profitability analysis not available.', 'error');
                    hideMessage();
                    return;
                }
                reportData = aggregated.map(item => {
                    const profit = item.TotalSales - item.TotalCost;
                    const profitMargin = item.TotalSales > 0 ? (profit / item.TotalSales) * 100 : 0;
                    return {
                        ...item,
                        'Total Cost': item.TotalCost,
                        'Profit': profit,
                        'Profit Margin %': profitMargin
                    };
                });
                reportData.sort((a, b) => b.Profit - a.Profit);
                reportHeaders.push('Total Cost', 'Profit', 'Profit Margin %');
                break;

            case 'price-vs-sales':
                if (!currentMappedColumns.cost) {
                    showMessage('Cost per unit column is not mapped. Price vs. Sales Trend analysis not available.', 'error');
                    hideMessage();
                    return;
                }
                reportData = aggregated.map(item => ({
                    ...item,
                    'Unit Price': item.Quantity > 0 ? item.TotalSales / item.Quantity : 0
                }));
                reportHeaders.push('Unit Price');
                reportData.sort((a, b) => b['Total Sales'] - a['Total Sales']);
                break;

            default:
                showMessage('Unknown report type selected.', 'error');
                hideMessage();
                return;
        }

        displayReport(reportData, reportHeaders);
        showMessage(`${capitalizeFirstLetter(reportType.replace('-', ' '))} report generated successfully!`, 'success');

    } catch (error) {
        showMessage(`Error generating report: ${error.message}`, 'error');
        console.error("Report generation error:", error);
    }
}

// --- Data Visualization Logic ---
function generateChart(reportType, chartType) {
    hideMessage();
    if (!currentReportData || currentReportData.length === 0 || !currentMappedColumns.description) {
        showMessage('No report data available to visualize. Please run a report first.', 'error');
        hideChart();
        return;
    }

    if (currentChart) {
        currentChart.destroy();
    }

    let labels = [];
    let dataValues = [];
    let datasetLabel = '';
    let backgroundColor = [];
    let borderColor = [];

    switch (reportType) {
        case 'fast-moving':
        case 'slow-moving':
            labels = currentReportData.map(item => item.Description);
            dataValues = currentReportData.map(item => item.Quantity);
            datasetLabel = 'Quantity Sold';
            backgroundColor = 'rgba(75, 192, 192, 0.6)';
            borderColor = 'rgba(75, 192, 192, 1)';
            break;
        case 'contribution':
            labels = currentReportData.map(item => item.Description);
            dataValues = currentReportData.map(item => item['Contribution %']);
            datasetLabel = 'Sales Contribution (%)';
            backgroundColor = 'rgba(153, 102, 255, 0.6)';
            borderColor = 'rgba(153, 102, 255, 1)';
            break;
        case 'profitability':
            labels = currentReportData.map(item => item.Description);
            dataValues = currentReportData.map(item => item.Profit);
            datasetLabel = 'Profit';
            backgroundColor = 'rgba(255, 159, 64, 0.6)';
            borderColor = 'rgba(255, 159, 64, 1)';
            break;
        case 'price-vs-sales':
            labels = currentReportData.map(item => item.Description);
            dataValues = currentReportData.map(item => item['Total Sales']);
            datasetLabel = 'Total Sales';
            backgroundColor = 'rgba(54, 162, 235, 0.6)';
            borderColor = 'rgba(54, 162, 235, 1)';
            break;
        default:
            showMessage('Visualization not supported for this report type or data.', 'error');
            hideChart();
            return;
    }

    if (['pie', 'doughnut', 'polarArea'].includes(chartType)) {
        backgroundColor = labels.map((_, i) => `hsl(${i * 360 / labels.length}, 70%, 60%)`);
        borderColor = 'white';
    }

    const ctx = myChartCanvas.getContext('2d');
    currentChart = new Chart(ctx, {
        type: chartType,
        data: {
            labels: labels,
            datasets: [{
                label: datasetLabel,
                data: dataValues,
                backgroundColor: Array.isArray(backgroundColor) ? backgroundColor : [backgroundColor],
                borderColor: Array.isArray(borderColor) ? borderColor : [borderColor],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    display: !['pie', 'doughnut', 'polarArea'].includes(chartType)
                },
                x: {
                     display: !['pie', 'doughnut', 'polarArea'].includes(chartType)
                }
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed.y !== undefined) {
                                if (reportType === 'contribution') {
                                    label += context.parsed.y.toFixed(2) + '%';
                                } else if (reportType === 'profitability') {
                                    label += formatCurrency(context.parsed.y);
                                } else {
                                    label += context.parsed.y.toLocaleString();
                                }
                            } else if (context.parsed !== undefined) {
                                if (reportType === 'contribution') {
                                    label += context.parsed.toFixed(2) + '%';
                                } else {
                                    label += context.parsed.toLocaleString();
                                }
                            }
                            return label;
                        }
                    }
                }
            }
        }
    });

    chartContainer.style.display = 'block';
    myChartCanvas.style.display = 'block';
    noVisualMsg.style.display = 'none';
    setControlButtonsVisibility(true, true, true);
    showMessage('Chart generated successfully!', 'success');
}

// --- Event Listeners ---
excelFileInput.addEventListener('change', async (event) => {
    const files = event.target.files;
    if (files.length > 0) {
        await processExcelFiles(files);
    } else {
        hideMessage();
        columnMappingArea.style.display = 'none';
        postMappingNavigation.style.display = 'none';
        uploadedData = [];
        showMessage('No files selected.', 'info');
    }
});

confirmMappingBtn.addEventListener('click', () => {
    hideMessage();
    const desc = descriptionColSelect.value;
    const qty = quantityColSelect.value;
    const sales = salesColSelect.value;
    const cost = costColSelect.value;

    const selectedCols = [desc, qty, sales].filter(Boolean);
    const uniqueCols = new Set(selectedCols);

    if (selectedCols.length !== 3 || uniqueCols.size !== 3) {
        showMessage('Please select unique columns for Description, Quantity, and Total Sales.', 'error');
        return;
    }

    currentMappedColumns = {
        description: desc,
        quantity: qty,
        sales: sales,
        cost: cost || null
    };

    showMessage('Columns mapped successfully!', 'success');
    columnMappingArea.style.display = 'none';
    postMappingNavigation.style.display = 'flex';

    if (!currentMappedColumns.cost) {
        optionalReportLinks.forEach(link => link.classList.add('disabled'));
        visReportDropdownContent.querySelectorAll('.optional-report').forEach(link => link.classList.add('disabled'));
    } else {
        optionalReportLinks.forEach(link => link.classList.remove('disabled'));
        visReportDropdownContent.querySelectorAll('.optional-report').forEach(link => link.classList.remove('disabled'));
    }
});

navSectionButtons.forEach(button => {
    button.addEventListener('click', () => {
        const targetSection = button.dataset.target;
        switchSection(targetSection);
    });
});

reportDropdownBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    reportDropdownContent.classList.toggle('active');
    chartTypeDropdownContent.classList.remove('active');
    visReportDropdownContent.classList.remove('active');
});

reportDropdownContent.addEventListener('click', (event) => {
    event.preventDefault();
    const targetLink = event.target.closest('a');
    if (targetLink && !targetLink.classList.contains('disabled')) {
        const reportType = targetLink.dataset.report;
        reportDropdownBtn.textContent = targetLink.textContent + ' ▼';
        reportDropdownContent.classList.remove('active');

        reportInputArea.style.display = 'flex';
        numItemsInput.style.display = 'none';
        numItemsLabel.style.display = 'none';
        dateRangeNote.style.display = 'none';
        generateVisualBtn.style.display = 'none';

        if (reportType === 'fast-moving' || reportType === 'slow-moving') {
            numItemsInput.style.display = 'inline-block';
            numItemsLabel.style.display = 'inline-block';
            numItemsLabel.textContent = 'Number of items (Max: ' + (uploadedData.length > 0 ? aggregateData().length : 0) + '):';
            numItemsInput.value = '';
        } else if (reportType === 'growth-rate') {
            dateRangeNote.style.display = 'block';
        }

        reportInputArea.dataset.selectedReport = reportType;
    }
});

runReportBtn.addEventListener('click', () => {
    const selectedReport = reportInputArea.dataset.selectedReport;
    if (selectedReport) {
        runSelectedReport(selectedReport);
    } else {
        showMessage('Please select a report type first.', 'info');
    }
});

visReportDropdownBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    visReportDropdownContent.classList.toggle('active');
    reportDropdownContent.classList.remove('active');
    chartTypeDropdownContent.classList.remove('active');
});

visReportDropdownContent.addEventListener('click', (event) => {
    event.preventDefault();
    const targetLink = event.target.closest('a');
    if (targetLink && !targetLink.classList.contains('disabled')) {
        const reportType = targetLink.dataset.report;
        visReportDropdownBtn.textContent = targetLink.textContent + ' ▼';
        visReportDropdownContent.classList.remove('active');
        visReportDropdownBtn.dataset.selectedReport = reportType;
        if (chartTypeDropdownBtn.dataset.selectedChartType) {
            generateVisualBtn.style.display = 'inline-block';
        }
    }
});

chartTypeDropdownBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    chartTypeDropdownContent.classList.toggle('active');
    reportDropdownContent.classList.remove('active');
    visReportDropdownContent.classList.remove('active');
});

chartTypeDropdownContent.addEventListener('click', (event) => {
    event.preventDefault();
    const targetLink = event.target.closest('a');
    if (targetLink) {
        const chartType = targetLink.dataset.chartType;
        chartTypeDropdownBtn.textContent = targetLink.textContent + ' ▼';
        chartTypeDropdownContent.classList.remove('active');
        chartTypeDropdownBtn.dataset.selectedChartType = chartType;
        if (visReportDropdownBtn.dataset.selectedReport) {
            generateVisualBtn.style.display = 'inline-block';
        }
    }
});

generateVisualBtn.addEventListener('click', () => {
    const selectedReport = visReportDropdownBtn.dataset.selectedReport;
    const selectedChartType = chartTypeDropdownBtn.dataset.selectedChartType;

    if (selectedReport && selectedChartType) {
        generateChart(selectedReport, selectedChartType);
    } else {
        showMessage('Please select both a report and a chart type.', 'info');
    }
});

document.addEventListener('click', (event) => {
    if (!reportDropdownBtn.contains(event.target) && !reportDropdownContent.contains(event.target)) {
        reportDropdownContent.classList.remove('active');
    }
    if (!visReportDropdownBtn.contains(event.target) && !visReportDropdownContent.contains(event.target)) {
        visReportDropdownContent.classList.remove('active');
    }
    if (!chartTypeDropdownBtn.contains(event.target) && !chartTypeDropdownContent.contains(event.target)) {
        chartTypeDropdownContent.classList.remove('active');
    }
    hideMessage();
});

// --- Control Button Actions ---
exportCsvBtn.addEventListener('click', () => {
    if (currentReportData && currentReportData.length > 0) {
        showMessage('Exporting CSV...', 'loading', true);
        const headers = Array.from(resultsTable.querySelectorAll('thead th')).map(th => th.textContent);
        const csvRows = [];
        csvRows.push(headers.join(','));

        currentReportData.forEach(row => {
            const values = headers.map(header => {
                let value = row[header];
                if (typeof value === 'string') {
                    value = `"${value.replace(/"/g, '""')}"`;
                } else if (typeof value === 'number') {
                    value = value.toFixed(2);
                }
                return value;
            });
            csvRows.push(values.join(','));
        });

        const csvString
