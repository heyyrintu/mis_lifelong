class ExcelMISGenerator {
    constructor() {
        this.rawData = [];
        this.loadedData = null; // Store loaded and processed data
        this.isDataLoaded = false; // Track if data is loaded
        this.b2cData = [];
        this.ecomData = [];
        this.offlineData = [];
        this.quickcomData = [];
        this.eboData = [];
        this.othersData = [];
        
        // Performance optimizations
        this.cache = new Map();
        this.processingQueue = [];
        this.isProcessing = false;
        this.compiledRegex = {
            b2c: /decathlon|flflipkart\(b2c\)|snapmint|shopify|tatacliq|amazon b2c|pepperfry/i,
            ecom: /amazon|flipkart/i,
            offline: /offline sales-b2b|offline – gt|offline - mt/i,
            quickcom: /blinkit|swiggy|bigbasket|zepto/i,
            ebo: /store 2-lucknow|store3-zirakpur/i,
            others: /sales to vendor|internal company|others/i
        };
        this.performanceMonitor = new PerformanceMonitor();
        
        this.initializeEventListeners();
    }
    
    // Cleanup function to prevent memory leaks
    cleanup() {
        // Clear caches
        this.cache.clear();
        
        // Clear cached quantities from data objects
        if (this.rawData && this.rawData.length > 0) {
            for (let i = 1; i < this.rawData.length; i++) {
                const row = this.rawData[i];
                if (row && row._cachedQuantity !== undefined) {
                    delete row._cachedQuantity;
                }
            }
        }
        
        // Clear performance metrics
        this.performanceMonitor.metrics.clear();
        
        console.log('🧹 Cleanup completed - memory optimized');
    }

    initializeEventListeners() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const loadDataBtn = document.getElementById('loadDataBtn');
        const generateReports = document.getElementById('generateReports');
        const downloadAll = document.getElementById('downloadAll');

        if (!uploadArea || !fileInput) {
            console.error("Upload elements not found!");
            return;
        }

        // File upload events (check upload permission)
        uploadArea.addEventListener('click', () => {
            if (window.auth && window.auth.canUpload()) {
                fileInput.click();
            } else {
                this.showError('You do not have permission to upload files. Contact administrator.');
            }
        });
        
        uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        uploadArea.addEventListener('drop', (e) => this.handleDrop(e));
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Load Data button
        if (loadDataBtn) {
            loadDataBtn.addEventListener('click', () => {
                if (window.auth && window.auth.canWrite()) {
                    this.loadData();
                } else {
                    this.showError('You do not have permission to load data. Contact administrator.');
                }
            });
        }

        // Report generation (check write permission)
        if (generateReports) {
            generateReports.addEventListener('click', () => {
                if (window.auth && window.auth.canWrite()) {
                    if (this.isDataLoaded) {
                        this.generateAllReports();
                    } else {
                        this.showError('Please load data first before generating reports.');
                    }
                } else {
                    this.showError('You do not have permission to generate reports. Contact administrator.');
                }
            });
        }
        
        // Download functionality (check download permission)
        if (downloadAll) {
            downloadAll.addEventListener('click', () => {
                if (window.auth && window.auth.canDownload()) {
                    this.downloadAllReports();
                } else {
                    this.showError('You do not have permission to download files. Contact administrator.');
                }
            });
        }

        // Add event listener for sort order changes
        const sortOrderSelect = document.getElementById('sortOrder');
        if (sortOrderSelect) {
            sortOrderSelect.addEventListener('change', () => {
                // Re-sort and re-display reports if data is already loaded
                if (this.ecomData.length > 0 || this.quickcomData.length > 0 || this.offlineData.length > 0) {
                    this.displayAllReports();
                }
            });
        }
    }

    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        const uploadArea = document.getElementById('uploadArea');
        if (uploadArea) {
            uploadArea.classList.add('dragover');
        }
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        const uploadArea = document.getElementById('uploadArea');
        if (uploadArea) {
            uploadArea.classList.remove('dragover');
        }
    }

    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        
        const uploadArea = document.getElementById('uploadArea');
        if (uploadArea) {
            uploadArea.classList.remove('dragover');
        }
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            this.processFile(file);
        }
    }

    processFile(file) {
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel',
            'text/csv',
            'application/csv'
        ];

        if (!allowedTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls|csv)$/i)) {
            this.showError('Please upload a valid Excel (.xlsx, .xls) or CSV file.');
            return;
        }

        // Show file info
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');
        const fileInfo = document.getElementById('fileInfo');
        
        if (fileName && fileSize && fileInfo) {
            fileName.textContent = file.name;
            fileSize.textContent = this.formatFileSize(file.size);
            fileInfo.style.display = 'block';
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                let workbook;
                if (file.name.toLowerCase().endsWith('.csv')) {
                    // Handle CSV files
                    const csvData = e.target.result;
                    workbook = XLSX.read(csvData, { type: 'string' });
                } else {
                    // Handle Excel files
                    const data = new Uint8Array(e.target.result);
                    workbook = XLSX.read(data, { type: 'array' });
                }
                
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                this.rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (this.rawData.length > 1) {
                    const recordCount = document.getElementById('recordCount');
                    const fileInfo = document.getElementById('fileInfo');
                    
                    if (recordCount && fileInfo) {
                        recordCount.textContent = `${this.rawData.length - 1} records`;
                        fileInfo.style.display = 'block';
                        
                        // Show load status as ready
                        const loadStatus = document.getElementById('loadStatus');
                        if (loadStatus) {
                            loadStatus.textContent = 'Ready to load data';
                            loadStatus.className = 'load-status';
                        }
                        
                        // Reset data loaded state
                        this.isDataLoaded = false;
                        this.loadedData = null;
                        
                        // Hide generation section until data is loaded
                        const generationSection = document.getElementById('generationSection');
                        if (generationSection) {
                            generationSection.style.display = 'none';
                        }
                    }
                    
                    this.hideError();
                } else {
                    this.showError('The uploaded file appears to be empty or has no data rows.');
                }
            } catch (error) {
                this.showError('Error reading file: ' + error.message);
                console.error(error);
            }
        };
        
        reader.onerror = (error) => {
            this.showError('Error reading file: ' + error.message);
            console.error(error);
        };
        
        if (file.name.toLowerCase().endsWith('.csv')) {
            reader.readAsText(file);
        } else {
            reader.readAsArrayBuffer(file);
        }
    }

    // Format CBM to have exactly two decimal places (ensuring non-negative values)
    formatCBM(value) {
        try {
            if (value === null || value === undefined || value === '') {
                return '0.00';
            }
            
            const cbm = parseFloat(value);
            
            // Handle NaN values
            if (isNaN(cbm)) {
                console.warn('Invalid CBM value:', value);
                return '0.00';
            }
            
            // Ensure we don't format negative CBM values (additional safety check)
            const safeCBM = Math.max(0, cbm);
            return safeCBM.toFixed(2);
        } catch (error) {
            console.error('Error formatting CBM:', error, value);
            return '0.00';
        }
    }

    // Optimized quantity calculation with caching
    getQuantity(row) {
        if (!row || typeof row !== 'object') return 0;
        
        // Check cache first
        if (row._cachedQuantity !== undefined) {
            return row._cachedQuantity;
        }

        try {
            // Try different quantity fields in order of preference
            let qty = row['SALES Invoice QTY'] || row['DELIVERY Note QTY'] || 0;
            
            if (qty === null || qty === undefined) {
                row._cachedQuantity = 0;
                return 0;
            }
            
            // Fast numeric conversion
            const numQty = +qty; // Faster than parseFloat for most cases
            
            // Cache the result
            row._cachedQuantity = isNaN(numQty) ? 0 : Math.max(0, numQty);
            return row._cachedQuantity;
            
        } catch (error) {
            row._cachedQuantity = 0;
            return 0;
        }
    }

    // Load and process data from uploaded Excel file
    loadData() {
        try {
            if (this.rawData.length === 0) {
                this.showError('No file uploaded. Please upload an Excel file first.');
                return;
            }

            // Update load status
            const loadStatus = document.getElementById('loadStatus');
            const loadDataBtn = document.getElementById('loadDataBtn');
            
            if (loadStatus) {
                loadStatus.textContent = 'Loading data...';
                loadStatus.className = 'load-status loading';
            }
            
            if (loadDataBtn) {
                loadDataBtn.disabled = true;
                loadDataBtn.textContent = 'Loading...';
            }

            this.performanceMonitor.start('Data Loading');

            // Process the raw data
            console.log('📊 Starting data loading process...');
            console.log(`Processing ${this.rawData.length} rows from Excel file`);

            // Convert raw data to objects
            const headers = this.rawData[0];
            if (!headers || headers.length === 0) {
                throw new Error('No headers found in the Excel file');
            }

            console.log('📋 Excel Headers:', headers);

            const dataObjects = [];
            for (let i = 1; i < this.rawData.length; i++) {
                const row = this.rawData[i];
                if (!row || row.length === 0) continue;

                const obj = {};
                for (let j = 0; j < headers.length; j++) {
                    obj[headers[j]] = row[j];
                }
                dataObjects.push(obj);
            }

            console.log(`✅ Converted ${dataObjects.length} data rows to objects`);

            // Filter out negative values
            console.log('🔍 Filtering negative values...');
            this.loadedData = this.filterNegativeValues(dataObjects);
            console.log(`✅ Filtered data: ${this.loadedData.length} valid rows`);

            // Initialize dashboard with loaded data
            if (window.dashboard) {
                console.log('🎯 Initializing dashboard with loaded data...');
                window.dashboard.setData(this.loadedData);
                
                // Show dashboard section
                const dashboardSection = document.getElementById('dashboardSection');
                if (dashboardSection) {
                    dashboardSection.style.display = 'block';
                }
            }

            // Mark data as loaded
            this.isDataLoaded = true;
            this.performanceMonitor.end('Data Loading');

            // Update UI
            if (loadStatus) {
                loadStatus.textContent = `✅ Data loaded successfully! ${this.loadedData.length} records ready for processing`;
                loadStatus.className = 'load-status success';
            }
            
            if (loadDataBtn) {
                loadDataBtn.disabled = false;
                loadDataBtn.textContent = 'Reload Data';
            }

            // Show generation section
            const generationSection = document.getElementById('generationSection');
            if (generationSection) {
                generationSection.style.display = 'block';
            }

            console.log('🎉 Data loading completed successfully!');
            this.hideError();

        } catch (error) {
            console.error('❌ Error loading data:', error);
            
            const loadStatus = document.getElementById('loadStatus');
            const loadDataBtn = document.getElementById('loadDataBtn');
            
            if (loadStatus) {
                loadStatus.textContent = `❌ Error loading data: ${error.message}`;
                loadStatus.className = 'load-status error';
            }
            
            if (loadDataBtn) {
                loadDataBtn.disabled = false;
                loadDataBtn.textContent = 'Load Data';
            }
            
            this.showError(`Failed to load data: ${error.message}`);
        }
    }

    generateAllReports() {
        try {
            if (!this.isDataLoaded || !this.loadedData) {
                this.showError('Please load data first before generating reports.');
                return;
            }

            this.performanceMonitor.start('Total Report Generation');

            console.log('📊 Starting report generation with loaded data...');
            console.log(`Processing ${this.loadedData.length} loaded records`);

            // Use the already loaded and filtered data
            const reportData = this.loadedData;
            
            // Generate all six reports with loaded data
            this.performanceMonitor.start('Report Generation');
            this.generateB2CReport(reportData);
            this.generateEcomReport(reportData);
            this.generateOfflineReport(reportData);
            this.generateQuickcomReport(reportData);
            this.generateEBOReport(reportData);
            this.generateOthersReport(reportData);
            this.performanceMonitor.end('Report Generation');

            // Display results
            this.displayAllReports();
            const resultsSection = document.getElementById('resultsSection');
            if (resultsSection) {
                resultsSection.style.display = 'block';
            }
            this.hideError();
            
            // Show success notification to user
            this.showSuccessNotification(reportData.length);
            
            // Dashboard is already updated with the data during loadData()
            
            // Store data in localStorage for LR Pending page
            try {
                localStorage.setItem('excelData', JSON.stringify(reportData));
                console.log('Stored report data in localStorage:', reportData.length, 'records');
                
                // Store in instance for direct access
                this.filteredData = reportData;
            } catch (error) {
                console.warn('Could not store data in localStorage:', error);
            }

            // Log total quantities for verification
            this.logTotalQuantities();
            
            // Log detailed quantity breakdown for debugging
            this.logQuantityBreakdown(reportData);
            
            // Note: Data integrity was already validated during loadData()
            
            this.performanceMonitor.end('Total Report Generation');

            // Show LR Missing section and update data
            showLRMissingSection();
            updateLRMissingData(reportData);
            
            // Debug: Log LR counts for comparison
            const dashboardStats = window.dashboard ? window.dashboard.calculateStats(reportData) : null;
            const lrMissingCount = reportData.filter(row => {
                const lrNo = row['SHIPMENT Awb NUMBER'] || '';
                return !lrNo || lrNo.toString().trim() === '';
            }).length;

            console.log('Dashboard LR Pending:', dashboardStats ? dashboardStats.lrPending : 'N/A');
            console.log('LR Missing Section Count:', lrMissingCount);
            console.log('Data consistency check:', dashboardStats ? (dashboardStats.lrPending === lrMissingCount) : 'N/A');

        } catch (error) {
            this.showError('Error generating reports: ' + error.message);
            console.error(error);
        }
    }

    // Enhanced method to filter out rows with negative CBM and quantity values
    filterNegativeValues(data) {
        if (!Array.isArray(data) || data.length === 0) {
            console.warn('⚠️ No data provided to filterNegativeValues');
            return [];
        }

        let filteredCount = 0;
        let negativeCBMCount = 0;
        let negativeQtyCount = 0;
        
        // Optimized filtering using for loop
        const filteredData = [];
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            
            // Fast numeric conversion using unary plus operator
            const siTotalCBM = +(row['SI Total CBM'] || 0);
            const dnTotalCBM = +(row['DN Total CBM'] || 0);
            const perUnitCBM = +(row['Per Unit CBM'] || 0);
            const salesQty = +(row['SALES Invoice QTY'] || 0);
            const deliveryQty = +(row['DELIVERY Note QTY'] || 0);
            
            // Check for negative values
            if (siTotalCBM < 0 || dnTotalCBM < 0 || perUnitCBM < 0 || 
                salesQty < 0 || deliveryQty < 0) {
                    filteredCount++;
                if (siTotalCBM < 0 || dnTotalCBM < 0 || perUnitCBM < 0) negativeCBMCount++;
                if (salesQty < 0 || deliveryQty < 0) negativeQtyCount++;
                continue; // Skip this row
            }
            
            filteredData.push(row);
        }
        
        // Log comprehensive filtering summary
        console.log('📊 FILTERING SUMMARY:');
        console.log(`   Original rows: ${data.length}`);
        console.log(`   Filtered out: ${filteredCount} rows`);
        console.log(`   - Negative CBM: ${negativeCBMCount} rows`);
        console.log(`   - Negative Quantity: ${negativeQtyCount} rows`);
        console.log(`   Valid rows remaining: ${filteredData.length}`);
        console.log(`   Filtering rate: ${((filteredCount / data.length) * 100).toFixed(2)}%`);
        
        return filteredData;
    }

    logTotalQuantities() {
        // Calculate total quantities for verification
        let ecomTotal = 0;
        let quickcomTotal = 0;
        let offlineTotal = 0;
        
        // Calculate LR Pending for each category
        let ecomLRPending = 0;
        let quickcomLRPending = 0;
        let offlineLRPending = 0;
        
        this.ecomData.forEach(item => {
            const qty = parseFloat(item['Invoice Qty']) || 0;
            ecomTotal += Math.max(0, qty); // Ensure non-negative
            
            // Check for LR Pending
            const lrNo = item['LR No.'] || '';
            if (!lrNo || lrNo.toString().trim() === '') {
                ecomLRPending++;
            }
        });
        
        this.quickcomData.forEach(item => {
            const qty = parseFloat(item['Invoice Qty']) || 0;
            quickcomTotal += Math.max(0, qty); // Ensure non-negative
            
            // Check for LR Pending
            const lrNo = item['LR No.'] || '';
            if (!lrNo || lrNo.toString().trim() === '') {
                quickcomLRPending++;
            }
        });
        
        this.offlineData.forEach(item => {
            const qty = parseFloat(item['Invoice Qty']) || 0;
            offlineTotal += Math.max(0, qty); // Ensure non-negative
            
            // Check for LR Pending
            const lrNo = item['LR No.'] || '';
            if (!lrNo || lrNo.toString().trim() === '') {
                offlineLRPending++;
            }
        });
        
        const grandTotal = ecomTotal + quickcomTotal + offlineTotal;
        const totalLRPending = ecomLRPending + quickcomLRPending + offlineLRPending;
        
        console.log('=== PROCESSED DATA SUMMARY (After filtering negative values) ===');
        console.log('E-commerce - Records:', this.ecomData.length, '| Total Quantity:', ecomTotal.toFixed(2), '| LR Pending:', ecomLRPending);
        console.log('Quick-commerce - Records:', this.quickcomData.length, '| Total Quantity:', quickcomTotal.toFixed(2), '| LR Pending:', quickcomLRPending);
        console.log('Offline - Records:', this.offlineData.length, '| Total Quantity:', offlineTotal.toFixed(2), '| LR Pending:', offlineLRPending);
        console.log('GRAND TOTAL - Records:', (this.ecomData.length + this.quickcomData.length + this.offlineData.length), '| Total Quantity:', grandTotal.toFixed(2), '| Total LR Pending:', totalLRPending);
        console.log('================================================================');
    }
    
    logQuantityBreakdown(data) {
        let salesInvoiceTotal = 0;
        let deliveryNoteTotal = 0;
        let fallbackUsed = 0;
        let zeroQuantities = 0;
        let negativeQuantities = 0;
        let invalidQuantities = 0;
        let maxSalesQty = 0;
        let maxDeliveryQty = 0;
        
        data.forEach(row => {
            const salesQty = parseFloat(row['SALES Invoice QTY'] || 0);
            const deliveryQty = parseFloat(row['DELIVERY Note QTY'] || 0);
            
            // Track maximum values
            if (salesQty > maxSalesQty) maxSalesQty = salesQty;
            if (deliveryQty > maxDeliveryQty) maxDeliveryQty = deliveryQty;
            
            // Count negative values
            if (salesQty < 0) negativeQuantities++;
            if (deliveryQty < 0) negativeQuantities++;
            
            // Count invalid values
            if (isNaN(salesQty) && row['SALES Invoice QTY']) invalidQuantities++;
            if (isNaN(deliveryQty) && row['DELIVERY Note QTY']) invalidQuantities++;
            
            if (salesQty > 0) {
                salesInvoiceTotal += salesQty;
            }
            if (deliveryQty > 0) {
                deliveryNoteTotal += deliveryQty;
            }
            
            // Check if fallback was used
            if ((!salesQty || salesQty === 0) && deliveryQty > 0) {
                fallbackUsed++;
            }
            
            // Check for zero quantities
            if (salesQty === 0 && deliveryQty === 0) {
                zeroQuantities++;
            }
        });
        
        console.log('📊 DETAILED QUANTITY BREAKDOWN ANALYSIS:');
        console.log(`   Total rows processed: ${data.length.toLocaleString()}`);
        console.log(`   Total SALES Invoice QTY: ${salesInvoiceTotal.toLocaleString()}`);
        console.log(`   Total DELIVERY Note QTY: ${deliveryNoteTotal.toLocaleString()}`);
        console.log(`   Rows using DELIVERY Note fallback: ${fallbackUsed.toLocaleString()}`);
        console.log(`   Rows with zero quantities: ${zeroQuantities.toLocaleString()}`);
        console.log(`   Rows with negative quantities: ${negativeQuantities.toLocaleString()}`);
        console.log(`   Rows with invalid quantities: ${invalidQuantities.toLocaleString()}`);
        console.log(`   Max SALES Invoice QTY: ${maxSalesQty.toLocaleString()}`);
        console.log(`   Max DELIVERY Note QTY: ${maxDeliveryQty.toLocaleString()}`);
        console.log(`   Expected total (with fallback): ${(salesInvoiceTotal + deliveryNoteTotal).toLocaleString()}`);
        console.log('===============================================================');
    }
    
    validateDataIntegrity(originalData, filteredData) {
        const originalCount = originalData.length;
        const filteredCount = filteredData.length;
        const removedCount = originalCount - filteredCount;
        
        // Calculate totals from original data
        let originalTotalQty = 0;
        let originalTotalCBM = 0;
        
        originalData.forEach(row => {
            const qty = this.getQuantity(row);
            const cbm = parseFloat(row['SI Total CBM'] || row['DN Total CBM'] || 0);
            originalTotalQty += qty;
            originalTotalCBM += Math.max(0, cbm);
        });
        
        // Calculate totals from filtered data
        let filteredTotalQty = 0;
        let filteredTotalCBM = 0;
        
        filteredData.forEach(row => {
            const qty = this.getQuantity(row);
            const cbm = parseFloat(row['SI Total CBM'] || row['DN Total CBM'] || 0);
            filteredTotalQty += qty;
            filteredTotalCBM += Math.max(0, cbm);
        });
        
        console.log('🔍 DATA INTEGRITY VALIDATION:');
        console.log(`   Original rows: ${originalCount.toLocaleString()}`);
        console.log(`   Filtered rows: ${filteredCount.toLocaleString()}`);
        console.log(`   Removed rows: ${removedCount.toLocaleString()} (${((removedCount/originalCount)*100).toFixed(2)}%)`);
        console.log(`   Original total quantity: ${originalTotalQty.toLocaleString()}`);
        console.log(`   Filtered total quantity: ${filteredTotalQty.toLocaleString()}`);
        console.log(`   Quantity difference: ${(originalTotalQty - filteredTotalQty).toLocaleString()}`);
        console.log(`   Original total CBM: ${originalTotalCBM.toFixed(2)}`);
        console.log(`   Filtered total CBM: ${filteredTotalCBM.toFixed(2)}`);
        console.log(`   CBM difference: ${(originalTotalCBM - filteredTotalCBM).toFixed(2)}`);
        
        // Check for potential issues
        if (removedCount > originalCount * 0.1) {
            console.warn('⚠️  WARNING: More than 10% of data was filtered out!');
        }
        if ((originalTotalQty - filteredTotalQty) > originalTotalQty * 0.1) {
            console.warn('⚠️  WARNING: More than 10% of quantity was lost in filtering!');
        }
        
        console.log('===============================================================');
    }

    generateEcomReport(data) {
        // Optimized filtering using pre-compiled regex
        const ecomData = data.filter(row => {
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            return this.compiledRegex.ecom.test(customerGroup);
        });

        // E-com Excel headers
        this.ecomData = ecomData.map(row => {
            const quantity = this.getQuantity(row);
            
            return {
                'Customer Group': row['Customer Group'] || '',
                'Vehicle Series': row['SHIPMENT Vehicle NO'] || '',
                'Dispatch Date': row['SHIPMENT Pickup DATE'] || row['DELIVERY Note DATE'] || '',
                'Customer Name': row['Customer'] || '',
                'Transporter Name': row['Transporter'] || '',
                'Vehicle No': row['SHIPMENT Vehicle NO'] || '',
                'LR No.': row['SHIPMENT Awb NUMBER'] || '',
                'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || '',
                'Invoice Date': row['SALES Invoice DATE'] || '',
                'Invoice SKU': row['SO Item'] || row['Description of Content'] || '',
                'Invoice Qty': quantity,
                'Total CBM': this.formatCBM(row['SI Total CBM'] || row['DN Total CBM'] || 0),
                'Number of Boxes': this.calculateBoxes(quantity)
            };
        });
    }

    generateQuickcomReport(data) {
        // Optimized filtering using pre-compiled regex
        const quickcomData = data.filter(row => {
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            return this.compiledRegex.quickcom.test(customerGroup);
        });

        // Quick-com Excel headers
        this.quickcomData = quickcomData.map(row => {
            const quantity = this.getQuantity(row);
            
            return {
                'Customer Group': row['Customer Group'] || '',
                'Transporter Name': row['Transporter'] || '',
                'LR No.': row['SHIPMENT Awb NUMBER'] || '',
                'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || '',
                'Invoice Date': row['SALES Invoice DATE'] || '',
                'Invoice SKU': row['SO Item'] || '',
                'Invoice Qty': quantity,
                'Per Unit CBM': this.formatCBM(row['Per Unit CBM'] || 0),
                'Total CBM': this.formatCBM(row['SI Total CBM'] || row['DN Total CBM'] || 0),
                'Number of Boxes': this.calculateBoxes(quantity)
            };
        });
    }

    generateOfflineReport(data) {
        // Optimized filtering for offline channels
        const offlineData = data.filter(row => {
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            
            // Use pre-compiled regex for faster exclusion check
            return !this.compiledRegex.ecom.test(customerGroup) && 
                   !this.compiledRegex.quickcom.test(customerGroup);
        });

        // Offline Excel headers
        this.offlineData = offlineData.map(row => {
            const quantity = this.getQuantity(row);
            
            return {
                'Customer': row['Customer'] || '',
                'Customer Group': row['Customer Group'] || '',
                'Transporter Name': row['Transporter'] || '',
                'LR No.': row['SHIPMENT Awb NUMBER'] || '',
                'Vehicle No': row['SHIPMENT Vehicle NO'] || '',
                'Sales Order No': row['Sales Order No'] || '',
                'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || '',
                'Invoice Date': row['SALES Invoice DATE'] || '',
                'Invoice SKU': row['SO Item'] || '',
                'Invoice Qty': quantity,
                'Per Unit CBM': this.formatCBM(row['Per Unit CBM'] || 0),
                'Total CBM': this.formatCBM(row['SI Total CBM'] || row['DN Total CBM'] || 0),
                'Pickup Date': row['SHIPMENT Pickup DATE'] || '',
                'Delivered Date': row['DELIVERED Date'] || '',
                'Number of Boxes': this.calculateBoxes(quantity)
            };
        });
    }

    calculateBoxes(qty) {
        // Calculate number of boxes based on quantity
        const quantity = parseFloat(qty) || 0;
        return Math.max(1, Math.ceil(quantity / 20));
    }
    
    generateB2CReport(data) {
        // Filter for B2C channels
        const b2cData = data.filter(row => {
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            return this.compiledRegex.b2c.test(customerGroup);
        });

        // B2C Excel headers
        this.b2cData = b2cData.map(row => {
            const quantity = this.getQuantity(row);
            
            return {
                'Customer Group': row['Customer Group'] || '',
                'Vehicle Series': row['SHIPMENT Vehicle NO'] || '',
                'Dispatch Date': row['SHIPMENT Pickup DATE'] || row['DELIVERY Note DATE'] || '',
                'Customer Name': row['Customer'] || '',
                'Transporter Name': row['Transporter'] || '',
                'Vehicle No': row['SHIPMENT Vehicle NO'] || '',
                'LR No.': row['SHIPMENT Awb NUMBER'] || '',
                'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || '',
                'Invoice Date': row['SALES Invoice DATE'] || '',
                'Invoice SKU': row['SO Item'] || row['Description of Content'] || '',
                'Invoice Qty': quantity,
                'Total CBM': this.formatCBM(row['SI Total CBM'] || row['DN Total CBM'] || 0),
                'Number of Boxes': this.calculateBoxes(quantity)
            };
        });
    }
    
    generateEBOReport(data) {
        // Filter for EBO channels
        const eboData = data.filter(row => {
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            return this.compiledRegex.ebo.test(customerGroup);
        });

        // EBO Excel headers
        this.eboData = eboData.map(row => {
            const quantity = this.getQuantity(row);
            
            return {
                'Store Name': row['Customer Group'] || '',
                'Customer': row['Customer'] || '',
                'Transporter Name': row['Transporter'] || '',
                'LR No.': row['SHIPMENT Awb NUMBER'] || '',
                'Vehicle No': row['SHIPMENT Vehicle NO'] || '',
                'Sales Order No': row['Sales Order No'] || '',
                'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || '',
                'Invoice Date': row['SALES Invoice DATE'] || '',
                'Invoice SKU': row['SO Item'] || '',
                'Invoice Qty': quantity,
                'Per Unit CBM': this.formatCBM(row['Per Unit CBM'] || 0),
                'Total CBM': this.formatCBM(row['SI Total CBM'] || row['DN Total CBM'] || 0),
                'Pickup Date': row['SHIPMENT Pickup DATE'] || '',
                'Delivered Date': row['DELIVERED Date'] || '',
                'Number of Boxes': this.calculateBoxes(quantity)
            };
        });
    }
    
    generateOthersReport(data) {
        // Filter for Others channels
        const othersData = data.filter(row => {
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            return this.compiledRegex.others.test(customerGroup);
        });

        // Others Excel headers
        this.othersData = othersData.map(row => {
            const quantity = this.getQuantity(row);
            
            return {
                'Category': row['Customer Group'] || '',
                'Customer': row['Customer'] || '',
                'Transporter Name': row['Transporter'] || '',
                'LR No.': row['SHIPMENT Awb NUMBER'] || '',
                'Vehicle No': row['SHIPMENT Vehicle NO'] || '',
                'Sales Order No': row['Sales Order No'] || '',
                'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || '',
                'Invoice Date': row['SALES Invoice DATE'] || '',
                'Invoice SKU': row['SO Item'] || '',
                'Invoice Qty': quantity,
                'Total CBM': this.formatCBM(row['SI Total CBM'] || row['DN Total CBM'] || 0),
                'Pickup Date': row['SHIPMENT Pickup DATE'] || '',
                'Delivered Date': row['DELIVERED Date'] || '',
                'Number of Boxes': this.calculateBoxes(quantity)
            };
        });
    }

    displayAllReports() {
        // Sort all data by SO Date before displaying
        this.sortDataByDate();
        
        this.displayReport('b2cTable', this.b2cData, 'b2cCount');
        this.displayReport('ecomTable', this.ecomData, 'ecomCount');
        this.displayReport('offlineTable', this.offlineData, 'offlineCount');
        this.displayReport('quickcomTable', this.quickcomData, 'quickcomCount');
        this.displayReport('eboTable', this.eboData, 'eboCount');
        this.displayReport('othersTable', this.othersData, 'othersCount');
    }

    // Sort data by SO Date only
    sortDataByDate() {
        // Get the selected sort order
        const sortOrderSelect = document.getElementById('sortOrder');
        const sortOrder = sortOrderSelect ? sortOrderSelect.value : 'newest';
        
        const sortByDate = (a, b) => {
            // Get SO Date from both records (only SO Date)
            const dateA = this.parseDate(a['SO Date'] || '');
            const dateB = this.parseDate(b['SO Date'] || '');
            
            // If dates are equal, maintain original order
            if (dateA.getTime() === dateB.getTime()) return 0;
            
            // Sort based on selected order
            if (sortOrder === 'oldest') {
                return dateA.getTime() - dateB.getTime(); // Oldest first
            } else {
                return dateB.getTime() - dateA.getTime(); // Newest first (default)
            }
        };

        // Sort each report data
        this.ecomData.sort(sortByDate);
        this.quickcomData.sort(sortByDate);
        this.offlineData.sort(sortByDate);
    }

    // Parse date string and return Date object
    parseDate(dateStr) {
        if (!dateStr || dateStr === '') {
            return new Date(0); // Return epoch for empty dates (will sort to end)
        }

        try {
            // Handle various date formats
            let date;
            
            // If it's already a Date object
            if (dateStr instanceof Date) {
                date = dateStr;
            }
            // If it's a string, try to parse it
            else if (typeof dateStr === 'string') {
                // Try different date formats
                const formats = [
                    /^\d{4}-\d{2}-\d{2}$/, // YYYY-MM-DD
                    /^\d{2}\/\d{2}\/\d{4}$/, // MM/DD/YYYY
                    /^\d{2}-\d{2}-\d{4}$/, // MM-DD-YYYY
                    /^\d{1,2}\/\d{1,2}\/\d{4}$/, // M/D/YYYY
                    /^\d{1,2}-\d{1,2}-\d{4}$/ // M-D-YYYY
                ];
                
                // Check if it matches any known format
                const isKnownFormat = formats.some(format => format.test(dateStr.trim()));
                
                if (isKnownFormat) {
                    date = new Date(dateStr);
                } else {
                    // Try to parse as-is
                    date = new Date(dateStr);
                }
            } else {
                date = new Date(dateStr);
            }
            
            // Check if the date is valid
            if (isNaN(date.getTime())) {
                console.warn('Invalid date format:', dateStr);
                return new Date(0); // Return epoch for invalid dates
            }
            
            return date;
        } catch (error) {
            console.error('Error parsing date:', dateStr, error);
            return new Date(0); // Return epoch for error cases
        }
    }

    displayReport(tableId, data, countId) {
        const table = document.getElementById(tableId);
        const countElement = document.getElementById(countId);
        
        if (!table) {
            console.error(`Table with ID ${tableId} not found`);
            return;
        }
        
        const thead = table.querySelector('thead');
        const tbody = table.querySelector('tbody');
        
        if (!thead || !tbody) {
            console.error(`Table with ID ${tableId} is missing thead or tbody`);
            return;
        }
        
        // Clear previous content
        thead.innerHTML = '';
        tbody.innerHTML = '';
        
        // Update count
        if (countElement) {
            countElement.textContent = `${data.length} records`;
        }
        
        if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="100%" style="text-align: center; padding: 20px;">No data available</td></tr>';
            return;
        }
        
        // Create headers
        const headers = Object.keys(data[0]);
        const headerRow = document.createElement('tr');
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        
        // Create data rows (show only first 100 rows for performance)
        const displayData = data.slice(0, 100);
        displayData.forEach(row => {
            const tr = document.createElement('tr');
            headers.forEach(header => {
                const td = document.createElement('td');
                
                // Format quantity values consistently
                if (header === 'Invoice Qty') {
                    td.textContent = parseFloat(row[header] || 0).toString();
                } else {
                    td.textContent = row[header] || '';
                }
                
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        // Add note if data was truncated
        if (data.length > 100) {
            const noteRow = document.createElement('tr');
            const noteCell = document.createElement('td');
            noteCell.colSpan = headers.length;
            noteCell.style.textAlign = 'center';
            noteCell.style.fontStyle = 'italic';
            noteCell.style.padding = '15px';
            noteCell.style.backgroundColor = '#f8f9ff';
            noteCell.textContent = `Showing first 100 rows. Total: ${data.length} records. Download Excel to see all data.`;
            noteRow.appendChild(noteCell);
            tbody.appendChild(noteRow);
        }
    }

    downloadAllReports() {
        this.downloadReport('b2c');
        setTimeout(() => this.downloadReport('ecom'), 500);
        setTimeout(() => this.downloadReport('offline'), 1000);
        setTimeout(() => this.downloadReport('quickcom'), 1500);
        setTimeout(() => this.downloadReport('ebo'), 2000);
        setTimeout(() => this.downloadReport('others'), 2500);
    }

    downloadReport(type) {
        let data, filename;
        
        switch(type) {
            case 'b2c':
                data = this.b2cData;
                filename = 'b2c_excel';
                break;
            case 'ecom':
                data = this.ecomData;
                filename = 'e-com_excel';
                break;
            case 'offline':
                data = this.offlineData;
                filename = 'offline_excel';
                break;
            case 'quickcom':
                data = this.quickcomData;
                filename = 'quick-com_excel';
                break;
            case 'ebo':
                data = this.eboData;
                filename = 'ebo_excel';
                break;
            case 'others':
                data = this.othersData;
                filename = 'others_excel';
                break;
            default:
                return;
        }
        
        if (data.length === 0) {
            this.showError(`No data available for ${filename}.`);
            return;
        }

        try {
            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'MIS Report');

            const today = new Date().toISOString().split('T')[0];
            XLSX.writeFile(wb, `${filename}_${today}.xlsx`);
        } catch (error) {
            this.showError(`Error downloading ${filename}: ${error.message}`);
            console.error(error);
        }
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    showError(message) {
        const errorText = document.getElementById('errorText');
        const errorMessage = document.getElementById('errorMessage');
        
        if (errorText && errorMessage) {
            errorText.textContent = message;
            errorMessage.style.display = 'block';
        } else {
            console.error('Error:', message);
        }
    }

    hideError() {
        const errorMessage = document.getElementById('errorMessage');
        if (errorMessage) {
            errorMessage.style.display = 'none';
        }
    }
    
    showFilteringNotification(originalCount, filteredCount) {
        const filteredRows = originalCount - filteredCount;
        if (filteredRows > 0) {
            // Create a temporary notification
            const notification = document.createElement('div');
            notification.style.cssText = `
                position: fixed;
                top: 20px;
                right: 20px;
                background: #fff3cd;
                border: 1px solid #ffeaa7;
                border-radius: 8px;
                padding: 15px 20px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                z-index: 1000;
                max-width: 400px;
                font-family: Arial, sans-serif;
            `;
            
            notification.innerHTML = `
                <div style="display: flex; align-items: center; margin-bottom: 8px;">
                    <span style="font-size: 20px; margin-right: 10px;">⚠️</span>
                    <strong style="color: #856404;">Data Filtering Applied</strong>
                </div>
                <div style="color: #856404; font-size: 14px;">
                    <p style="margin: 0 0 5px 0;">Filtered out <strong>${filteredRows}</strong> rows with negative CBM or quantity values.</p>
                    <p style="margin: 0; font-size: 12px;">Processing <strong>${filteredCount}</strong> valid rows for reports and dashboard.</p>
                </div>
            `;
            
            document.body.appendChild(notification);
            
            // Auto-remove after 8 seconds
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 8000);
        }
    }

    showSuccessNotification(recordCount) {
        // Create a success notification
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: #d4edda;
            border: 1px solid #c3e6cb;
            border-radius: 8px;
            padding: 15px 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 1000;
            max-width: 400px;
            font-family: Arial, sans-serif;
        `;
        
        notification.innerHTML = `
            <div style="display: flex; align-items: center; margin-bottom: 8px;">
                <span style="font-size: 20px; margin-right: 10px;">✅</span>
                <strong style="color: #155724;">Reports Generated Successfully!</strong>
            </div>
            <div style="color: #155724; font-size: 14px;">
                Successfully processed ${recordCount.toLocaleString()} records and generated all reports.
            </div>
        `;
        
        document.body.appendChild(notification);
        
        // Auto-remove after 5 seconds
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 5000);
    }
}

// Global function for download buttons
function downloadReport(type) {
    if (window.auth && !window.auth.canDownload()) {
        if (window.misGenerator) {
            window.misGenerator.showError('You do not have permission to download files. Contact administrator.');
        }
        return;
    }
    
    if (window.misGenerator) {
        window.misGenerator.downloadReport(type);
    } else {
        console.error('MIS Generator not initialized');
    }
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    // Check authentication before initializing
    if (window.auth && window.auth.isUserAuthenticated()) {
        // Update UI based on user permissions
        updateUIForUserPermissions();
        
        window.misGenerator = new ExcelMISGenerator();
    } else {
        // Redirect to login if not authenticated
        window.location.href = 'login.html';
    }
});

// Update UI based on user permissions
function updateUIForUserPermissions() {
    if (!window.auth || !window.auth.isUserAuthenticated()) return;
    
    const user = window.auth.getCurrentUser();
    const accessLevelIndicator = document.getElementById('accessLevelIndicator');
    
    // Update access level indicator
    if (accessLevelIndicator) {
        const accessLevel = user.accessLevel === 'admin' ? 'Admin Access (Read/Write)' : 'User Access (Read-Only)';
        const icon = user.accessLevel === 'admin' ? '🔓' : '🔒';
        accessLevelIndicator.innerHTML = `${icon} ${accessLevel}`;
        accessLevelIndicator.style.color = user.accessLevel === 'admin' ? '#4caf50' : '#ff9800';
    }
    
    // Disable upload area for read-only users
    if (!window.auth.canUpload()) {
        const uploadArea = document.getElementById('uploadArea');
        if (uploadArea) {
            uploadArea.style.opacity = '0.6';
            uploadArea.style.cursor = 'not-allowed';
            uploadArea.title = 'Upload not available - Read-only access';
        }
        
        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            fileInput.disabled = true;
        }
    }
    
    // Disable generate reports button for read-only users
    if (!window.auth.canWrite()) {
        const generateBtn = document.getElementById('generateReports');
        if (generateBtn) {
            generateBtn.disabled = true;
            generateBtn.style.opacity = '0.6';
            generateBtn.style.cursor = 'not-allowed';
            generateBtn.title = 'Report generation not available - Read-only access';
        }
    }
    
    // Disable download buttons for read-only users
    if (!window.auth.canDownload()) {
        const downloadAllBtn = document.getElementById('downloadAll');
        if (downloadAllBtn) {
            downloadAllBtn.disabled = true;
            downloadAllBtn.style.opacity = '0.6';
            downloadAllBtn.style.cursor = 'not-allowed';
            downloadAllBtn.title = 'Download not available - Read-only access';
        }
        
        // Disable individual download buttons
        const downloadBtns = document.querySelectorAll('.btn-download');
        downloadBtns.forEach(btn => {
            btn.disabled = true;
            btn.style.opacity = '0.6';
            btn.style.cursor = 'not-allowed';
            btn.title = 'Download not available - Read-only access';
        });
    }
}

// Method to expose filtered data for LR Pending page
function getFilteredData() {
    if (window.excelMISGenerator) {
        return window.excelMISGenerator.filteredData || [];
    }
    return [];
}

// Initialize LR Missing Dashboard
function initializeLRMissingDashboard() {
    // Hide all sections initially
    const summarySection = document.getElementById('lrMissingSummarySection');
    const tableSection = document.getElementById('lrMissingTableSection');
    const selectedDateInfo = document.getElementById('selectedDateInfo');
    
    if (summarySection) summarySection.style.display = 'none';
    if (tableSection) tableSection.style.display = 'none';
    if (selectedDateInfo) selectedDateInfo.style.display = 'none';
    
    // Reset global variables
    selectedDate = null;
    selectedDateData = [];
    filteredLRData = [];
    currentCategory = 'all';
    currentLRPage = 1;
    
    // Set up date picker
    const datePicker = document.getElementById('lrDatePicker');
    if (datePicker) {
        const today = new Date().toISOString().split('T')[0];
        datePicker.value = '';
        datePicker.max = today;
    }
}

// LR Missing functionality
function showLRMissingSection() {
    const lrMissingSection = document.getElementById('lrMissingSection');
    if (lrMissingSection) {
        lrMissingSection.style.display = 'block';
    }
    
    // Initialize the dashboard properly
    initializeLRMissingDashboard();
}

function updateLRMissingData(data) {
    if (!data || data.length === 0) {
        hideLRMissingData();
        return;
    }

    // Filter for LR Missing records
    const lrMissingData = data.filter(row => {
        const lrNo = row['SHIPMENT Awb NUMBER'] || '';
        return !lrNo || lrNo.toString().trim() === '';
    });

    console.log('LR Missing records found:', lrMissingData.length);
    
    // Ensure dashboard is also updated with the same data
    if (window.dashboard) {
        const dashboardStats = window.dashboard.calculateStats(data);
        console.log('Dashboard LR Pending count:', dashboardStats.lrPending);
        console.log('Counts match:', dashboardStats.lrPending === lrMissingData.length);
    }

    // Update summary counts immediately
    updateLRMissingSummary(lrMissingData);

    // Store all LR missing data globally
    allLRMissingData = lrMissingData.map(row => {
        const customerGroup = (row['Customer Group'] || '').toLowerCase();
        let category = 'offline';
        
        if (customerGroup.includes('amazon') || customerGroup.includes('flipkart')) {
            category = 'ecom';
        } else if (customerGroup.includes('bigbasket') || customerGroup.includes('blinkit') || 
                  customerGroup.includes('zepto') || customerGroup.includes('swiggy')) {
            category = 'quickcom';
        }

        const quantity = getQuantityForLR(row);
        let priority = 'low';
        if (quantity >= 100) priority = 'high';
        else if (quantity >= 10) priority = 'medium';

        return {
            ...row,
            category,
            priority
        };
    });

    // Sort by date (newest first), then by priority, then by quantity
    allLRMissingData.sort((a, b) => {
        const dateA = new Date(a['SO Date'] || '');
        const dateB = new Date(b['SO Date'] || '');
        
        if (dateA.getTime() !== dateB.getTime()) {
            return dateB.getTime() - dateA.getTime();
        }
        
        const priorityOrder = { high: 3, medium: 2, low: 1 };
        const aPriority = priorityOrder[a.priority] || 0;
        const bPriority = priorityOrder[b.priority] || 0;
        
        if (aPriority !== bPriority) {
            return bPriority - aPriority;
        }
        
        return getQuantityForLR(b) - getQuantityForLR(a);
    });

    // Don't wait for date selection - show summary and data immediately
    if (lrMissingData.length > 0) {
        // Set filtered data to all data initially
        filteredLRData = allLRMissingData;
        
        // Show sections immediately with all data
        const summarySection = document.getElementById('lrMissingSummarySection');
        const tableSection = document.getElementById('lrMissingTableSection');
        const recordCount = document.getElementById('recordCount');
        
        if (summarySection) summarySection.style.display = 'block';
        if (tableSection) tableSection.style.display = 'block';
        
        // Update the header record count and show all data
        if (recordCount) {
            recordCount.textContent = `${lrMissingData.length} records`;
        }
        
        // Display all LR missing records grouped by date
        displayLRMissingByDay(lrMissingData);
    } else {
        displayLRMissingByDay(lrMissingData);
    }

    // Initialize filters
    initializeLRMissingFilters(lrMissingData);
}

function updateLRMissingSummary(data) {
    const totalCount = data.length;
    
    // B2C category count
    const b2cCount = data.filter(row => {
        const customerGroup = (row['Customer Group'] || '').toLowerCase();
        return customerGroup.includes('decathlon') || customerGroup.includes('flflipkart(b2c)') || 
               customerGroup.includes('snapmint') || customerGroup.includes('shopify') || 
               customerGroup.includes('tatacliq') || customerGroup.includes('amazon b2c') ||
               customerGroup.includes('pepperfry');
    }).length;
    
    // E-commerce category count
    const ecomCount = data.filter(row => {
        const customerGroup = (row['Customer Group'] || '').toLowerCase();
        return customerGroup.includes('amazon') || customerGroup.includes('flipkart');
    }).length;
    
    // Offline category count
    const offlineCount = data.filter(row => {
        const customerGroup = (row['Customer Group'] || '').toLowerCase();
        return customerGroup.includes('offline sales-b2b') || customerGroup.includes('offline – gt') ||
               customerGroup.includes('offline - mt');
    }).length;
    
    // Quick-commerce category count
    const quickcomCount = data.filter(row => {
        const customerGroup = (row['Customer Group'] || '').toLowerCase();
        return customerGroup.includes('bigbasket') || customerGroup.includes('blinkit') || 
               customerGroup.includes('zepto') || customerGroup.includes('swiggy');
    }).length;
    
    // EBO category count
    const eboCount = data.filter(row => {
        const customerGroup = (row['Customer Group'] || '').toLowerCase();
        return customerGroup.includes('store 2-lucknow') || customerGroup.includes('store3-zirakpur');
    }).length;
    
    // Others category count (everything else not captured in other categories)
    const othersCount = data.filter(row => {
        const customerGroup = (row['Customer Group'] || '').toLowerCase();
        return customerGroup.includes('sales to vendor') || customerGroup.includes('internal company') ||
               customerGroup.includes('others');
    }).length;

    // Update stat cards
    document.getElementById('totalLRMissingCount').textContent = totalCount.toLocaleString();
    document.getElementById('b2cLRMissingCount').textContent = b2cCount.toLocaleString();
    document.getElementById('ecomLRMissingCount').textContent = ecomCount.toLocaleString();
    document.getElementById('offlineLRMissingCount').textContent = offlineCount.toLocaleString();
    document.getElementById('quickcomLRMissingCount').textContent = quickcomCount.toLocaleString();
    document.getElementById('eboLRMissingCount').textContent = eboCount.toLocaleString();
    document.getElementById('othersLRMissingCount').textContent = othersCount.toLocaleString();

    // Update filter dropdown options
    updateFilterCounts(totalCount, b2cCount, ecomCount, offlineCount, quickcomCount, eboCount, othersCount);
}

function updateFilterCounts(total, b2c, ecom, offline, quickcom, ebo, others) {
    const categoryFilter = document.getElementById('lrCategoryFilter');
    if (categoryFilter) {
        categoryFilter.innerHTML = `
            <option value="all">All Categories (${total})</option>
            <option value="b2c">B2C (${b2c})</option>
            <option value="ecom">E-commerce (${ecom})</option>
            <option value="offline">Offline (${offline})</option>
            <option value="quickcom">Quick-commerce (${quickcom})</option>
            <option value="ebo">EBO (${ebo})</option>
            <option value="others">Others (${others})</option>
        `;
    }
}

// Global variables for LR Missing functionality
let currentLRPage = 1;
let lrRecordsPerPage = 50;
let allLRMissingData = [];
let selectedDate = null;
let selectedDateData = [];
let filteredLRData = [];
let currentCategory = 'all';

function displayLRMissingByDay(data) {
    const container = document.getElementById('lrMissingByDay');
    if (!container) return;

    if (data.length === 0) {
        container.innerHTML = '<div class="no-lr-missing">No LR Missing records found. All shipments have LR numbers.</div>';
        return;
    }

    // Store all data globally for pagination
    allLRMissingData = data.map(row => {
        const customerGroup = (row['Customer Group'] || '').toLowerCase();
        let category = 'others'; // Default category
        
        // B2C category
        if (customerGroup.includes('decathlon') || customerGroup.includes('flflipkart(b2c)') || 
            customerGroup.includes('snapmint') || customerGroup.includes('shopify') || 
            customerGroup.includes('tatacliq') || customerGroup.includes('amazon b2c') ||
            customerGroup.includes('pepperfry')) {
            category = 'b2c';
        }
        // E-Commerce category
        else if (customerGroup.includes('amazon') || customerGroup.includes('flipkart')) {
            category = 'ecom';
        }
        // Offline category
        else if (customerGroup.includes('offline sales-b2b') || customerGroup.includes('offline – gt') ||
                customerGroup.includes('offline - mt')) {
            category = 'offline';
        }
        // Quick Commerce category
        else if (customerGroup.includes('bigbasket') || customerGroup.includes('blinkit') || 
                customerGroup.includes('zepto') || customerGroup.includes('swiggy')) {
            category = 'quickcom';
        }
        // EBO category
        else if (customerGroup.includes('store 2-lucknow') || customerGroup.includes('store3-zirakpur')) {
            category = 'ebo';
        }
        // Others category handled by default

        const quantity = getQuantityForLR(row);
        let priority = 'low';
        if (quantity >= 100) priority = 'high';
        else if (quantity >= 10) priority = 'medium';

        return {
            ...row,
            category,
            priority
        };
    });

    // Sort by date (newest first), then by priority, then by quantity
    allLRMissingData.sort((a, b) => {
        const dateA = new Date(a['SO Date'] || '');
        const dateB = new Date(b['SO Date'] || '');
        
        if (dateA.getTime() !== dateB.getTime()) {
            return dateB.getTime() - dateA.getTime();
        }
        
        const priorityOrder = { high: 3, medium: 2, low: 1 };
        const aPriority = priorityOrder[a.priority] || 0;
        const bPriority = priorityOrder[b.priority] || 0;
        
        if (aPriority !== bPriority) {
            return bPriority - aPriority;
        }
        
        return getQuantityForLR(b) - getQuantityForLR(a);
    });

    // Don't auto-display - wait for date selection
    // displayLRMissingPage();
}

function displayLRMissingPage() {
    const container = document.getElementById('lrMissingByDay');
    if (!container) return;

    const startIndex = (currentLRPage - 1) * lrRecordsPerPage;
    const endIndex = startIndex + lrRecordsPerPage;
    const pageData = allLRMissingData.slice(startIndex, endIndex);
    
    const totalPages = Math.ceil(allLRMissingData.length / lrRecordsPerPage);

    // Generate table HTML
    const tableHTML = `
        <div class="table-container">
            <div class="table-header">
                <div class="table-info">
                    Showing ${startIndex + 1}-${Math.min(endIndex, allLRMissingData.length)} of ${allLRMissingData.length} records
                </div>
                <div class="table-actions">
                    <select id="lrRecordsPerPage" onchange="changeLRRecordsPerPage()">
                        <option value="25" ${lrRecordsPerPage === 25 ? 'selected' : ''}>25 per page</option>
                        <option value="50" ${lrRecordsPerPage === 50 ? 'selected' : ''}>50 per page</option>
                        <option value="100" ${lrRecordsPerPage === 100 ? 'selected' : ''}>100 per page</option>
                        <option value="200" ${lrRecordsPerPage === 200 ? 'selected' : ''}>200 per page</option>
                    </select>
                </div>
            </div>
            <div class="lr-missing-content">
                <table class="lr-missing-table">
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Invoice No</th>
                            <th>Customer</th>
                            <th>Customer Group</th>
                            <th>SKU</th>
                            <th>Quantity</th>
                            <th>Transporter</th>
                            <th>Priority</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                <tbody>
                    ${pageData.map(row => `
                        <tr>
                            <td>${formatDateForExport(row['SALES Invoice DATE'] || '')}</td>
                            <td>${row['SALES Invoice NO'] || row['DELIVERY Note NO'] || 'N/A'}</td>
                            <td>${row['Customer'] || 'N/A'}</td>
                            <td>${row['Customer Group'] || 'N/A'}</td>
                            <td>${row['SO Item'] || row['Description of Content'] || 'N/A'}</td>
                            <td class="priority-${row.priority}">${getQuantityForLR(row).toLocaleString()}</td>
                            <td>${row['Transporter'] || 'Not Assigned'}</td>
                            <td><span class="priority-${row.priority}">${row.priority.toUpperCase()}</span></td>
                            <td><span class="status-pending">LR Missing</span></td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            </div>
            ${totalPages > 1 ? `
                <div class="pagination">
                    <button onclick="previousLRPage()" ${currentLRPage === 1 ? 'disabled' : ''}>← Previous</button>
                    <span class="page-info">Page ${currentLRPage} of ${totalPages}</span>
                    <button onclick="nextLRPage()" ${currentLRPage === totalPages ? 'disabled' : ''}>Next →</button>
                </div>
            ` : ''}
        </div>
    `;

    container.innerHTML = tableHTML;
}

// New function to display filtered data
function displayFilteredData() {
    const container = document.getElementById('lrMissingByDay');
    const recordCount = document.getElementById('recordCount');
    
    if (!container) return;
    
    if (filteredLRData.length === 0) {
        container.innerHTML = '<div class="no-lr-missing">No LR Missing records found for the selected criteria.</div>';
        recordCount.textContent = '0 records';
        return;
    }

    // Reset pagination for filtered data
    currentLRPage = 1;
    
    // Update record count
    recordCount.textContent = `${filteredLRData.length} records`;
    
    // Display paginated filtered data
    displayFilteredPage();
}

// Updated function to display paginated filtered data
function displayFilteredPage() {
    const container = document.getElementById('lrMissingByDay');
    if (!container) return;

    const startIndex = (currentLRPage - 1) * lrRecordsPerPage;
    const endIndex = startIndex + lrRecordsPerPage;
    const pageData = filteredLRData.slice(startIndex, endIndex);
    
    const totalPages = Math.ceil(filteredLRData.length / lrRecordsPerPage);

    // Generate table HTML
    const tableHTML = `
        <div class="table-container">
            <div class="table-header">
                <div class="table-info">
                    Showing ${startIndex + 1}-${Math.min(endIndex, filteredLRData.length)} of ${filteredLRData.length} records
                </div>
                <div class="table-actions">
                    <select id="lrRecordsPerPage" onchange="changeLRRecordsPerPage()">
                        <option value="25" ${lrRecordsPerPage === 25 ? 'selected' : ''}>25 per page</option>
                        <option value="50" ${lrRecordsPerPage === 50 ? 'selected' : ''}>50 per page</option>
                        <option value="100" ${lrRecordsPerPage === 100 ? 'selected' : ''}>100 per page</option>
                        <option value="200" ${lrRecordsPerPage === 200 ? 'selected' : ''}>200 per page</option>
                    </select>
                </div>
            </div>
            <div class="lr-missing-content">
                <table class="lr-missing-table">
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Invoice No</th>
                            <th>Customer</th>
                            <th>Customer Group</th>
                            <th>SKU</th>
                            <th>Quantity</th>
                            <th>Transporter</th>
                            <th>Priority</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${pageData.map(row => `
                            <tr>
                                <td>${formatDateForExport(row['SALES Invoice DATE'] || '')}</td>
                                <td>${row['SALES Invoice NO'] || row['DELIVERY Note NO'] || 'N/A'}</td>
                                <td>${row['Customer'] || 'N/A'}</td>
                                <td>${row['Customer Group'] || 'N/A'}</td>
                                <td>${row['SO Item'] || row['Description of Content'] || 'N/A'}</td>
                                <td class="priority-${row.priority}">${getQuantityForLR(row).toLocaleString()}</td>
                                <td>${row['Transporter'] || 'Not Assigned'}</td>
                                <td><span class="priority-${row.priority}">${row.priority.toUpperCase()}</span></td>
                                <td><span class="status-pending">LR Missing</span></td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
            ${totalPages > 1 ? `
                <div class="pagination">
                    <button onclick="previousLRPage()" ${currentLRPage === 1 ? 'disabled' : ''}>← Previous</button>
                    <span class="page-info">Page ${currentLRPage} of ${totalPages}</span>
                    <button onclick="nextLRPage()" ${currentLRPage === totalPages ? 'disabled' : ''}>Next →</button>
                </div>
            ` : ''}
        </div>
    `;

    container.innerHTML = tableHTML;
}

function changeLRRecordsPerPage() {
    const select = document.getElementById('lrRecordsPerPage');
    lrRecordsPerPage = parseInt(select.value);
    currentLRPage = 1;
    
    // Use filtered data if available, otherwise use original display
    if (filteredLRData.length > 0) {
        displayFilteredPage();
    } else {
        displayLRMissingPage();
    }
}

function previousLRPage() {
    if (currentLRPage > 1) {
        currentLRPage--;
        // Use filtered data if available
        if (filteredLRData.length > 0) {
            displayFilteredPage();
        } else {
            displayLRMissingPage();
        }
    }
}

function nextLRPage() {
    const dataLength = filteredLRData.length > 0 ? filteredLRData.length : allLRMissingData.length;
    const totalPages = Math.ceil(dataLength / lrRecordsPerPage);
    if (currentLRPage < totalPages) {
        currentLRPage++;
        // Use filtered data if available
        if (filteredLRData.length > 0) {
            displayFilteredPage();
        } else {
            displayLRMissingPage();
        }
    }
}

function getQuantityForLR(row) {
    if (!row || typeof row !== 'object') return 0;

    try {
        let qty = row['SALES Invoice QTY'];
        
        if (!qty || qty === 0 || qty === '0') {
            qty = row['DELIVERY Note QTY'];
        }
        
        if (qty === null || qty === undefined) return 0;
        
        if (typeof qty === 'string') {
            qty = qty.replace(/[^\d.-]/g, '');
            if (qty === '' || qty === '-') return 0;
        }
        
        const numQty = parseFloat(qty);
        return isNaN(numQty) ? 0 : Math.max(0, numQty);
    } catch (error) {
        return 0;
    }
}

function formatDateForGrouping(dateStr) {
    if (!dateStr) return 'Unknown Date';
    
    try {
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return 'Invalid Date';
        
        return date.toISOString().split('T')[0]; // YYYY-MM-DD format
    } catch (error) {
        return 'Invalid Date';
    }
}

function initializeLRMissingFilters(data) {
    // Category filter
    const categoryFilter = document.getElementById('lrCategoryFilter');
    if (categoryFilter) {
        categoryFilter.addEventListener('change', () => {
            if (selectedDate) {
                filterSelectedDateData();
            } else {
                filterLRMissingData(data);
            }
        });
    }

    // Date filter
    const dateFilter = document.getElementById('lrDateFilter');
    if (dateFilter) {
        dateFilter.addEventListener('change', () => {
            if (selectedDate) {
                filterSelectedDateData();
            } else {
                filterLRMissingData(data);
            }
        });
    }

    // Search filter
    const searchInput = document.getElementById('lrSearchInput');
    if (searchInput) {
        searchInput.addEventListener('input', () => {
            if (selectedDate) {
                filterSelectedDateData();
            } else {
                filterLRMissingData(data);
            }
        });
    }

    // Initialize date picker without auto-selecting date
    const datePicker = document.getElementById('lrDatePicker');
    if (datePicker) {
        const today = new Date().toISOString().split('T')[0];
        datePicker.value = ''; // Don't auto-select today
        datePicker.max = today; // Set max date to today
        // Don't call filterBySelectedDate() automatically
    }
}

function filterBySelectedDate() {
    const datePicker = document.getElementById('lrDatePicker');
    const selectedDateInfo = document.getElementById('selectedDateInfo');
    const summarySection = document.getElementById('lrMissingSummarySection');
    const tableSection = document.getElementById('lrMissingTableSection');
    
    if (!datePicker || !datePicker.value) {
        selectedDate = null;
        selectedDateData = [];
        selectedDateInfo.style.display = 'none';
        summarySection.style.display = 'none';
        tableSection.style.display = 'none';
        return;
    }

    selectedDate = datePicker.value;
    
    // Filter data for selected date
    selectedDateData = allLRMissingData.filter(row => {
        const invoiceDate = row['SO Date'] || '';
        const dateStr = formatDateForGrouping(invoiceDate);
        return dateStr === selectedDate;
    });

    // Update selected date info
    const dateObj = new Date(selectedDate);
    const formattedDate = dateObj.toLocaleDateString('en-GB', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });

    document.getElementById('selectedDateText').textContent = formattedDate;
    document.getElementById('selectedDateCount').textContent = `${selectedDateData.length} records`;
    
    // Show sections
    selectedDateInfo.style.display = 'block';
    
    if (selectedDateData.length > 0) {
        summarySection.style.display = 'block';
        tableSection.style.display = 'block';

        // Update summary for selected date data only
        updateLRMissingSummary(selectedDateData);

        // Reset category filter and display data
        currentCategory = 'all';
        const categoryFilter = document.getElementById('lrCategoryFilter');
        if (categoryFilter) {
            categoryFilter.value = 'all';
        }

        // Set filtered data and display immediately
        filteredLRData = selectedDateData;
        displayFilteredData();
    } else {
        summarySection.style.display = 'none';
        tableSection.style.display = 'none';

        // Update record count to show 0 when no data
        const recordCount = document.getElementById('recordCount');
        if (recordCount) {
            recordCount.textContent = '0 records';
        }
    }
}

function filterSelectedDateData() {
    if (!selectedDate || selectedDateData.length === 0) {
        filteredLRData = [];
        displayFilteredData();
        return;
    }

    // Apply category filter
    if (currentCategory === 'all') {
        filteredLRData = selectedDateData;
    } else {
        filteredLRData = selectedDateData.filter(row => row.category === currentCategory);
    }
    
    // Apply search filter if search input exists
    const searchInput = document.getElementById('lrSearchInput');
    if (searchInput && searchInput.value.trim()) {
        const searchTerm = searchInput.value.toLowerCase().trim();
        filteredLRData = filteredLRData.filter(row => {
            return Object.values(row).some(value => 
                value && value.toString().toLowerCase().includes(searchTerm)
            );
        });
    }
    
    displayFilteredData();
}

// Function to filter by category
function filterByCategory() {
    const categoryFilter = document.getElementById('lrCategoryFilter');
    if (!categoryFilter) return;
    
    currentCategory = categoryFilter.value;
    currentLRPage = 1;
    
    // Apply category filter to selected date data
    if (selectedDateData.length > 0) {
        filterSelectedDateData();
    }
}

function filterLRMissingData(data) {
    const categoryFilter = document.getElementById('lrCategoryFilter')?.value || 'all';
    const dateFilter = document.getElementById('lrDateFilter')?.value || 'all';
    const searchTerm = document.getElementById('lrSearchInput')?.value.toLowerCase() || '';

    let filteredData = data.filter(row => {
        // Category filter
        if (categoryFilter !== 'all') {
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            let rowCategory = 'others'; // Default category
            
            // B2C category
            if (customerGroup.includes('decathlon') || customerGroup.includes('flflipkart(b2c)') || 
                customerGroup.includes('snapmint') || customerGroup.includes('shopify') || 
                customerGroup.includes('tatacliq') || customerGroup.includes('amazon b2c') ||
                customerGroup.includes('pepperfry')) {
                rowCategory = 'b2c';
            }
            // E-Commerce category
            else if (customerGroup.includes('amazon') || customerGroup.includes('flipkart')) {
                rowCategory = 'ecom';
            }
            // Offline category
            else if (customerGroup.includes('offline sales-b2b') || customerGroup.includes('offline – gt') ||
                    customerGroup.includes('offline - mt')) {
                rowCategory = 'offline';
            }
            // Quick Commerce category
            else if (customerGroup.includes('bigbasket') || customerGroup.includes('blinkit') || 
                    customerGroup.includes('zepto') || customerGroup.includes('swiggy')) {
                rowCategory = 'quickcom';
            }
            // EBO category
            else if (customerGroup.includes('store 2-lucknow') || customerGroup.includes('store3-zirakpur')) {
                rowCategory = 'ebo';
            }
            // Others category (already set as default)
            
            if (rowCategory !== categoryFilter) return false;
        }

        // Date filter
        if (dateFilter !== 'all') {
            const invoiceDate = new Date(row['SO Date'] || '');
            const today = new Date();
            const yesterday = new Date(today);
            yesterday.setDate(yesterday.getDate() - 1);
            
            let shouldInclude = false;
            switch (dateFilter) {
                case 'today':
                    shouldInclude = invoiceDate.toDateString() === today.toDateString();
                    break;
                case 'yesterday':
                    shouldInclude = invoiceDate.toDateString() === yesterday.toDateString();
                    break;
                case 'week':
                    const weekAgo = new Date(today);
                    weekAgo.setDate(weekAgo.getDate() - 7);
                    shouldInclude = invoiceDate >= weekAgo;
                    break;
                case 'month':
                    const monthAgo = new Date(today);
                    monthAgo.setMonth(monthAgo.getMonth() - 1);
                    shouldInclude = invoiceDate >= monthAgo;
                    break;
            }
            
            if (!shouldInclude) return false;
        }

        // Search filter
        if (searchTerm) {
            const searchFields = [
                row['SALES Invoice NO'] || '',
                row['DELIVERY Note NO'] || '',
                row['Customer'] || '',
                row['Customer Group'] || ''
            ];
            
            const searchText = searchFields.join(' ').toLowerCase();
            if (!searchText.includes(searchTerm)) return false;
        }

        return true;
    });

    // Update the global data and reset pagination
    allLRMissingData = filteredData;
    currentLRPage = 1;
    displayLRMissingPage();
}

function hideLRMissingData() {
    const container = document.getElementById('lrMissingByDay');
    if (container) {
        container.innerHTML = '<div class="no-lr-missing">No data available. Please upload and process an Excel file first.</div>';
    }

    // Reset summary counts
    document.getElementById('totalLRMissingCount').textContent = '0';
    document.getElementById('ecomLRMissingCount').textContent = '0';
    document.getElementById('quickcomLRMissingCount').textContent = '0';
    document.getElementById('offlineLRMissingCount').textContent = '0';
}

function refreshLRMissing() {
    if (window.excelMISGenerator && window.excelMISGenerator.filteredData) {
        updateLRMissingData(window.excelMISGenerator.filteredData);
    } else {
        hideLRMissingData();
    }
}

function exportLRMissingToExcel() {
    if (!window.auth || !window.auth.canDownload()) {
        alert('You do not have permission to download files. Contact administrator.');
        return;
    }

    if (allLRMissingData.length === 0) {
        alert('No LR Missing records found to export.');
        return;
    }

    try {
        // Show loading message
        const exportBtn = document.getElementById('exportLRMissing');
        const originalText = exportBtn.textContent;
        exportBtn.textContent = '📊 Exporting...';
        exportBtn.disabled = true;

        // Use setTimeout to allow UI to update
        setTimeout(() => {
            try {
                // Prepare data for export (use all data, not just current page)
                const exportData = allLRMissingData.map(row => {
                    // Get category using the same logic as in updateLRMissingSummary for consistency
                    const customerGroup = (row['Customer Group'] || '').toLowerCase();
                    let category = 'Others';
                    
                    // B2C category
                    if (customerGroup.includes('decathlon') || customerGroup.includes('flflipkart(b2c)') || 
                        customerGroup.includes('snapmint') || customerGroup.includes('shopify') || 
                        customerGroup.includes('tatacliq') || customerGroup.includes('amazon b2c') ||
                        customerGroup.includes('pepperfry')) {
                        category = 'B2C';
                    }
                    // E-Commerce category
                    else if (customerGroup.includes('amazon') || customerGroup.includes('flipkart')) {
                        category = 'E-Commerce';
                    }
                    // Offline category
                    else if (customerGroup.includes('offline sales-b2b') || customerGroup.includes('offline – gt') ||
                            customerGroup.includes('offline - mt')) {
                        category = 'Offline';
                    }
                    // Quick Commerce category
                    else if (customerGroup.includes('bigbasket') || customerGroup.includes('blinkit') || 
                            customerGroup.includes('zepto') || customerGroup.includes('swiggy')) {
                        category = 'Quick-Commerce';
                    }
                    // EBO category
                    else if (customerGroup.includes('store 2-lucknow') || customerGroup.includes('store3-zirakpur')) {
                        category = 'EBO';
                    }

                    const quantity = getQuantityForLR(row);
                    let priority = 'Low';
                    if (quantity >= 100) priority = 'High';
                    else if (quantity >= 10) priority = 'Medium';

                    return {
                        'Date': formatDateForExport(row['SO Date'] || ''),
                        'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || 'N/A',
                        'Customer': row['Customer'] || 'N/A',
                        'Customer Group': row['Customer Group'] || 'N/A',
                        'Category': category,
                        'SKU': row['SO Item'] || row['Description of Content'] || 'N/A',
                        'Quantity': quantity,
                        'CBM': parseFloat(row['SI Total CBM'] || row['DN Total CBM'] || 0).toFixed(2),
                        'Per Unit CBM': parseFloat(row['Per Unit CBM'] || 0).toFixed(2),
                        'Transporter': row['Transporter'] || 'N/A',
                        'Vehicle No': row['SHIPMENT Vehicle NO'] || 'N/A',
                        'Pickup Date': formatDateForExport(row['SHIPMENT Pickup DATE'] || ''),
                        'Delivered Date': formatDateForExport(row['DELIVERED Date'] || ''),
                        'Sales Order No': row['Sales Order No'] || 'N/A',
                        'Priority': priority,
                        'Status': 'LR Missing'
                    };
                });

                // Create workbook
                const ws = XLSX.utils.json_to_sheet(exportData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'LR Missing Report');

                // Generate filename
                const today = new Date().toISOString().split('T')[0];
                const filename = `LR_Missing_Report_${today}.xlsx`;

                // Download file
                XLSX.writeFile(wb, filename);
                
                alert(`LR Missing report exported successfully as ${filename} (${exportData.length} records)`);
            } catch (error) {
                console.error('Export error:', error);
                alert('Error exporting data: ' + error.message);
            } finally {
                // Reset button
                exportBtn.textContent = originalText;
                exportBtn.disabled = false;
            }
        }, 100);

    } catch (error) {
        console.error('Export error:', error);
        alert('Error exporting data: ' + error.message);
    }
}

function exportSelectedDateLRMissing() {
    if (!window.auth || !window.auth.canDownload()) {
        alert('You do not have permission to download files. Contact administrator.');
        return;
    }

    if (!selectedDate || selectedDateData.length === 0) {
        alert('No data available for the selected date. Please select a date with LR missing records.');
        return;
    }

    try {
        // Show loading message
        const exportBtn = document.getElementById('exportSelectedDate');
        const originalText = exportBtn.textContent;
        exportBtn.textContent = '📅 Exporting...';
        exportBtn.disabled = true;

        // Use setTimeout to allow UI to update
        setTimeout(() => {
            try {
                // Prepare data for export (use selected date data)
                const exportData = selectedDateData.map(row => {
                    // Get category using the same logic as in updateLRMissingSummary for consistency
                    const customerGroup = (row['Customer Group'] || '').toLowerCase();
                    let category = 'Others';
                    
                    // B2C category
                    if (customerGroup.includes('decathlon') || customerGroup.includes('flflipkart(b2c)') || 
                        customerGroup.includes('snapmint') || customerGroup.includes('shopify') || 
                        customerGroup.includes('tatacliq') || customerGroup.includes('amazon b2c') ||
                        customerGroup.includes('pepperfry')) {
                        category = 'B2C';
                    }
                    // E-Commerce category
                    else if (customerGroup.includes('amazon') || customerGroup.includes('flipkart')) {
                        category = 'E-Commerce';
                    }
                    // Offline category
                    else if (customerGroup.includes('offline sales-b2b') || customerGroup.includes('offline – gt') ||
                            customerGroup.includes('offline - mt')) {
                        category = 'Offline';
                    }
                    // Quick Commerce category
                    else if (customerGroup.includes('bigbasket') || customerGroup.includes('blinkit') || 
                            customerGroup.includes('zepto') || customerGroup.includes('swiggy')) {
                        category = 'Quick-Commerce';
                    }
                    // EBO category
                    else if (customerGroup.includes('store 2-lucknow') || customerGroup.includes('store3-zirakpur')) {
                        category = 'EBO';
                    }

                    const quantity = getQuantityForLR(row);
                    let priority = 'Low';
                    if (quantity >= 100) priority = 'High';
                    else if (quantity >= 10) priority = 'Medium';

                    return {
                        'Date': formatDateForExport(row['SO Date'] || ''),
                        'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || 'N/A',
                        'Customer': row['Customer'] || 'N/A',
                        'Customer Group': row['Customer Group'] || 'N/A',
                        'Category': category,
                        'SKU': row['SO Item'] || row['Description of Content'] || 'N/A',
                        'Quantity': quantity,
                        'CBM': parseFloat(row['SI Total CBM'] || row['DN Total CBM'] || 0).toFixed(2),
                        'Per Unit CBM': parseFloat(row['Per Unit CBM'] || 0).toFixed(2),
                        'Transporter': row['Transporter'] || 'N/A',
                        'Vehicle No': row['SHIPMENT Vehicle NO'] || 'N/A',
                        'Pickup Date': formatDateForExport(row['SHIPMENT Pickup DATE'] || ''),
                        'Delivered Date': formatDateForExport(row['DELIVERED Date'] || ''),
                        'Sales Order No': row['Sales Order No'] || 'N/A',
                        'Priority': priority,
                        'Status': 'LR Missing'
                    };
                });

                // Create workbook
                const ws = XLSX.utils.json_to_sheet(exportData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'LR Missing Report');

                // Generate filename with selected date
                const dateStr = selectedDate.replace(/-/g, '');
                const filename = `LR_Missing_${dateStr}.xlsx`;

                // Download file
                XLSX.writeFile(wb, filename);
                
                alert(`LR Missing report for ${formatDateForExport(selectedDate)} exported successfully as ${filename} (${exportData.length} records)`);
            } catch (error) {
                console.error('Export error:', error);
                alert('Error exporting data: ' + error.message);
            } finally {
                // Reset button
                exportBtn.textContent = originalText;
                exportBtn.disabled = false;
            }
        }, 100);

    } catch (error) {
        console.error('Export error:', error);
        alert('Error exporting data: ' + error.message);
    }
}

function formatDateForExport(dateStr) {
    if (!dateStr) return 'N/A';
    
    try {
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return dateStr;
        
        return date.toLocaleDateString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        });
    } catch (error) {
        return dateStr;
    }
}

// Download filtered LR Missing data
function downloadLRMissingExcel() {
    if (!filteredLRData || filteredLRData.length === 0) {
        alert('No data available to export. Please select a date first.');
        return;
    }

    const downloadBtn = document.getElementById('downloadLRMissing');
    if (downloadBtn) {
        downloadBtn.textContent = '📥 Exporting...';
        downloadBtn.disabled = true;
    }

    try {
        // Prepare data for export
        const exportData = filteredLRData.map(row => {
            const quantity = getQuantityForLR(row);

            // Get category using the same logic as in updateLRMissingSummary for consistency
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            let category = 'Others';
            
            // B2C category
            if (customerGroup.includes('decathlon') || customerGroup.includes('flflipkart(b2c)') || 
                customerGroup.includes('snapmint') || customerGroup.includes('shopify') || 
                customerGroup.includes('tatacliq') || customerGroup.includes('amazon b2c') ||
                customerGroup.includes('pepperfry')) {
                category = 'B2C';
            }
            // E-Commerce category
            else if (customerGroup.includes('amazon') || customerGroup.includes('flipkart')) {
                category = 'E-Commerce';
            }
            // Offline category
            else if (customerGroup.includes('offline sales-b2b') || customerGroup.includes('offline – gt') ||
                    customerGroup.includes('offline - mt')) {
                category = 'Offline';
            }
            // Quick Commerce category
            else if (customerGroup.includes('bigbasket') || customerGroup.includes('blinkit') || 
                    customerGroup.includes('zepto') || customerGroup.includes('swiggy')) {
                category = 'Quick-Commerce';
            }
            // EBO category
            else if (customerGroup.includes('store 2-lucknow') || customerGroup.includes('store3-zirakpur')) {
                category = 'EBO';
            }

            return {
                'Date': formatDateForExport(row['SO Date'] || ''),
                'Invoice No': row['SALES Invoice NO'] || row['DELIVERY Note NO'] || 'N/A',
                'Customer': row['Customer'] || 'N/A',
                'Customer Group': row['Customer Group'] || 'N/A',
                'Category': category,
                'SKU': row['SO Item'] || row['Description of Content'] || 'N/A',
                'Quantity': quantity,
                'CBM': parseFloat(row['SI Total CBM'] || row['DN Total CBM'] || 0).toFixed(2),
                'Per Unit CBM': parseFloat(row['Per Unit CBM'] || 0).toFixed(2),
                'Transporter': row['Transporter'] || 'N/A',
                'Vehicle No': row['SHIPMENT Vehicle NO'] || 'N/A',
                'Pickup Date': formatDateForExport(row['SHIPMENT Pickup DATE'] || ''),
                'Delivered Date': formatDateForExport(row['DELIVERED Date'] || ''),
                'Sales Order No': row['Sales Order No'] || 'N/A',
                'Priority': row.priority ? (row.priority.charAt(0).toUpperCase() + row.priority.slice(1)) : 'Low',
                'Status': 'LR Missing'
            };
        });

        // Create workbook
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'LR Missing Report');

        // Auto-size columns
        const colWidths = [
            { width: 12 }, // Date
            { width: 15 }, // Category
            { width: 20 }, // Invoice No
            { width: 25 }, // Customer
            { width: 20 }, // Customer Group
            { width: 10 }, // Quantity
            { width: 8 },  // CBM
            { width: 10 }, // Priority
            { width: 12 }, // Status
            { width: 20 }, // Transporter
            { width: 15 }, // Vehicle No
            { width: 12 }, // Pickup Date
            { width: 12 }  // Delivery Date
        ];
        ws['!cols'] = colWidths;

        // Generate filename
        const selectedDateFormatted = selectedDate ? selectedDate.replace(/-/g, '_') : new Date().toISOString().split('T')[0].replace(/-/g, '_');
        const categoryText = currentCategory === 'all' ? 'All_Categories' : currentCategory.charAt(0).toUpperCase() + currentCategory.slice(1);
        const filename = `LR_Missing_Report_${selectedDateFormatted}_${categoryText}.xlsx`;

        // Download file
        XLSX.writeFile(wb, filename);
        
        alert(`LR Missing report exported successfully!\n\nFile: ${filename}\nRecords: ${exportData.length}\nDate: ${selectedDate || 'All dates'}\nCategory: ${categoryText}`);
    } catch (error) {
        console.error('Export error:', error);
        alert('Error exporting data: ' + error.message);
    } finally {
        if (downloadBtn) {
            downloadBtn.textContent = '📥 Download LR Missing Report';
            downloadBtn.disabled = false;
        }
    }
}

// Performance monitoring utility
class PerformanceMonitor {
    constructor() {
        this.metrics = new Map();
    }
    
    start(label) {
        this.metrics.set(label, performance.now());
    }
    
    end(label) {
        const startTime = this.metrics.get(label);
        if (startTime) {
            const duration = performance.now() - startTime;
            console.log(`⚡ ${label}: ${duration.toFixed(2)}ms`);
            this.metrics.delete(label);
            return duration;
        }
        return 0;
    }
    
    measure(label, fn) {
        this.start(label);
        const result = fn();
        this.end(label);
        return result;
    }
}