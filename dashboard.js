class MISDashboard {
    constructor() {
        this.processedData = null;
        this.currentView = 'date'; // 'date' or 'month'
        
        // Performance optimizations
        this.statsCache = new Map();
        this.lastDataHash = null;
        this.compiledRegex = {
            b2c: /decathlon|flflipkart\(b2c\)|snapmint|shopify|tatacliq|amazon b2c|pepperfry/i,
            ecom: /amazon|flipkart/i,
            offline: /offline sales-b2b|offline – gt|offline - mt/i,
            quickcom: /blinkit|swiggy|bigbasket|zepto/i,
            ebo: /store 2-lucknow|store3-zirakpur/i,
            others: /sales to vendor|internal company|others/i
        };
        this.excludedPlatforms = ['flipkart', 'amazon', 'bigbasket', 'blinkit', 'zepto', 'swiggy'];
        
        // Warehouse location mapping
        this.warehouseLocationMapping = {
            // Mumbai warehouses
            'MH4': 'Mumbai',
            'MH5': 'Mumbai',
            // Gurgaon warehouses
            'HR3': 'Gurgaon',
            'HR10': 'Gurgaon',
            'HR11': 'Gurgaon',
            // Howrah warehouses
            'WB4': 'Howrah',
            // Bangalore warehouses
            'KA3': 'Bangalore',
            'KA4': 'Bangalore',
            // Ludhiana warehouses
            'PB2': 'Ludhiana'
        };
        
        // Add flexible code pattern matching - this will catch variants with different formatting
        this.warehouseCodePatterns = {
            mumbai: /\b(MH[0-9]+)/i,
            gurgaon: /\b(HR[0-9]+)/i,
            howrah: /\b(WB[0-9]+)/i,
            bangalore: /\b(KA[0-9]+)/i,
            ludhiana: /\b(PB[0-9]+)/i
        };
        
        // Transporter mapping for FTL/PTL classification
        this.transporterMapping = {
            FTL: [
                "Loadit Supply Services Pvt Ltd",
                "SR ENTERPRISES", 
                "Surya Freight Carrier",
                "Self Pick Up",
                "Loadit"
            ],
            PTL: [
                "SKYLARK EXPRESS (DELHI) PVT LTD",
                "DTDC EXPRESS LIMITED",
                "Safexpress Private Limited", 
                "V TRANS (INDIA) LIMITED",
                "DTDC Biker",
                "Safexpress",
                "DTDC",
                "Only Invoicing",
                "Skylark",
                "Self",
                "Safe Xpress",
                "V-Trans"
            ]
        };
        
        // Initialize warehouse filter data
        this.locations = [];
        this.areaCodes = {};
        this.warehouses = {};
        
        this.initializeUI();
        this.initializeOrderView();
        this.initializeEventListeners();
    }
    
    initializeUI() {
        const dashboardDiv = document.getElementById('dashboardSection');
        if (!dashboardDiv) {
            console.error("Dashboard section not found");
            return;
        }
        
        // Create the modern dashboard layout
        dashboardDiv.innerHTML = `
            <div class="filter-section">
                <div class="filter-group">
                    <label>View By:</label>
                    <div class="toggle-buttons">
                        <button id="dateViewBtn" class="toggle-btn active">Date</button>
                        <button id="monthViewBtn" class="toggle-btn">Month</button>
                    </div>
                </div>
                
                <div id="dateFilterGroup" class="filter-group">
                    <label for="datePicker">Select Date:</label>
                    <input type="date" id="datePicker" class="date-picker">
                </div>
                
                <div id="monthFilterGroup" class="filter-group" style="display: none;">
                    <label for="monthPicker">Select Month:</label>
                    <input type="month" id="monthPicker" class="month-picker">
                </div>
            </div>
            
            <div class="dashboard-grid">
                <!-- Total Card -->
                <div class="dashboard-card total-card">
                    <div class="card-header">
                        <h3>📊 Total</h3>
                        <div class="card-icon">🎯</div>
                    </div>
                    <div class="card-body">
                        <div class="metric">
                            <div class="metric-label">Sales Order Qty</div>
                            <div class="metric-value" id="totalSOQty">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO Numbers</div>
                            <div class="metric-value" id="totalSONumbers">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total Invoices</div>
                            <div class="metric-value" id="totalInvoices">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Invoice Quantity</div>
                            <div class="metric-value" id="totalQuantity">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO-SI Qty Diff</div>
                            <div class="metric-value" id="totalQtyDiff">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total CBM</div>
                            <div class="metric-value" id="totalCBM">0.00</div>
                        </div>
                        <div class="metric" data-metric="lr-pending">
                            <div class="metric-label">LR Pending</div>
                            <div class="metric-value" id="totalLRPending">0</div>
                        </div>
                    </div>
                </div>
                
                <!-- B2C Card -->
                <div class="dashboard-card b2c-card">
                    <div class="card-header">
                        <h3>👤 B2C</h3>
                        <div class="card-icon">🛍️</div>
                    </div>
                    <div class="card-body">
                        <div class="metric">
                            <div class="metric-label">Sales Order Qty</div>
                            <div class="metric-value" id="b2cSOQty">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO Numbers</div>
                            <div class="metric-value" id="b2cSONumbers">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total Invoices</div>
                            <div class="metric-value" id="b2cInvoices">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Invoice Quantity</div>
                            <div class="metric-value" id="b2cQuantity">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO-SI Qty Diff</div>
                            <div class="metric-value" id="b2cQtyDiff">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total CBM</div>
                            <div class="metric-value" id="b2cCBM">0.00</div>
                        </div>
                        <div class="metric" data-metric="lr-pending">
                            <div class="metric-label">LR Pending</div>
                            <div class="metric-value" id="b2cLRPending">0</div>
                        </div>
                    </div>
                </div>
                
                <!-- E-Commerce Card -->
                <div class="dashboard-card ecom-card">
                    <div class="card-header">
                        <h3>🛒 E-Commerce</h3>
                        <div class="card-icon">📱</div>
                    </div>
                    <div class="card-body">
                        <div class="metric">
                            <div class="metric-label">Sales Order Qty</div>
                            <div class="metric-value" id="ecomSOQty">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO Numbers</div>
                            <div class="metric-value" id="ecomSONumbers">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total Invoices</div>
                            <div class="metric-value" id="ecomInvoices">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Invoice Quantity</div>
                            <div class="metric-value" id="ecomQuantity">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO-SI Qty Diff</div>
                            <div class="metric-value" id="ecomQtyDiff">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total CBM</div>
                            <div class="metric-value" id="ecomCBM">0.00</div>
                        </div>
                        <div class="metric" data-metric="lr-pending">
                            <div class="metric-label">LR Pending</div>
                            <div class="metric-value" id="ecomLRPending">0</div>
                        </div>
                    </div>
                </div>
                
                <!-- Offline Card -->
                <div class="dashboard-card offline-card">
                    <div class="card-header">
                        <h3>🏪 Offline</h3>
                        <div class="card-icon">🏬</div>
                    </div>
                    <div class="card-body">
                        <div class="metric">
                            <div class="metric-label">Sales Order Qty</div>
                            <div class="metric-value" id="offlineSOQty">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO Numbers</div>
                            <div class="metric-value" id="offlineSONumbers">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total Invoices</div>
                            <div class="metric-value" id="offlineInvoices">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Invoice Quantity</div>
                            <div class="metric-value" id="offlineQuantity">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO-SI Qty Diff</div>
                            <div class="metric-value" id="offlineQtyDiff">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total CBM</div>
                            <div class="metric-value" id="offlineCBM">0.00</div>
                        </div>
                        <div class="metric" data-metric="lr-pending">
                            <div class="metric-label">LR Pending</div>
                            <div class="metric-value" id="offlineLRPending">0</div>
                        </div>
                    </div>
                </div>
                
                <!-- Quick-Commerce Card -->
                <div class="dashboard-card quickcom-card">
                    <div class="card-header">
                        <h3>⚡ Quick-Commerce</h3>
                        <div class="card-icon">🚀</div>
                    </div>
                    <div class="card-body">
                        <div class="metric">
                            <div class="metric-label">Sales Order Qty</div>
                            <div class="metric-value" id="quickcomSOQty">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO Numbers</div>
                            <div class="metric-value" id="quickcomSONumbers">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total Invoices</div>
                            <div class="metric-value" id="quickcomInvoices">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Invoice Quantity</div>
                            <div class="metric-value" id="quickcomQuantity">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO-SI Qty Diff</div>
                            <div class="metric-value" id="quickcomQtyDiff">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total CBM</div>
                            <div class="metric-value" id="quickcomCBM">0.00</div>
                        </div>
                        <div class="metric" data-metric="lr-pending">
                            <div class="metric-label">LR Pending</div>
                            <div class="metric-value" id="quickcomLRPending">0</div>
                        </div>
                    </div>
                </div>
                
                <!-- EBO Card -->
                <div class="dashboard-card ebo-card">
                    <div class="card-header">
                        <h3>🏢 EBO</h3>
                        <div class="card-icon">🏢</div>
                    </div>
                    <div class="card-body">
                        <div class="metric">
                            <div class="metric-label">Sales Order Qty</div>
                            <div class="metric-value" id="eboSOQty">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO Numbers</div>
                            <div class="metric-value" id="eboSONumbers">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total Invoices</div>
                            <div class="metric-value" id="eboInvoices">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Invoice Quantity</div>
                            <div class="metric-value" id="eboQuantity">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO-SI Qty Diff</div>
                            <div class="metric-value" id="eboQtyDiff">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total CBM</div>
                            <div class="metric-value" id="eboCBM">0.00</div>
                        </div>
                        <div class="metric" data-metric="lr-pending">
                            <div class="metric-label">LR Pending</div>
                            <div class="metric-value" id="eboLRPending">0</div>
                        </div>
                    </div>
                </div>
                
                <!-- Others Card -->
                <div class="dashboard-card others-card">
                    <div class="card-header">
                        <h3>📋 Others</h3>
                        <div class="card-icon">📋</div>
                    </div>
                    <div class="card-body">
                        <div class="metric">
                            <div class="metric-label">Sales Order Qty</div>
                            <div class="metric-value" id="othersSOQty">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO Numbers</div>
                            <div class="metric-value" id="othersSONumbers">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total Invoices</div>
                            <div class="metric-value" id="othersInvoices">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Invoice Quantity</div>
                            <div class="metric-value" id="othersQuantity">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">SO-SI Qty Diff</div>
                            <div class="metric-value" id="othersQtyDiff">0</div>
                        </div>
                        <div class="metric">
                            <div class="metric-label">Total CBM</div>
                            <div class="metric-value" id="othersCBM">0.00</div>
                        </div>
                        <div class="metric" data-metric="lr-pending">
                            <div class="metric-label">LR Pending</div>
                            <div class="metric-value" id="othersLRPending">0</div>
                        </div>
                    </div>
                </div>
            </div>
            
            <div id="dashboardNoData" class="no-data-message" style="display: none;">
                <div class="no-data-content">
                    <div class="no-data-icon">📄</div>
                    <h3>No Data Available</h3>
                    <p>Please upload an Excel file to view dashboard metrics.</p>
                </div>
            </div>
        `;
    }
    
    initializeEventListeners() {
        // Toggle between date and month view
        const dateViewBtn = document.getElementById('dateViewBtn');
        const monthViewBtn = document.getElementById('monthViewBtn');
        
        if (dateViewBtn) {
            dateViewBtn.addEventListener('click', () => this.switchView('date'));
        }
        
        if (monthViewBtn) {
            monthViewBtn.addEventListener('click', () => this.switchView('month'));
        }
        
        // Date picker change handler
        const datePicker = document.getElementById('datePicker');
        if (datePicker) {
            datePicker.addEventListener('change', (e) => {
                this.filterDataByDate(e.target.value);
            });
        }
        
        // Month picker change handler
        const monthPicker = document.getElementById('monthPicker');
        if (monthPicker) {
            monthPicker.addEventListener('change', (e) => {
                this.filterDataByMonth(e.target.value);
            });
        }
        
        // Location filter change handler
        const locationFilter = document.getElementById('locationFilter');
        if (locationFilter) {
            locationFilter.addEventListener('change', () => {
                this.updateAreaCodeFilter();
                this.applyWarehouseFilters();
            });
        }
        
        // Area code filter change handler
        const areaCodeFilter = document.getElementById('areaCodeFilter');
        if (areaCodeFilter) {
            areaCodeFilter.addEventListener('change', () => {
                this.updateWarehouseFilter();
                this.applyWarehouseFilters();
            });
        }
        
        // Warehouse filter change handler
        const warehouseFilter = document.getElementById('warehouseFilter');
        if (warehouseFilter) {
            warehouseFilter.addEventListener('change', () => {
                this.applyWarehouseFilters();
            });
        }
        
        // Sort order change handler
        const sortOrderSelect = document.getElementById('sortOrder');
        if (sortOrderSelect) {
            sortOrderSelect.addEventListener('change', () => {
                if (this.processedData && this.processedData.length > 0) {
                    // Re-apply current filter with new sorting
                    if (this.currentView === 'date') {
                        const datePicker = document.getElementById('datePicker');
                        if (datePicker && datePicker.value) {
                            this.filterDataByDate(datePicker.value);
                        } else {
                            this.updateDashboard(this.processedData);
                        }
                    } else if (this.currentView === 'month') {
                        const monthPicker = document.getElementById('monthPicker');
                        if (monthPicker && monthPicker.value) {
                            this.filterDataByMonth(monthPicker.value);
                        } else {
                            this.updateDashboard(this.processedData);
                        }
                    }
                }
            });
        }
        
        // Set default dates
        const today = new Date();
        if (datePicker) {
            datePicker.valueAsDate = today;
        }
        
        if (monthPicker) {
            const monthValue = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
            monthPicker.value = monthValue;
        }
    }
    
    switchView(viewType) {
        this.currentView = viewType;
        
        const dateViewBtn = document.getElementById('dateViewBtn');
        const monthViewBtn = document.getElementById('monthViewBtn');
        const dateFilterGroup = document.getElementById('dateFilterGroup');
        const monthFilterGroup = document.getElementById('monthFilterGroup');
        
        if (!dateViewBtn || !monthViewBtn || !dateFilterGroup || !monthFilterGroup) {
            console.error("Some view elements not found");
            return;
        }
        
        // Update UI
        if (viewType === 'date') {
            dateViewBtn.classList.add('active');
            monthViewBtn.classList.remove('active');
            dateFilterGroup.style.display = 'flex';
            monthFilterGroup.style.display = 'none';
            
            // Apply date filter
            const dateValue = document.getElementById('datePicker').value;
            if (dateValue) {
                this.filterDataByDate(dateValue);
            }
        } else {
            dateViewBtn.classList.remove('active');
            monthViewBtn.classList.add('active');
            dateFilterGroup.style.display = 'none';
            monthFilterGroup.style.display = 'flex';
            
            // Apply month filter
            const monthValue = document.getElementById('monthPicker').value;
            if (monthValue) {
                this.filterDataByMonth(monthValue);
            }
        }
    }
    
    // Apply warehouse filters to the data
    applyWarehouseFilters() {
        if (!this.processedData || !Array.isArray(this.processedData)) return;
        
        // Get filter values
        const locationFilter = document.getElementById('locationFilter');
        const areaCodeFilter = document.getElementById('areaCodeFilter');
        const warehouseFilter = document.getElementById('warehouseFilter');
        
        if (!locationFilter || !areaCodeFilter || !warehouseFilter) return;
        
        const selectedLocation = locationFilter.value;
        const selectedAreaCode = areaCodeFilter.value;
        const selectedWarehouse = warehouseFilter.value;
        
        // Apply current view filter (date or month)
        if (this.currentView === 'date') {
            const datePicker = document.getElementById('datePicker');
            if (datePicker && datePicker.value) {
                this.filterDataByDate(datePicker.value);
            } else {
                this.updateDashboardWithWarehouseFilters(this.processedData);
            }
        } else if (this.currentView === 'month') {
            const monthPicker = document.getElementById('monthPicker');
            if (monthPicker && monthPicker.value) {
                this.filterDataByMonth(monthPicker.value);
            } else {
                this.updateDashboardWithWarehouseFilters(this.processedData);
            }
        }
    }

    // Return current data after applying warehouse/location/area filters WITHOUT updating UI
    getWarehouseFilteredData(baseData = null) {
        const data = baseData || this.processedData || [];
        const locationFilter = document.getElementById('locationFilter');
        const areaCodeFilter = document.getElementById('areaCodeFilter');
        const warehouseFilter = document.getElementById('warehouseFilter');
        if (!locationFilter || !areaCodeFilter || !warehouseFilter) return data;
        const selectedLocation = locationFilter.value;
        const selectedAreaCode = areaCodeFilter.value;
        const selectedWarehouse = warehouseFilter.value;
        let filtered = [...data];
        if (selectedLocation !== 'all') {
            filtered = filtered.filter(row => {
                const warehouse = row && row['Set Source Warehouse'];
                if (!warehouse) return false;
                const code = this.getWarehouseCode(warehouse);
                return this.getWarehouseLocation(code) === selectedLocation;
            });
            if (selectedAreaCode !== 'all') {
                filtered = filtered.filter(row => {
                    const warehouse = row && row['Set Source Warehouse'];
                    if (!warehouse) return false;
                    const match = warehouse.match(/\b([A-Z]{2}\d+)/);
                    return match && match[1] === selectedAreaCode;
                });
                if (selectedWarehouse !== 'all') {
                    filtered = filtered.filter(r => (r['Set Source Warehouse'] || '') === selectedWarehouse);
                }
            }
        }
        return filtered;
    }
    
    // Apply warehouse filters to data and update dashboard
    updateDashboardWithWarehouseFilters(data) {
        if (!data || !Array.isArray(data)) return;
        
        // Get filter values
        const locationFilter = document.getElementById('locationFilter');
        const areaCodeFilter = document.getElementById('areaCodeFilter');
        const warehouseFilter = document.getElementById('warehouseFilter');
        
        if (!locationFilter || !areaCodeFilter || !warehouseFilter) {
            // If filters not found, update dashboard with all data
            this.updateDashboard(data);
            return;
        }
        
        const selectedLocation = locationFilter.value;
        const selectedAreaCode = areaCodeFilter.value;
        const selectedWarehouse = warehouseFilter.value;
        
        // Filter data based on selected filters
        let filteredData = [...data];
        
        if (selectedLocation !== 'all') {
            filteredData = filteredData.filter(row => {
                if (!row || typeof row !== 'object') return false;
                
                const warehouse = row['Set Source Warehouse'] || '';
                if (!warehouse) return false;
                
                const warehouseCode = this.getWarehouseCode(warehouse);
                const location = this.getWarehouseLocation(warehouseCode);
                
                return location === selectedLocation;
            });
            
            if (selectedAreaCode !== 'all') {
                filteredData = filteredData.filter(row => {
                    if (!row || typeof row !== 'object') return false;
                    
                    const warehouse = row['Set Source Warehouse'] || '';
                    if (!warehouse) return false;
                    
                    // Extract area code
                    const areaCodeMatch = warehouse.match(/\b([A-Z]{2}\d+)/);
                    if (!areaCodeMatch) return false;
                    
                    return areaCodeMatch[1] === selectedAreaCode;
                });
                
                if (selectedWarehouse !== 'all') {
                    filteredData = filteredData.filter(row => {
                        if (!row || typeof row !== 'object') return false;
                        
                        const warehouse = row['Set Source Warehouse'] || '';
                        return warehouse === selectedWarehouse;
                    });
                }
            }
        }
        
        // Sort filtered data by date
        filteredData = this.sortDataByDate(filteredData);
        
        // Update dashboard with filtered data
        this.updateDashboard(filteredData);
    }
    
    setData(data) {
        // Store the raw data
        this.processedData = data;
        
        const noDataMessage = document.getElementById('dashboardNoData');
        if (noDataMessage) {
            noDataMessage.style.display = 'none';
        }
        
        // Extract warehouse structure and populate filters
        this.extractWarehouseStructure(data);
        this.populateWarehouseFilters();
        
        // Apply the current filter
        if (this.currentView === 'date') {
            const datePicker = document.getElementById('datePicker');
            if (datePicker && datePicker.value) {
                this.filterDataByDate(datePicker.value);
            } else {
                this.updateDashboardWithWarehouseFilters(this.processedData);
            }
        } else {
            const monthPicker = document.getElementById('monthPicker');
            if (monthPicker && monthPicker.value) {
                this.filterDataByMonth(monthPicker.value);
            } else {
                this.updateDashboardWithWarehouseFilters(this.processedData);
            }
        }
    }
    
    filterDataByDate(dateStr) {
        if (!this.processedData || !dateStr) return;
        
        // Format date to match the format in the data (assuming 'SALES Invoice DATE' or 'DELIVERY Note DATE')
        const searchDate = new Date(dateStr);
        const formattedDate = searchDate.toISOString().split('T')[0];
        
        // Filter data for the selected date
        const filteredData = this.processedData.filter(row => {
            const invoiceDate = row['SALES Invoice DATE'] || '';
            const deliveryDate = row['DELIVERY Note DATE'] || '';
            
            return invoiceDate.includes(formattedDate) || deliveryDate.includes(formattedDate);
        });
        
        this.updateDashboardWithWarehouseFilters(filteredData);
    }
    
    filterDataByMonth(monthStr) {
        if (!this.processedData || !monthStr) return;
        
        // Extract year and month from the input (format: YYYY-MM)
        const [year, month] = monthStr.split('-');
        
        // Filter data for the selected month
        const filteredData = this.processedData.filter(row => {
            const invoiceDate = row['SALES Invoice DATE'] || '';
            const deliveryDate = row['DELIVERY Note DATE'] || '';
            
            // Check if either date is in the selected month
            const checkDate = (dateStr) => {
                if (!dateStr) return false;
                // Try to extract year and month from the date string
                const parts = dateStr.split('-');
                return parts[0] === year && parts[1] === month;
            };
            
            return checkDate(invoiceDate) || checkDate(deliveryDate);
        });
        
        this.updateDashboardWithWarehouseFilters(filteredData);
    }
    
    updateDashboard(filteredData) {
        if (!filteredData || !Array.isArray(filteredData)) return;
        
        // Extract warehouse structure and populate filters on first data load
        if (!this.locations || this.locations.length === 0) {
            this.extractWarehouseStructure(this.processedData);
            this.populateWarehouseFilters();
        }
        
        // Sort the data by date
        filteredData = this.sortDataByDate(filteredData);
        
        // Check cache first
        const dataHash = this.hashData(filteredData);
        if (this.lastDataHash === dataHash && this.statsCache.has('dashboard_stats')) {
            const cachedStats = this.statsCache.get('dashboard_stats');
            this.updateAllCardStats(cachedStats);
            return;
        }
        
        // Single-pass categorization and calculation for maximum performance
        const stats = this.calculateStatsOptimized(filteredData);
        
        // Cache results
        this.statsCache.set('dashboard_stats', stats);
        this.lastDataHash = dataHash;
        
        // Update UI
        this.updateAllCardStats(stats);
        
        // Update ORDER VIEW dashboard
        this.updateOrderView(filteredData);
    }
    
    // Optimized single-pass calculation
    calculateStatsOptimized(data) {
        // Log available columns for debugging
        if (data && data.length > 0) {
            console.log("Available columns in data:", Object.keys(data[0]));
            
            // Check for possible LR/AWB column names
            const possibleLRColumns = Object.keys(data[0]).filter(col => 
                col.toLowerCase().includes('awb') || 
                col.toLowerCase().includes('lr') || 
                col.toLowerCase().includes('shipment') ||
                col.toLowerCase().includes('tracking') ||
                col.toLowerCase().includes('docket')
            );
            console.log("Possible LR/AWB columns found:", possibleLRColumns);
            
            // Log first few rows to see LR data
            console.log("Sample LR data from first 3 rows:");
            for (let i = 0; i < Math.min(3, data.length); i++) {
                const row = data[i];
                console.log(`Row ${i + 1}:`);
                possibleLRColumns.forEach(col => {
                    console.log(`  ${col}: "${row[col]}"`);
                });
                console.log(`  SHIPMENT Awb NUMBER: "${row['SHIPMENT Awb NUMBER']}"`);
            }
        }
        
        const stats = {
            total: { invoices: 0, quantity: 0, soQty: 0, soNumbers: 0, qtyDiff: 0, cbm: 0, lrPending: 0 },
            b2c: { invoices: 0, quantity: 0, soQty: 0, soNumbers: 0, qtyDiff: 0, cbm: 0, lrPending: 0 },
            ecom: { invoices: 0, quantity: 0, soQty: 0, soNumbers: 0, qtyDiff: 0, cbm: 0, lrPending: 0 },
            offline: { invoices: 0, quantity: 0, soQty: 0, soNumbers: 0, qtyDiff: 0, cbm: 0, lrPending: 0 },
            quickcom: { invoices: 0, quantity: 0, soQty: 0, soNumbers: 0, qtyDiff: 0, cbm: 0, lrPending: 0 },
            ebo: { invoices: 0, quantity: 0, soQty: 0, soNumbers: 0, qtyDiff: 0, cbm: 0, lrPending: 0 },
            others: { invoices: 0, quantity: 0, soQty: 0, soNumbers: 0, qtyDiff: 0, cbm: 0, lrPending: 0 }
        };
        
        // Track unique invoices for each category
        const uniqueInvoices = {
            total: new Set(),
            b2c: new Set(),
            ecom: new Set(),
            offline: new Set(),
            quickcom: new Set(),
            ebo: new Set(),
            others: new Set()
        };
        
        // Track unique SO numbers for each category
        const uniqueSONumbers = {
            total: new Set(),
            b2c: new Set(),
            ecom: new Set(),
            offline: new Set(),
            quickcom: new Set(),
            ebo: new Set(),
            others: new Set()
        };
        
        // Define the customer groups for each category
        const categoryGroups = {
            b2c: ['decathlon', 'flflipkart(b2c)', 'snapmint', 'shopify', 'tatacliq', 'amazon b2c', 'pepperfry'],
            ecom: ['amazon', 'flipkart'],
            offline: ['offline sales-b2b', 'offline – gt', 'offline - mt'],
            quickcom: ['blinkit', 'swiggy', 'bigbasket', 'zepto'],
            ebo: ['store 2-lucknow', 'store3-zirakpur'],
            others: ['sales to vendor', 'internal company', 'others']
        };
        
        // Single pass through data
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            
            // Determine category based on customer group
            let category = 'others'; // Default category
            
            for (const [cat, groups] of Object.entries(categoryGroups)) {
                if (groups.some(group => customerGroup.includes(group))) {
                    category = cat;
                    break;
                }
            }
            
            // Calculate values once
            const invoiceNo = row['SALES Invoice NO'] || row['DELIVERY Note NO'];
            const soNumber = row['Sales Order No'];
            const quantity = this.getQuantity(row);
            const soQty = this.getSalesOrderQty(row);
            const qtyDiff = this.getQtyDifference(row);
            const cbm = parseFloat(row['SI Total CBM'] || row['DN Total CBM'] || 0);
            
            // Try multiple possible LR/AWB column names
            const lrNo = row['SHIPMENT Awb NUMBER'] || 
                        row['Shipment Awb Number'] || 
                        row['AWB Number'] || 
                        row['LR Number'] || 
                        row['Tracking Number'] || 
                        row['Docket Number'] || 
                        '';
            const hasLR = lrNo && lrNo.toString().trim() !== '';
            
            // Debug LR information for first few rows
            if (i < 3) {
                console.log(`Row ${i + 1} LR info: "${lrNo}" (hasLR: ${hasLR})`);
            }
            
            // Update stats for total and specific category
            [stats.total, stats[category]].forEach((statObj, index) => {
                const invoiceSet = index === 0 ? uniqueInvoices.total : uniqueInvoices[category];
                const soNumberSet = index === 0 ? uniqueSONumbers.total : uniqueSONumbers[category];
                
                if (invoiceNo && invoiceNo.toString().trim() !== '') {
                    invoiceSet.add(invoiceNo.toString().trim());
                }
                
                if (soNumber && soNumber.toString().trim() !== '') {
                    soNumberSet.add(soNumber.toString().trim());
                }
                
                statObj.quantity += quantity;
                statObj.soQty += soQty;
                statObj.qtyDiff += qtyDiff;
                if (!isNaN(cbm) && cbm >= 0) {
                    statObj.cbm += cbm;
                }
                if (!hasLR) {
                    statObj.lrPending++;
                }
            });
        }
        
        // Set invoice counts from unique sets
        for (const category in uniqueInvoices) {
            stats[category].invoices = uniqueInvoices[category].size;
        }
        
        // Set SO Numbers counts from unique sets
        for (const category in uniqueSONumbers) {
            stats[category].soNumbers = uniqueSONumbers[category].size;
        }
        
        // Format CBM values
        Object.values(stats).forEach(stat => {
            stat.cbm = stat.cbm.toFixed(2);
        });
        
        return stats;
    }
    
    // Initialize ORDER VIEW Dashboard
    initializeOrderView() {
        const orderViewDiv = document.getElementById('orderViewSection');
        if (!orderViewDiv) {
            console.error("Order View section not found");
            return;
        }
        
        orderViewDiv.innerHTML = `
            <div class="dashboard-grid">
                <!-- Total Card -->
                <div class="dashboard-card total-card">
                    <div class="card-header">
                        <h3>📊 Total</h3>
                        <div class="card-icon">🎯</div>
                    </div>
                    <div class="card-body">
                        <div class="transport-section">
                            <div class="transport-type ftl">
                                <h4>🚛 FTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="totalFtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="totalFtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="totalFtlManpower">0</div>
                                </div>
                            </div>
                            <div class="transport-type ptl">
                                <h4>📦 PTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="totalPtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="totalPtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="totalPtlManpower">0</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- B2C Card -->
                <div class="dashboard-card b2c-card">
                    <div class="card-header">
                        <h3>🛍️ B2C</h3>
                        <div class="card-icon">🛒</div>
                    </div>
                    <div class="card-body">
                        <div class="transport-section">
                            <div class="transport-type ftl">
                                <h4>🚛 FTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="b2cFtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="b2cFtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="b2cFtlManpower">0</div>
                                </div>
                            </div>
                            <div class="transport-type ptl">
                                <h4>📦 PTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="b2cPtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="b2cPtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="b2cPtlManpower">0</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- E-commerce Card -->
                <div class="dashboard-card ecom-card">
                    <div class="card-header">
                        <h3>🛒 E-commerce</h3>
                        <div class="card-icon">📱</div>
                    </div>
                    <div class="card-body">
                        <div class="transport-section">
                            <div class="transport-type ftl">
                                <h4>🚛 FTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="ecomFtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="ecomFtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="ecomFtlManpower">0</div>
                                </div>
                            </div>
                            <div class="transport-type ptl">
                                <h4>📦 PTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="ecomPtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="ecomPtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="ecomPtlManpower">0</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Offline Card -->
                <div class="dashboard-card offline-card">
                    <div class="card-header">
                        <h3>🏪 Offline</h3>
                        <div class="card-icon">🏬</div>
                    </div>
                    <div class="card-body">
                        <div class="transport-section">
                            <div class="transport-type ftl">
                                <h4>🚛 FTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="offlineFtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="offlineFtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="offlineFtlManpower">0</div>
                                </div>
                            </div>
                            <div class="transport-type ptl">
                                <h4>📦 PTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="offlinePtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="offlinePtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="offlinePtlManpower">0</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Quick Commerce Card -->
                <div class="dashboard-card quickcom-card">
                    <div class="card-header">
                        <h3>⚡ Quick Commerce</h3>
                        <div class="card-icon">🏃‍♂️</div>
                    </div>
                    <div class="card-body">
                        <div class="transport-section">
                            <div class="transport-type ftl">
                                <h4>🚛 FTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="quickcomFtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="quickcomFtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="quickcomFtlManpower">0</div>
                                </div>
                            </div>
                            <div class="transport-type ptl">
                                <h4>📦 PTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="quickcomPtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="quickcomPtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="quickcomPtlManpower">0</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- EBO Card -->
                <div class="dashboard-card ebo-card">
                    <div class="card-header">
                        <h3>🏢 EBO</h3>
                        <div class="card-icon">🏬</div>
                    </div>
                    <div class="card-body">
                        <div class="transport-section">
                            <div class="transport-type ftl">
                                <h4>🚛 FTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="eboFtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="eboFtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="eboFtlManpower">0</div>
                                </div>
                            </div>
                            <div class="transport-type ptl">
                                <h4>📦 PTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="eboPtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="eboPtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="eboPtlManpower">0</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Others Card -->
                <div class="dashboard-card others-card">
                    <div class="card-header">
                        <h3>📋 Others</h3>
                        <div class="card-icon">📝</div>
                    </div>
                    <div class="card-body">
                        <div class="transport-section">
                            <div class="transport-type ftl">
                                <h4>🚛 FTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="othersFtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="othersFtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="othersFtlManpower">0</div>
                                </div>
                            </div>
                            <div class="transport-type ptl">
                                <h4>📦 PTL</h4>
                                <div class="metric">
                                    <div class="metric-label">SO Numbers</div>
                                    <div class="metric-value" id="othersPtlSONumbers">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">SO Qty</div>
                                    <div class="metric-value" id="othersPtlSOQty">0</div>
                                </div>
                                <div class="metric">
                                    <div class="metric-label">Manpower</div>
                                    <div class="metric-value" id="othersPtlManpower">0</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }
    
    // Calculate ORDER VIEW statistics
    calculateOrderViewStats(data) {
        const orderStats = {
            total: { ftl: { soNumbers: 0, soQty: 0, manpower: 0 }, ptl: { soNumbers: 0, soQty: 0, manpower: 0 } },
            b2c: { ftl: { soNumbers: 0, soQty: 0, manpower: 0 }, ptl: { soNumbers: 0, soQty: 0, manpower: 0 } },
            ecom: { ftl: { soNumbers: 0, soQty: 0, manpower: 0 }, ptl: { soNumbers: 0, soQty: 0, manpower: 0 } },
            offline: { ftl: { soNumbers: 0, soQty: 0, manpower: 0 }, ptl: { soNumbers: 0, soQty: 0, manpower: 0 } },
            quickcom: { ftl: { soNumbers: 0, soQty: 0, manpower: 0 }, ptl: { soNumbers: 0, soQty: 0, manpower: 0 } },
            ebo: { ftl: { soNumbers: 0, soQty: 0, manpower: 0 }, ptl: { soNumbers: 0, soQty: 0, manpower: 0 } },
            others: { ftl: { soNumbers: 0, soQty: 0, manpower: 0 }, ptl: { soNumbers: 0, soQty: 0, manpower: 0 } }
        };
        
        // Track unique SO numbers for each category and transport type
        const uniqueSONumbers = {
            total: { ftl: new Set(), ptl: new Set() },
            b2c: { ftl: new Set(), ptl: new Set() },
            ecom: { ftl: new Set(), ptl: new Set() },
            offline: { ftl: new Set(), ptl: new Set() },
            quickcom: { ftl: new Set(), ptl: new Set() },
            ebo: { ftl: new Set(), ptl: new Set() },
            others: { ftl: new Set(), ptl: new Set() }
        };
        
        // Define the customer groups for each category
        const categoryGroups = {
            b2c: ['decathlon', 'flflipkart(b2c)', 'snapmint', 'shopify', 'tatacliq', 'amazon b2c', 'pepperfry'],
            ecom: ['amazon', 'flipkart'],
            offline: ['offline sales-b2b', 'offline – gt', 'offline - mt'],
            quickcom: ['blinkit', 'swiggy', 'bigbasket', 'zepto'],
            ebo: ['store 2-lucknow', 'store3-zirakpur'],
            others: ['sales to vendor', 'internal company', 'others']
        };
        
        // Single pass through data
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const customerGroup = (row['Customer Group'] || '').toLowerCase();
            
            // Determine category based on customer group
            let category = 'others'; // Default category
            
            for (const [cat, groups] of Object.entries(categoryGroups)) {
                if (groups.some(group => customerGroup.includes(group))) {
                    category = cat;
                    break;
                }
            }
            
            // Get transport type
            const transportType = this.getTransportType(row).toLowerCase();
            if (transportType !== 'ftl' && transportType !== 'ptl') {
                continue; // Skip unknown transport types
            }
            
            // Calculate values
            const soNumber = row['Sales Order No'];
            const soQty = this.getSalesOrderQty(row);
            
            // Update stats for total and specific category
            [orderStats.total, orderStats[category]].forEach((statObj, index) => {
                const soNumberSet = index === 0 ? uniqueSONumbers.total[transportType] : uniqueSONumbers[category][transportType];
                
                if (soNumber && soNumber.toString().trim() !== '') {
                    soNumberSet.add(soNumber.toString().trim());
                }
                
                statObj[transportType].soQty += soQty;
            });
        }

        // Derive manpower after accumulating quantities
        const computeManpower = (qty) => Math.floor((qty || 0) / 1000);
        Object.values(orderStats).forEach(catObj => {
            ['ftl','ptl'].forEach(t => {
                catObj[t].manpower = computeManpower(catObj[t].soQty);
            });
        });
        
        // Set SO Numbers counts from unique sets
        for (const category in uniqueSONumbers) {
            for (const transportType of ['ftl', 'ptl']) {
                orderStats[category][transportType].soNumbers = uniqueSONumbers[category][transportType].size;
            }
        }
        
        return orderStats;
    }
    
    // Update ORDER VIEW dashboard
    updateOrderView(data) {
        if (!data || !Array.isArray(data)) return;
        
        const orderStats = this.calculateOrderViewStats(data);
        this.updateOrderViewCards(orderStats);
    }
    
    // Update ORDER VIEW card displays
    updateOrderViewCards(orderStats) {
        const categories = ['total', 'b2c', 'ecom', 'offline', 'quickcom', 'ebo', 'others'];
        const transportTypes = ['ftl', 'ptl'];
        
        // Helper function to animate value update
        const animateValueUpdate = (element, newValue) => {
            if (!element) return;
            element.textContent = newValue;
            element.style.animation = 'valueChange 0.7s ease-in-out';
            element.addEventListener('animationend', () => {
                element.style.animation = '';
            }, { once: true });
        };
        
        categories.forEach(category => {
            transportTypes.forEach(transportType => {
                const soNumbersElement = document.getElementById(`${category}${transportType.charAt(0).toUpperCase() + transportType.slice(1)}SONumbers`);
                const soQtyElement = document.getElementById(`${category}${transportType.charAt(0).toUpperCase() + transportType.slice(1)}SOQty`);
                const manpowerElement = document.getElementById(`${category}${transportType.charAt(0).toUpperCase() + transportType.slice(1)}Manpower`);
                
                if (soNumbersElement && orderStats[category] && orderStats[category][transportType]) {
                    animateValueUpdate(soNumbersElement, orderStats[category][transportType].soNumbers.toLocaleString());
                }
                
                if (soQtyElement && orderStats[category] && orderStats[category][transportType]) {
                    animateValueUpdate(soQtyElement, orderStats[category][transportType].soQty.toLocaleString());
                }

                if (manpowerElement && orderStats[category] && orderStats[category][transportType]) {
                    animateValueUpdate(manpowerElement, orderStats[category][transportType].manpower.toLocaleString());
                }
            });
        });
    }
    
    
    
    // Simple hash function for data change detection
    hashData(data) {
        if (!Array.isArray(data) || data.length === 0) return 'empty';
        
        // Create a simple hash based on data length and first/last rows
        const firstRow = data[0];
        const lastRow = data[data.length - 1];
        const hash = `${data.length}_${JSON.stringify(firstRow)}_${JSON.stringify(lastRow)}`;
        
        // Simple string hash
        let hashValue = 0;
        for (let i = 0; i < hash.length; i++) {
            const char = hash.charCodeAt(i);
            hashValue = ((hashValue << 5) - hashValue) + char;
            hashValue = hashValue & hashValue; // Convert to 32-bit integer
        }
        
        return hashValue.toString();
    }
    
    // Helper method to update all card stats
    updateAllCardStats(stats) {
        this.updateCardStats('total', stats.total);
        this.updateCardStats('b2c', stats.b2c);
        this.updateCardStats('ecom', stats.ecom);
        this.updateCardStats('offline', stats.offline);
        this.updateCardStats('quickcom', stats.quickcom);
        this.updateCardStats('ebo', stats.ebo);
        this.updateCardStats('others', stats.others);
    }

    // Extract warehouse code from warehouse name
    getWarehouseCode(warehouse) {
        if (!warehouse) return '';
        
        // Look for common patterns like MH4, HR3, etc.
        const codeMatch = warehouse.match(/\b([A-Z]{2}\d+)/);
        if (codeMatch && codeMatch[1]) {
            return codeMatch[1];
        }
        
        return warehouse; // Return original if no code found
    }
    
    // Get location for a warehouse code
    getWarehouseLocation(warehouseCode) {
        // Extract the base warehouse code (without suffixes)
        const baseCode = warehouseCode.match(/^([A-Z]{2}\d+)/);
        const code = baseCode && baseCode[1] ? baseCode[1] : warehouseCode;
        
        return this.warehouseLocationMapping[code] || 'Other';
    }
    
    // Sort data by warehouse and then by date
    sortDataByWarehouseAndDate(data) {
        if (!Array.isArray(data) || data.length === 0) {
            return data;
        }
        
        // Get warehouse sort option
        const warehouseSortSelect = document.getElementById('warehouseSort');
        const warehouseSort = warehouseSortSelect ? warehouseSortSelect.value : 'none';
        
        // If no warehouse sorting, just sort by date
        if (warehouseSort === 'none') {
            return this.sortDataByDate(data);
        }
        
        // Make a copy of the data to avoid modifying original
        const sortedData = [...data];
        
        // First sort by date (this will be secondary sort)
        this.sortDataByDate(sortedData);
        
        // Then sort by warehouse information
        return sortedData.sort((a, b) => {
            const warehouseA = a['Set Source Warehouse'] || '';
            const warehouseB = b['Set Source Warehouse'] || '';
            
            if (warehouseSort === 'location') {
                // Two-level sort: first by location, then by warehouse code
                const codeA = this.getWarehouseCode(warehouseA);
                const codeB = this.getWarehouseCode(warehouseB);
                
                const locationA = this.getWarehouseLocation(codeA);
                const locationB = this.getWarehouseLocation(codeB);
                
                // First compare locations
                if (locationA !== locationB) {
                    return locationA.localeCompare(locationB);
                }
                
                // Then compare warehouse codes
                return codeA.localeCompare(codeB);
            } else {
                // Just sort by warehouse code
                const codeA = this.getWarehouseCode(warehouseA);
                const codeB = this.getWarehouseCode(warehouseB);
                
                return codeA.localeCompare(codeB);
            }
        });
    }

    // Sort data by SO Date only
    sortDataByDate(data) {
        if (!Array.isArray(data) || data.length === 0) {
            return data;
        }

        // Get the selected sort order from the main page
        const sortOrderSelect = document.getElementById('sortOrder');
        const sortOrder = sortOrderSelect ? sortOrderSelect.value : 'newest';

        return data.sort((a, b) => {
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
        });
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
                console.warn('Invalid date format in dashboard:', dateStr);
                return new Date(0); // Return epoch for invalid dates
            }
            
            return date;
        } catch (error) {
            console.error('Error parsing date in dashboard:', dateStr, error);
            return new Date(0); // Return epoch for error cases
        }
    }
    
    calculateStats(data) {
        if (!Array.isArray(data) || data.length === 0) {
            return {
                invoices: 0,
                quantity: 0,
                cbm: '0.00',
                lrPending: 0
            };
        }

        try {
            // Count unique invoices (using SALES Invoice NO or DELIVERY Note NO)
            const uniqueInvoices = new Set();
            data.forEach(row => {
                if (row && typeof row === 'object') {
                    const invoiceNo = row['SALES Invoice NO'] || row['DELIVERY Note NO'];
                    if (invoiceNo && invoiceNo.toString().trim() !== '') {
                        uniqueInvoices.add(invoiceNo.toString().trim());
                    }
                }
            });
            
            // Calculate total quantity using the same logic as script.js (with fallback)
            let totalQuantity = 0;
            data.forEach(row => {
                if (row && typeof row === 'object') {
                    const qty = this.getQuantity(row);
                    totalQuantity += qty;
                }
            });
            
            // Calculate total CBM (ensuring non-negative values)
            let totalCBM = 0;
            data.forEach(row => {
                if (row && typeof row === 'object') {
                    const cbm = parseFloat(row['SI Total CBM'] || row['DN Total CBM'] || 0);
                    if (!isNaN(cbm) && cbm >= 0) {
                        totalCBM += cbm;
                    }
                }
            });
            
            // Calculate LR Pending - records without LR Number
            let lrPending = 0;
            data.forEach(row => {
                if (row && typeof row === 'object') {
                    const lrNo = row['SHIPMENT Awb NUMBER'] || '';
                    // Count as pending if LR Number is empty, null, or just whitespace
                    if (!lrNo || lrNo.toString().trim() === '') {
                        lrPending++;
                    }
                }
            });
            
            return {
                invoices: uniqueInvoices.size,
                quantity: Math.round(totalQuantity),
                cbm: totalCBM.toFixed(2),
                lrPending: lrPending
            };
        } catch (error) {
            console.error('Error calculating stats:', error);
            return {
                invoices: 0,
                quantity: 0,
                cbm: '0.00',
                lrPending: 0
            };
        }
    }
    
    // Get the quantity as a proper number (ensuring non-negative values)
    // This method matches the logic from script.js to ensure consistency
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
    
    // Get the Sales Order quantity
    getSalesOrderQty(row) {
        if (!row || typeof row !== 'object') return 0;
        
        // Check cache first
        if (row._cachedSOQuantity !== undefined) {
            return row._cachedSOQuantity;
        }
        
        try {
            // Try different possible field names for Sales Order Quantity
            let qty = row['SALES Order QTY'] || 
                      row['SO QTY'] || 
                      row['Sales Order QTY'] || 
                      row['Sales Order Qty'] || 
                      row['SO Quantity'] || 
                      row['SALES ORDER QUANTITY'] || 
                      row['ORDER QTY'] ||
                      row['Order Qty'] || 0;
            
            if (qty === null || qty === undefined) {
                // If none of the above fields exist, check for SO in field names
                for (const key in row) {
                    if (key.includes('SO') && key.includes('QTY') || 
                        key.includes('Sales Order') && key.includes('Qty') ||
                        key.includes('Order') && key.includes('Qty')) {
                        qty = row[key];
                        console.log('Found Sales Order Qty field:', key);
                        break;
                    }
                }
                
                // If still not found, return 0
                if (qty === null || qty === undefined) {
                    row._cachedSOQuantity = 0;
                    return 0;
                }
            }
            
            // Fast numeric conversion
            const numQty = +qty;
            
            // Cache the result
            row._cachedSOQuantity = isNaN(numQty) ? 0 : Math.max(0, numQty);
            return row._cachedSOQuantity;
            
        } catch (error) {
            console.error('Error getting Sales Order Qty:', error);
            row._cachedSOQuantity = 0;
            return 0;
        }
    }
    
    // Calculate the difference between Sales Order QTY and Sales Invoice QTY
    getQtyDifference(row) {
        const soQty = this.getSalesOrderQty(row);
        const siQty = this.getQuantity(row);
        return Math.max(0, soQty - siQty); // Ensure non-negative result
    }
    
    // Determine transport type (FTL/PTL) based on transporter
    getTransportType(row) {
        const transporter = (row['Transporter'] || '').toString().trim();
        
        if (!transporter) return 'Unknown';
        
        // Check FTL transporters
        for (const ftlTransporter of this.transporterMapping.FTL) {
            if (transporter.toLowerCase().includes(ftlTransporter.toLowerCase())) {
                return 'FTL';
            }
        }
        
        // Check PTL transporters  
        for (const ptlTransporter of this.transporterMapping.PTL) {
            if (transporter.toLowerCase().includes(ptlTransporter.toLowerCase())) {
                return 'PTL';
            }
        }
        
        // If not found in either list, return Unknown
        return 'Unknown';
    }
    
    updateCardStats(cardType, stats) {
        const invoicesElement = document.getElementById(`${cardType}Invoices`);
        const quantityElement = document.getElementById(`${cardType}Quantity`);
        const soQtyElement = document.getElementById(`${cardType}SOQty`);
        const soNumbersElement = document.getElementById(`${cardType}SONumbers`);
        const qtyDiffElement = document.getElementById(`${cardType}QtyDiff`);
        const cbmElement = document.getElementById(`${cardType}CBM`);
        const lrPendingElement = document.getElementById(`${cardType}LRPending`);
        
        // Helper function to animate value update
        const animateValueUpdate = (element, newValue) => {
            if (!element) return;
            element.textContent = newValue;
            element.style.animation = 'valueChange 0.7s ease-in-out';
            element.addEventListener('animationend', () => {
                element.style.animation = '';
            }, { once: true });
        };
        
        if (invoicesElement) {
            animateValueUpdate(invoicesElement, stats.invoices.toLocaleString());
        }
        if (soQtyElement) {
            animateValueUpdate(soQtyElement, stats.soQty ? stats.soQty.toLocaleString() : '0');
        }
        if (soNumbersElement) {
            animateValueUpdate(soNumbersElement, stats.soNumbers ? stats.soNumbers.toLocaleString() : '0');
        }
        if (quantityElement) {
            animateValueUpdate(quantityElement, stats.quantity.toLocaleString());
        }
        if (qtyDiffElement) {
            animateValueUpdate(qtyDiffElement, stats.qtyDiff.toLocaleString());
        }
        if (cbmElement) {
            animateValueUpdate(cbmElement, parseFloat(stats.cbm).toLocaleString(undefined, {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }));
        }
        if (lrPendingElement) {
            animateValueUpdate(lrPendingElement, stats.lrPending.toLocaleString());
        }
    }
    
    // Extract warehouse structure information from data
    extractWarehouseStructure(data) {
        if (!Array.isArray(data) || data.length === 0) {
            console.warn("No data or invalid data provided to extractWarehouseStructure");
            return;
        }
        
        console.log(`Extracting warehouse structure from ${data.length} rows`);
        
        const locationSet = new Set();
        const areaCodeMap = {};
        const warehouseMap = {};
        const warehouseValues = new Set();
        
        // Find the warehouse column name - try multiple possible naming conventions
        const possibleWarehouseColumns = [
            'Set Source Warehouse',
            'Source Warehouse',
            'Warehouse',
            'Set Warehouse',
            'SourceWarehouse',
            'Source_Warehouse',
            'Set_Source_Warehouse',
            'Warehouse Code'
        ];
        
        // Determine which column name exists in the data
        let warehouseColumnName = '';
        if (data.length > 0) {
            const sampleRow = data[0];
            for (const colName of possibleWarehouseColumns) {
                if (colName in sampleRow) {
                    warehouseColumnName = colName;
                    console.log(`Found warehouse column: "${warehouseColumnName}"`);
                    break;
                }
            }
        }
        
        if (!warehouseColumnName) {
            console.warn("No warehouse column found in the data. Available columns:", 
                data.length > 0 ? Object.keys(data[0]) : "No data rows");
            
            // If no specific warehouse column found, try to extract from all columns
            if (data.length > 0) {
                // Check all columns in the first few rows for warehouse patterns
                const sampleRow = data[0];
                console.log("Searching all columns for warehouse codes...");
                Object.keys(sampleRow).forEach(column => {
                    console.log(`Column '${column}': ${sampleRow[column]}`);
                });
            }
        }
        
        // First pass: extract unique locations, area codes, and warehouses
        data.forEach((row, index) => {
            if (!row || typeof row !== 'object') return;
            
            // Debug: Check first 5 rows to understand data structure
            if (index < 5) {
                console.log(`Row ${index} keys:`, Object.keys(row));
            }
            
            // Get warehouse value from the identified column or try Set Source Warehouse as default
            const warehouse = warehouseColumnName ? row[warehouseColumnName] || '' : row['Set Source Warehouse'] || '';
            
            // Track all warehouse values for debugging
            if (warehouse) {
                warehouseValues.add(warehouse);
                if (index < 10) {
                    console.log(`Row ${index} warehouse value:`, warehouse);
                }
            }
            
            if (!warehouse) return;
            
            // Try to extract warehouse code using our different patterns
            let areaCode = '';
            let location = '';
            
            // Extract area code (e.g., MH4, HR3) from warehouse name
            const areaCodeMatch = warehouse.match(/\b([A-Z]{2}\d+)/i);
            
            if (areaCodeMatch) {
                areaCode = areaCodeMatch[1].toUpperCase();
                location = this.warehouseLocationMapping[areaCode] || 'Other';
                
                // If still 'Other', try to determine location from code pattern
                if (location === 'Other') {
                    if (/^MH/i.test(areaCode)) location = 'Mumbai';
                    else if (/^HR/i.test(areaCode)) location = 'Gurgaon';
                    else if (/^WB/i.test(areaCode)) location = 'Howrah';
                    else if (/^KA/i.test(areaCode)) location = 'Bangalore';
                    else if (/^PB/i.test(areaCode)) location = 'Ludhiana';
                }
                
                console.log(`Extracted: Warehouse=${warehouse}, Code=${areaCode}, Location=${location}`);
            } else {
                console.log(`No area code found in warehouse: "${warehouse}"`);
                
                // Try to extract location directly from warehouse name
                const lowerWarehouse = warehouse.toLowerCase();
                if (lowerWarehouse.includes('mumbai')) {
                    location = 'Mumbai';
                    areaCode = 'MH';
                } else if (lowerWarehouse.includes('gurgaon')) {
                    location = 'Gurgaon';
                    areaCode = 'HR';
                } else if (lowerWarehouse.includes('howrah')) {
                    location = 'Howrah';
                    areaCode = 'WB';
                } else if (lowerWarehouse.includes('bangalore')) {
                    location = 'Bangalore';
                    areaCode = 'KA';
                } else if (lowerWarehouse.includes('ludhiana')) {
                    location = 'Ludhiana';
                    areaCode = 'PB';
                } else {
                    location = 'Other';
                    areaCode = 'XX';
                }
            }
            
            // Add location
            locationSet.add(location);
            
            // Add area code to location
            if (!areaCodeMap[location]) {
                areaCodeMap[location] = new Set();
            }
            areaCodeMap[location].add(areaCode);
            
            // Add warehouse to area code
            const key = `${location}:${areaCode}`;
            if (!warehouseMap[key]) {
                warehouseMap[key] = new Set();
            }
            warehouseMap[key].add(warehouse);
        });
        
        // If we didn't extract any data, add sample data for testing
        if (locationSet.size === 0) {
            console.warn("No warehouse data extracted from the Excel file. Adding sample data for testing.");
            
            // Add sample data for each location
            ['Mumbai', 'Gurgaon', 'Bangalore', 'Howrah', 'Ludhiana'].forEach(location => {
                locationSet.add(location);
                areaCodeMap[location] = new Set();
                
                // Add sample area codes for each location
                const areaCodes = {
                    'Mumbai': ['MH4', 'MH5'],
                    'Gurgaon': ['HR3', 'HR10', 'HR11'],
                    'Bangalore': ['KA3', 'KA4'],
                    'Howrah': ['WB4'],
                    'Ludhiana': ['PB2']
                };
                
                (areaCodes[location] || []).forEach(code => {
                    areaCodeMap[location].add(code);
                    const key = `${location}:${code}`;
                    warehouseMap[key] = new Set([`${code} - LORPL`]);
                });
            });
        }
        
        // Convert sets to sorted arrays
        this.locations = Array.from(locationSet).sort();
        
        this.areaCodes = {};
        Object.keys(areaCodeMap).forEach(location => {
            this.areaCodes[location] = Array.from(areaCodeMap[location]).sort();
        });
        
        this.warehouses = {};
        Object.keys(warehouseMap).forEach(key => {
            this.warehouses[key] = Array.from(warehouseMap[key])
                .filter(Boolean)
                .sort();
        });
        
        console.log('🏭 Extracted Warehouse Structure:');
        console.log('Locations:', this.locations);
        console.log('Area Codes by Location:', this.areaCodes);
        console.log('Warehouses by Location/Area:', this.warehouses);
        
        // Log the total number of unique warehouse values found
        console.log(`Total unique warehouse values found: ${warehouseValues.size}`);
        if (warehouseValues.size > 0) {
            console.log('Sample warehouse values:', Array.from(warehouseValues).slice(0, 10));
        }
    }
    
    // Populate dropdown options
    populateWarehouseFilters() {
        console.log("Populating warehouse filters with:", {
            locations: this.locations,
            areaCodes: this.areaCodes,
            warehouseCount: Object.keys(this.warehouses).length
        });
        
        // Populate location dropdown
        const locationFilter = document.getElementById('locationFilter');
        if (locationFilter) {
            // Clear existing options except the first one
            while (locationFilter.options.length > 1) {
                locationFilter.remove(1);
            }
            
            // Add location options
            this.locations.forEach(location => {
                const option = document.createElement('option');
                option.value = location;
                option.textContent = location;
                locationFilter.appendChild(option);
                console.log(`Added location option: ${location}`);
            });
        } else {
            console.warn("Location filter dropdown not found in the DOM");
        }
        
        // Update dependent dropdowns
        this.updateAreaCodeFilter();
    }
    
    // Update area code dropdown based on selected location
    updateAreaCodeFilter() {
        const locationFilter = document.getElementById('locationFilter');
        const areaCodeFilter = document.getElementById('areaCodeFilter');
        
        if (!locationFilter || !areaCodeFilter) return;
        
        const selectedLocation = locationFilter.value;
        
        // Clear existing options except the first one
        while (areaCodeFilter.options.length > 1) {
            areaCodeFilter.remove(1);
        }
        
        // If "All Locations" is selected or no valid location, stop here
        if (selectedLocation === 'all') {
            areaCodeFilter.disabled = true;
            this.updateWarehouseFilter();
            return;
        }
        
        // Enable the dropdown
        areaCodeFilter.disabled = false;
        
        // Add area code options for the selected location
        const areaCodes = this.areaCodes[selectedLocation] || [];
        areaCodes.forEach(areaCode => {
            const option = document.createElement('option');
            option.value = areaCode;
            option.textContent = areaCode;
            areaCodeFilter.appendChild(option);
        });
        
        // Update warehouse dropdown
        this.updateWarehouseFilter();
    }
    
    // Update warehouse dropdown based on selected location and area code
    updateWarehouseFilter() {
        const locationFilter = document.getElementById('locationFilter');
        const areaCodeFilter = document.getElementById('areaCodeFilter');
        const warehouseFilter = document.getElementById('warehouseFilter');
        
        if (!locationFilter || !areaCodeFilter || !warehouseFilter) return;
        
        const selectedLocation = locationFilter.value;
        const selectedAreaCode = areaCodeFilter.value;
        
        // Clear existing options except the first one
        while (warehouseFilter.options.length > 1) {
            warehouseFilter.remove(1);
        }
        
        // If "All Locations" or "All Area Codes" is selected, disable the warehouse dropdown
        if (selectedLocation === 'all' || selectedAreaCode === 'all') {
            warehouseFilter.disabled = true;
            return;
        }
        
        // Enable the dropdown
        warehouseFilter.disabled = false;
        
        // Add warehouse options for the selected location and area code
        const key = `${selectedLocation}:${selectedAreaCode}`;
        const warehouses = this.warehouses[key] || [];
        warehouses.forEach(warehouse => {
            const option = document.createElement('option');
            option.value = warehouse;
            option.textContent = warehouse;
            warehouseFilter.appendChild(option);
        });
    }
    
    // Group data by warehouse location and code for display
    groupWarehouseData(data) {
        if (!Array.isArray(data) || data.length === 0) {
            return {};
        }
        
        const warehouseGroups = {};
        
        // First pass: group data by location and warehouse code
        data.forEach(row => {
            if (!row || typeof row !== 'object') return;
            
            const warehouse = row['Set Source Warehouse'] || '';
            const warehouseCode = this.getWarehouseCode(warehouse);
            const location = this.getWarehouseLocation(warehouseCode);
            
            // Initialize location group if not exists
            if (!warehouseGroups[location]) {
                warehouseGroups[location] = {};
            }
            
            // Get the base warehouse code (without suffixes)
            const baseCode = warehouseCode.match(/^([A-Z]{2}\d+)/);
            const code = baseCode && baseCode[1] ? baseCode[1] : warehouseCode;
            
            // Initialize warehouse code group if not exists
            if (!warehouseGroups[location][code]) {
                warehouseGroups[location][code] = {
                    warehouses: new Set(),
                    rows: []
                };
            }
            
            // Add the full warehouse name to the set
            warehouseGroups[location][code].warehouses.add(warehouse);
            
            // Add the row to this warehouse group
            warehouseGroups[location][code].rows.push(row);
        });
        
        // Convert Sets to arrays for easier display
        Object.keys(warehouseGroups).forEach(location => {
            Object.keys(warehouseGroups[location]).forEach(code => {
                warehouseGroups[location][code].warehouses = Array.from(
                    warehouseGroups[location][code].warehouses
                ).filter(Boolean).sort();
            });
        });
        
        return warehouseGroups;
    }
    
    // Log warehouse grouping information - useful for debugging
    logWarehouseStructure(data) {
        const groupedData = this.groupWarehouseData(data);
        
        console.log('🏭 Warehouse Structure:');
        Object.keys(groupedData).sort().forEach(location => {
            console.log(`Location: ${location}`);
            
            Object.keys(groupedData[location]).sort().forEach(code => {
                console.log(`  Warehouse Code: ${code}`);
                console.log(`    Warehouses: ${groupedData[location][code].warehouses.join(', ')}`);
                console.log(`    Row count: ${groupedData[location][code].rows.length}`);
            });
        });
    }

    // LR Missing Dashboard Functions
    generateLRMissingReport() {
        const datePicker = document.getElementById('lrDatePicker');
        if (!datePicker.value) {
            alert('Please select a date to generate the LR Missing report.');
            return;
        }

        const selectedDate = datePicker.value;
        
        if (!this.processedData || this.processedData.length === 0) {
            alert('No data available. Please upload an Excel file first.');
            return;
        }
        // Start from warehouse-filtered subset
        const warehouseSubset = this.getWarehouseFilteredData();
        // Filter data for selected date and missing LR only within current warehouse scope
        const filteredData = this.filterLRMissingData(selectedDate, warehouseSubset);
        
        // Show loading state
        const generateBtn = document.getElementById('generateLRReport');
        const originalContent = generateBtn.innerHTML;
        generateBtn.innerHTML = '<span class="btn-icon">⏳</span><span class="btn-text">Generating...</span>';
        generateBtn.disabled = true;

        // Simulate processing time for better UX
        setTimeout(() => {
            this.displayLRMissingResults(filteredData, selectedDate);
            
            // Restore button
            generateBtn.innerHTML = originalContent;
            generateBtn.disabled = false;
        }, 800);
    }

    filterLRMissingData(selectedDate, sourceData = null) {
        const targetDate = new Date(selectedDate);
        const dateString = targetDate.toISOString().split('T')[0];
        const data = sourceData || this.processedData || [];
        return data.filter(row => {
            if (!row || typeof row !== 'object') return false;

            // Check if date matches
            const invoiceDate = row['SALES Invoice DATE'];
            const deliveryDate = row['DELIVERY Note DATE'];
            
            let rowDate = null;
            if (invoiceDate) {
                const parsedDate = new Date(invoiceDate);
                if (!isNaN(parsedDate.getTime())) {
                    rowDate = parsedDate.toISOString().split('T')[0];
                }
            } else if (deliveryDate) {
                const parsedDate = new Date(deliveryDate);
                if (!isNaN(parsedDate.getTime())) {
                    rowDate = parsedDate.toISOString().split('T')[0];
                }
            }

            if (rowDate !== dateString) return false;

            // Check if LR is missing (using same logic as LR pending calculation)
            const lrNo = row['SHIPMENT Awb NUMBER'] || 
                        row['Shipment Awb Number'] || 
                        row['AWB Number'] || 
                        row['LR Number'] || 
                        row['Tracking Number'] || 
                        row['Docket Number'] || 
                        '';
            
            return !lrNo || lrNo.toString().trim() === '';
        });
    }

    displayLRMissingResults(data, selectedDate) {
        // Update selected date info
        const selectedDateInfo = document.getElementById('selectedDateInfo');
        const selectedDateText = document.getElementById('selectedDateText');
        const selectedDateCount = document.getElementById('selectedDateCount');
        
        selectedDateText.textContent = `Selected Date: ${new Date(selectedDate).toLocaleDateString()}`;
        selectedDateCount.textContent = `${data.length} LR Missing records`;
        selectedDateInfo.style.display = 'flex';

        // Show summary section
        const summarySection = document.getElementById('lrMissingSummarySection');
        summarySection.style.display = 'block';

        // Calculate and display stats
        this.calculateLRMissingStats(data);
        
        // Display the records table
        this.displayLRMissingTable(data);
    }

    calculateLRMissingStats(data) {
        const stats = {
            total: data.length,
            b2c: 0,
            ecom: 0,
            offline: 0,
            quickcom: 0,
            ebo: 0,
            others: 0
        };

        // Count by category
        data.forEach(row => {
            const cat = this.getCategoryForRow(row);
            stats[cat]++;
        });

        // Update UI elements matching HTML IDs
        this.updateElement('totalLRMissingCount', stats.total);
        this.updateElement('b2cLRMissingCount', stats.b2c);
        this.updateElement('ecomLRMissingCount', stats.ecom);
        this.updateElement('offlineLRMissingCount', stats.offline);
        this.updateElement('quickcomLRMissingCount', stats.quickcom);
        this.updateElement('eboLRMissingCount', stats.ebo);
        this.updateElement('othersLRMissingCount', stats.others);

        // Also update filter option counts if present inside select (safe no-op if span inside option unsupported visually)
        this.updateElement('totalFilterCount', stats.total);
        this.updateElement('b2cFilterCount', stats.b2c);
        this.updateElement('ecomFilterCount', stats.ecom);
        this.updateElement('offlineFilterCount', stats.offline);
        this.updateElement('quickcomFilterCount', stats.quickcom);
        this.updateElement('eboFilterCount', stats.ebo);
        this.updateElement('othersFilterCount', stats.others);

        // Cache last LR missing dataset for filtering
        this.lastLRMissingData = data;

        // Update option labels to include counts (spans inside option often not displayed)
        const select = document.getElementById('lrCategoryFilter');
        if (select) {
            const labelMap = {
                all: `All Categories (${stats.total})`,
                b2c: `B2C (${stats.b2c})`,
                ecom: `E-commerce (${stats.ecom})`,
                offline: `Offline (${stats.offline})`,
                quickcom: `Quick-commerce (${stats.quickcom})`,
                ebo: `EBO (${stats.ebo})`,
                others: `Others (${stats.others})`
            };
            Array.from(select.options).forEach(opt => {
                const v = opt.value;
                if (labelMap[v]) opt.textContent = labelMap[v];
            });
        }
    }

    displayLRMissingTable(data) {
        const tableSection = document.getElementById('lrMissingTableSection');
        if (tableSection) tableSection.style.display = 'block';

        const container = document.getElementById('lrMissingByDay');
        if (!container) return;

        if (!data || data.length === 0) {
            container.innerHTML = '<div class="no-data">No LR missing records for selected date.</div>';
            this.updateElement('recordCount', '0 records');
            return;
        }

        // Build table
        const headers = [
            'Date','Customer Group','Sales Order No','SALES Invoice NO','DELIVERY Note NO','Quantity','SO Qty','CBM','Warehouse','Remarks'
        ];

        const rowsHtml = data.map(row => {
            const invoiceDate = row['SALES Invoice DATE'] || row['DELIVERY Note DATE'] || '';
            const dateDisp = invoiceDate ? new Date(invoiceDate).toLocaleDateString() : '';
            const qty = this.getQuantity(row);
            const soQty = this.getSalesOrderQty(row);
            const cbm = row['SI Total CBM'] || row['DN Total CBM'] || '';
            const warehouse = row['Set Source Warehouse'] || '';
            const category = this.getCategoryForRow(row);
            return `<tr data-category="${category}">
                <td>${dateDisp}</td>
                <td>${row['Customer Group'] || ''}</td>
                <td>${row['Sales Order No'] || ''}</td>
                <td>${row['SALES Invoice NO'] || ''}</td>
                <td>${row['DELIVERY Note NO'] || ''}</td>
                <td class="num">${qty}</td>
                <td class="num">${soQty}</td>
                <td class="num">${cbm}</td>
                <td>${warehouse}</td>
                <td>LR Missing</td>
            </tr>`;
        }).join('');

        container.innerHTML = `<table class="lr-missing-table">
            <thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
            <tbody>${rowsHtml}</tbody>
        </table>`;

        this.updateElement('recordCount', `${data.length} records`);
    }

    // Helper to determine category consistent across features
    getCategoryForRow(row) {
        const cg = (row['Customer Group'] || '').toLowerCase();
        if (this.compiledRegex.b2c.test(cg)) return 'b2c';
        if (this.compiledRegex.ecom.test(cg)) return 'ecom';
        if (this.compiledRegex.offline.test(cg)) return 'offline';
        if (this.compiledRegex.quickcom.test(cg)) return 'quickcom';
        if (this.compiledRegex.ebo.test(cg)) return 'ebo';
        return 'others';
    }

    updateElement(id, value) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = value;
        }
    }
}

// Global functions for LR Missing Dashboard
function generateLRMissingReport() {
    if (window.dashboard) {
        window.dashboard.generateLRMissingReport();
    }
}

function refreshLRMissing() {
    if (window.dashboard) {
        // Reset the form and clear results
        document.getElementById('lrDatePicker').value = '';
        document.getElementById('selectedDateInfo').style.display = 'none';
        document.getElementById('lrMissingSummarySection').style.display = 'none';
        const tableSection = document.getElementById('lrMissingTableSection');
        if (tableSection) tableSection.style.display = 'none';
        const container = document.getElementById('lrMissingByDay');
        if (container) container.innerHTML = '';
        
        // Optionally reload data
        if (window.dashboard.processedData) {
            console.log('LR Missing dashboard refreshed');
        }
    }
}

// Category filter for LR missing records
function filterByCategory() {
    if (!window.dashboard || !window.dashboard.lastLRMissingData) return;
    const select = document.getElementById('lrCategoryFilter');
    if (!select) return;
    const val = select.value;
    let data = window.dashboard.lastLRMissingData;
    if (val !== 'all') {
        data = data.filter(row => window.dashboard.getCategoryForRow(row) === val);
    }
    window.dashboard.displayLRMissingTable(data);
}

// Initialize the dashboard when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    // Check authentication before initializing dashboard
    if (window.auth && window.auth.isUserAuthenticated()) {
        window.dashboard = new MISDashboard();
    }
});