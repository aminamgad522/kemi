class ETAInvoiceExporter {
  constructor() {
    this.invoiceData = [];
    this.totalCount = 0;
    this.currentPage = 1;
    this.totalPages = 1;
    this.isProcessing = false;
    
    this.initializeElements();
    this.attachEventListeners();
    this.checkCurrentPage();
    this.setupProgressListener();
  }
  
  initializeElements() {
    this.elements = {
      countInfo: document.getElementById('countInfo'),
      totalCountText: document.getElementById('totalCountText'),
      status: document.getElementById('status'),
      closeBtn: document.getElementById('closeBtn'),
      jsonBtn: document.getElementById('jsonBtn'),
      excelBtn: document.getElementById('excelBtn'),
      pdfBtn: document.getElementById('pdfBtn'),
      progressContainer: null, // Will be created dynamically
      progressBar: null,
      progressText: null,
      checkboxes: {
        // Complete mapping of all checkboxes based on the images
        serialNumber: document.getElementById('option-serial-number'),
        detailsButton: document.getElementById('option-details-button'),
        documentType: document.getElementById('option-document-type'),
        documentVersion: document.getElementById('option-document-version'),
        status: document.getElementById('option-status'),
        issueDate: document.getElementById('option-issue-date'),
        submissionDate: document.getElementById('option-submission-date'),
        invoiceCurrency: document.getElementById('option-invoice-currency'),
        invoiceValue: document.getElementById('option-invoice-value'),
        vatAmount: document.getElementById('option-vat-amount'),
        taxDiscount: document.getElementById('option-tax-discount'),
        totalInvoice: document.getElementById('option-total-invoice'),
        internalNumber: document.getElementById('option-internal-number'),
        electronicNumber: document.getElementById('option-electronic-number'),
        sellerTaxNumber: document.getElementById('option-seller-tax-number'),
        sellerName: document.getElementById('option-seller-name'),
        sellerAddress: document.getElementById('option-seller-address'),
        buyerTaxNumber: document.getElementById('option-buyer-tax-number'),
        buyerName: document.getElementById('option-buyer-name'),
        buyerAddress: document.getElementById('option-buyer-address'),
        purchaseOrderRef: document.getElementById('option-purchase-order-ref'),
        purchaseOrderDesc: document.getElementById('option-purchase-order-desc'),
        salesOrderRef: document.getElementById('option-sales-order-ref'),
        electronicSignature: document.getElementById('option-electronic-signature'),
        foodDrugGuide: document.getElementById('option-food-drug-guide'),
        externalLink: document.getElementById('option-external-link'),
        // Special options
        downloadDetails: document.getElementById('option-download-details'),
        combineAll: document.getElementById('option-combine-all'),
        downloadAll: document.getElementById('option-download-all'),
        selectAll: document.getElementById('option-select-all')
      }
    };
    
    this.createProgressElements();
  }
  
  createProgressElements() {
    // Create progress container
    this.elements.progressContainer = document.createElement('div');
    this.elements.progressContainer.className = 'progress-container';
    this.elements.progressContainer.style.cssText = `
      margin-top: 15px;
      padding: 15px;
      background: #f8f9fa;
      border-radius: 8px;
      border: 1px solid #e9ecef;
      display: none;
    `;
    
    // Create progress bar
    this.elements.progressBar = document.createElement('div');
    this.elements.progressBar.className = 'progress-bar';
    this.elements.progressBar.style.cssText = `
      width: 100%;
      height: 20px;
      background: #e9ecef;
      border-radius: 10px;
      overflow: hidden;
      margin-bottom: 10px;
      position: relative;
    `;
    
    const progressFill = document.createElement('div');
    progressFill.className = 'progress-fill';
    progressFill.style.cssText = `
      height: 100%;
      background: linear-gradient(90deg, #28a745, #20c997);
      border-radius: 10px;
      width: 0%;
      transition: width 0.3s ease;
      position: relative;
    `;
    
    this.elements.progressBar.appendChild(progressFill);
    
    // Create progress text
    this.elements.progressText = document.createElement('div');
    this.elements.progressText.className = 'progress-text';
    this.elements.progressText.style.cssText = `
      text-align: center;
      font-size: 14px;
      color: #495057;
      font-weight: 500;
    `;
    
    // Assemble progress container
    this.elements.progressContainer.appendChild(this.elements.progressBar);
    this.elements.progressContainer.appendChild(this.elements.progressText);
    
    // Add to DOM
    const statusElement = this.elements.status;
    statusElement.parentNode.insertBefore(this.elements.progressContainer, statusElement.nextSibling);
  }
  
  attachEventListeners() {
    this.elements.closeBtn.addEventListener('click', () => window.close());
    this.elements.excelBtn.addEventListener('click', () => this.handleExport('excel'));
    this.elements.jsonBtn.addEventListener('click', () => this.handleExport('json'));
    this.elements.pdfBtn.addEventListener('click', () => this.handleExport('pdf'));
    
    // Add listener for download all checkbox
    if (this.elements.checkboxes.downloadAll) {
      this.elements.checkboxes.downloadAll.addEventListener('change', (e) => {
        if (e.target.checked) {
          this.showMultiPageWarning();
        }
      });
    }

    // Add listener for select all checkbox
    if (this.elements.checkboxes.selectAll) {
      this.elements.checkboxes.selectAll.addEventListener('change', (e) => {
        this.toggleAllCheckboxes(e.target.checked);
      });
    }
  }

  toggleAllCheckboxes(checked) {
    // Toggle all field checkboxes except special options
    const fieldCheckboxes = [
      'serialNumber', 'detailsButton', 'documentType', 'documentVersion', 'status',
      'issueDate', 'submissionDate', 'invoiceCurrency', 'invoiceValue', 'vatAmount',
      'taxDiscount', 'totalInvoice', 'internalNumber', 'electronicNumber',
      'sellerTaxNumber', 'sellerName', 'sellerAddress', 'buyerTaxNumber',
      'buyerName', 'buyerAddress', 'purchaseOrderRef', 'purchaseOrderDesc',
      'salesOrderRef', 'electronicSignature', 'foodDrugGuide', 'externalLink'
    ];

    fieldCheckboxes.forEach(field => {
      const checkbox = this.elements.checkboxes[field];
      if (checkbox) {
        checkbox.checked = checked;
      }
    });
  }
  
  setupProgressListener() {
    // Listen for progress updates from content script
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
      if (message.action === 'progressUpdate') {
        this.updateProgress(message.progress);
      }
    });
  }
  
  showMultiPageWarning() {
    const warningText = `تحذير: سيتم تحميل جميع الصفحات (${this.totalPages} صفحة) وقد يستغرق وقتاً أطول.`;
    this.showStatus(warningText, 'loading');
    
    setTimeout(() => {
      this.elements.status.textContent = '';
      this.elements.status.className = 'status';
    }, 3000);
  }
  
  async checkCurrentPage() {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      
      if (!tab.url.includes('invoicing.eta.gov.eg')) {
        this.showStatus('يرجى الانتقال إلى بوابة الفواتير الإلكترونية المصرية', 'error');
        this.disableButtons();
        return;
      }
      
      await this.loadInvoiceData();
    } catch (error) {
      this.showStatus('خطأ في فحص الصفحة الحالية', 'error');
      console.error('Error:', error);
    }
  }
  
  async loadInvoiceData() {
    try {
      this.showStatus('جاري تحميل بيانات الفواتير...', 'loading');
      
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      const response = await chrome.tabs.sendMessage(tab.id, { action: 'getInvoiceData' });
      
      if (!response || !response.success) {
        throw new Error('فشل في الحصول على بيانات الفواتير');
      }
      
      this.invoiceData = response.data.invoices || [];
      this.totalCount = response.data.totalCount || this.invoiceData.length;
      this.currentPage = response.data.currentPage || 1;
      this.totalPages = response.data.totalPages || 1;
      
      this.updateUI();
      this.showStatus('تم تحميل البيانات بنجاح', 'success');
      
    } catch (error) {
      this.showStatus('خطأ في تحميل البيانات: ' + error.message, 'error');
      console.error('Load error:', error);
    }
  }
  
  updateUI() {
    const currentPageCount = this.invoiceData.length;
    this.elements.countInfo.textContent = `الصفحة الحالية: ${currentPageCount} فاتورة | المجموع: ${this.totalCount} فاتورة`;
    
    if (this.elements.totalCountText) {
      this.elements.totalCountText.textContent = this.totalCount;
    }
    
    // Update download all option text
    const downloadAllLabel = this.elements.checkboxes.downloadAll?.parentElement.querySelector('label');
    if (downloadAllLabel) {
      downloadAllLabel.innerHTML = `تحميل جميع الصفحات - <span id="totalCountText">${this.totalCount}</span> فاتورة (${this.totalPages} صفحة)`;
    }
  }
  
  getSelectedOptions() {
    const options = {};
    Object.keys(this.elements.checkboxes).forEach(key => {
      const checkbox = this.elements.checkboxes[key];
      if (checkbox) {
        options[key] = checkbox.checked;
      } else {
        console.warn(`Checkbox not found: ${key}`);
        options[key] = false;
      }
    });
    return options;
  }
  
  async handleExport(format) {
    if (this.isProcessing) {
      this.showStatus('جاري المعالجة... يرجى الانتظار', 'loading');
      return;
    }
    
    const options = this.getSelectedOptions();
    
    if (!this.validateOptions(options)) {
      return;
    }
    
    this.isProcessing = true;
    this.disableButtons();
    
    try {
      if (options.downloadAll) {
        await this.exportAllPages(format, options);
      } else {
        await this.exportCurrentPage(format, options);
      }
    } catch (error) {
      this.showStatus('خطأ في التصدير: ' + error.message, 'error');
      console.error('Export error:', error);
    } finally {
      this.isProcessing = false;
      this.enableButtons();
      this.hideProgress();
    }
  }
  
  validateOptions(options) {
    const hasBasicField = options.serialNumber || options.electronicNumber || options.internalNumber || 
                         options.totalInvoice || options.invoiceValue || options.documentType;
    if (!hasBasicField) {
      this.showStatus('يرجى اختيار حقل واحد على الأقل للتصدير', 'error');
      return false;
    }
    return true;
  }
  
  async exportCurrentPage(format, options) {
    this.showStatus('جاري تصدير الصفحة الحالية...', 'loading');
    
    let dataToExport = [...this.invoiceData];
    
    if (options.downloadDetails) {
      this.showStatus('جاري تحميل تفاصيل الفواتير...', 'loading');
      dataToExport = await this.loadInvoiceDetails(dataToExport);
    }
    
    await this.generateFile(dataToExport, format, options);
    this.showStatus(`تم تصدير ${dataToExport.length} فاتورة بنجاح!`, 'success');
  }
  
  async exportAllPages(format, options) {
    this.showProgress();
    this.showStatus('جاري تحميل جميع الصفحات...', 'loading');
    
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    
    // Send message to content script to process all pages
    const allData = await chrome.tabs.sendMessage(tab.id, { 
      action: 'getAllPagesData',
      options: { ...options, progressCallback: true }
    });
    
    if (!allData || !allData.success) {
      throw new Error('فشل في تحميل جميع الصفحات: ' + (allData?.error || 'خطأ غير معروف'));
    }
    
    let dataToExport = allData.data;
    
    if (options.downloadDetails && dataToExport.length > 0) {
      this.updateProgress({
        currentPage: this.totalPages,
        totalPages: this.totalPages,
        message: 'جاري تحميل تفاصيل جميع الفواتير...'
      });
      
      dataToExport = await this.loadInvoiceDetails(dataToExport);
    }
    
    this.updateProgress({
      currentPage: this.totalPages,
      totalPages: this.totalPages,
      message: 'جاري إنشاء الملف...'
    });
    
    await this.generateFile(dataToExport, format, options);
    this.showStatus(`تم تصدير ${dataToExport.length} فاتورة من جميع الصفحات بنجاح!`, 'success');
  }
  
  showProgress() {
    this.elements.progressContainer.style.display = 'block';
    this.updateProgress({ currentPage: 0, totalPages: this.totalPages, message: 'جاري البدء...' });
  }
  
  hideProgress() {
    this.elements.progressContainer.style.display = 'none';
  }
  
  updateProgress(progress) {
    if (!this.elements.progressContainer || this.elements.progressContainer.style.display === 'none') {
      return;
    }
    
    const percentage = progress.percentage || (progress.totalPages > 0 ? (progress.currentPage / progress.totalPages) * 100 : 0);
    
    const progressFill = this.elements.progressBar.querySelector('.progress-fill');
    if (progressFill) {
      progressFill.style.width = `${Math.min(percentage, 100)}%`;
    }
    
    if (this.elements.progressText) {
      this.elements.progressText.textContent = progress.message || `الصفحة ${progress.currentPage} من ${progress.totalPages} (${Math.round(percentage)}%)`;
    }
  }
  
  async loadInvoiceDetails(invoices) {
    const detailedInvoices = [];
    const batchSize = 3; // Reduced batch size for better performance
    
    // Process invoices in batches to avoid overwhelming the server
    for (let i = 0; i < invoices.length; i += batchSize) {
      const batch = invoices.slice(i, i + batchSize);
      const batchPromises = batch.map(async (invoice, batchIndex) => {
        const globalIndex = i + batchIndex;
        this.showStatus(`جاري تحميل تفاصيل الفاتورة ${globalIndex + 1} من ${invoices.length}...`, 'loading');
        
        try {
          const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
          const detailResponse = await chrome.tabs.sendMessage(tab.id, {
            action: 'getInvoiceDetails',
            invoiceId: invoice.electronicNumber
          });
          
          if (detailResponse && detailResponse.success) {
            return {
              ...invoice,
              details: detailResponse.data
            };
          } else {
            return invoice;
          }
        } catch (error) {
          console.warn(`Failed to load details for invoice ${invoice.electronicNumber}:`, error);
          return invoice;
        }
      });
      
      // Wait for current batch to complete
      const batchResults = await Promise.all(batchPromises);
      detailedInvoices.push(...batchResults);
      
      // Small delay between batches to avoid rate limiting
      if (i + batchSize < invoices.length) {
        await new Promise(resolve => setTimeout(resolve, 300));
      }
    }
    
    return detailedInvoices;
  }
  
  async generateFile(data, format, options) {
    switch (format) {
      case 'excel':
        this.generateExcelFileWithCorrectLayout(data, options);
        break;
      case 'json':
        this.generateJSONFile(data, options);
        break;
      case 'pdf':
        this.showStatus('تصدير PDF غير متاح حاليًا', 'error');
        break;
    }
  }
  
  generateExcelFileWithCorrectLayout(data, options) {
    const wb = XLSX.utils.book_new();
    
    // Create main summary sheet with the exact layout from the screenshot
    this.createMainSheetWithCorrectLayout(wb, data, options);
    
    // Create details sheets for each invoice if requested
    if (options.downloadDetails) {
      this.createDetailsSheetsWithLinks(wb, data, options);
    }
    
    // Add statistics sheet
    this.createStatisticsSheet(wb, data, options);
    
    const timestamp = new Date().toISOString().split('T')[0].replace(/-/g, '');
    const pageInfo = options.downloadAll ? 'AllPages' : `Page${this.currentPage}`;
    const filename = `ETA_Invoices_${pageInfo}_${timestamp}.xlsx`;
    
    XLSX.writeFile(wb, filename);
  }
  
  createMainSheetWithCorrectLayout(wb, data, options) {
    // Headers based on the exact layout from the screenshot (left to right as shown)
    const headers = [
      'تسلسل',           // A - Serial number
      'عرض',             // B - View button
      'نوع الفاتورة',     // C - Invoice type
      'إصدار',           // D - Version
      'الحالة',          // E - Status
      'تاريخ الإصدار',    // F - Issue date
      'تاريخ التقديم',    // G - Submission date
      'عملة الفاتورة',    // H - Currency
      'قيمة الفاتورة',    // I - Invoice value
      'ضريبة القيمة المضافة', // J - VAT
      'الخصم تحت حساب الضريبة', // K - Tax discount
      'إجمالي الفاتورة',  // L - Total
      'الرقم الداخلي',    // M - Internal number
      'الرقم الإلكتروني'  // N - Electronic number
    ];

    const rows = [headers];
    
    data.forEach((invoice, index) => {
      const row = [
        index + 1,                                    // A - تسلسل
        'عرض',                                        // B - عرض (will be converted to hyperlink)
        invoice.documentType || 'فاتورة',             // C - نوع الفاتورة
        invoice.documentVersion || '1.0',             // D - إصدار
        invoice.status || '',                         // E - الحالة
        invoice.issueDate || '',                      // F - تاريخ الإصدار
        invoice.submissionDate || invoice.issueDate || '', // G - تاريخ التقديم
        invoice.currency || 'EGP',                    // H - عملة الفاتورة
        this.formatCurrency(invoice.invoiceValue || invoice.totalAmount), // I - قيمة الفاتورة
        this.formatCurrency(invoice.vatAmount),       // J - ضريبة القيمة المضافة
        this.formatCurrency(invoice.taxDiscount || '0'), // K - الخصم تحت حساب الضريبة
        this.formatCurrency(invoice.totalAmount),     // L - إجمالي الفاتورة
        invoice.internalNumber || '',                 // M - الرقم الداخلي
        invoice.electronicNumber || ''                // N - الرقم الإلكتروني
      ];
      
      rows.push(row);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    
    // Add hyperlinks to the "عرض" column (column B) that link to detail sheets
    this.addInternalHyperlinks(ws, data);
    
    // Format the worksheet
    this.formatMainWorksheet(ws, headers, data.length);
    
    XLSX.utils.book_append_sheet(wb, ws, 'فواتير مصلحة الضرائب');
  }
  
  addInternalHyperlinks(ws, data) {
    // Add hyperlinks to the "عرض" column (column B, index 1) that link to detail sheets
    data.forEach((invoice, index) => {
      const rowIndex = index + 2; // +2 because Excel is 1-indexed and we have a header row
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex - 1, c: 1 }); // Column B (index 1)
      
      if (ws[cellAddress] && invoice.electronicNumber) {
        const sheetName = `تفاصيل_${index + 1}`;
        
        // Create internal hyperlink to the detail sheet
        ws[cellAddress].l = { 
          Target: `#'${sheetName}'!A1`, 
          Tooltip: `عرض تفاصيل الفاتورة ${invoice.internalNumber || index + 1}` 
        };
        
        // Style the hyperlink
        ws[cellAddress].s = {
          font: { 
            color: { rgb: "0000FF" }, 
            underline: true,
            bold: true
          },
          alignment: { horizontal: "center", vertical: "center" }
        };
      }
    });
  }
  
  createDetailsSheetsWithLinks(wb, data, options) {
    data.forEach((invoice, index) => {
      // Create detail sheet for each invoice
      const sheetName = `تفاصيل_${index + 1}`;
      
      // Headers for invoice details matching the screenshot layout (right to left)
      const detailHeaders = [
        'اسم الصنف',                    // A - Item name (rightmost)
        'كود الوحدة',                  // B - Unit code  
        'اسم الوحدة',                  // C - Unit name
        'الكمية',                     // D - Quantity
        'السعر',                      // E - Price
        'القيمة',                     // F - Value
        'الضريبة',                    // G - Tax
        'ضريبة القيمة المضافة',        // H - VAT
        'الإجمالي'                    // I - Total (leftmost)
      ];
      
      const detailRows = [
        // Invoice header information (9 columns to match headers)
        ['معلومات الفاتورة', '', '', '', '', '', '', '', ''],
        ['الرقم الإلكتروني:', invoice.electronicNumber || '', '', '', '', '', '', '', ''],
        ['الرقم الداخلي:', invoice.internalNumber || '', '', '', '', '', '', '', ''],
        ['التاريخ:', invoice.issueDate || '', '', '', '', '', '', '', ''],
        ['البائع:', invoice.sellerName || '', '', '', '', '', '', '', ''],
        ['المشتري:', invoice.buyerName || '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''], // Empty row
        detailHeaders
      ];
      
      // Add invoice line items if available
      if (invoice.details && invoice.details.length > 0) {
        invoice.details.forEach(item => {
          detailRows.push([
            item.description || item.codeName || '',           // A - اسم الصنف
            item.unitCode || 'EA',                            // B - كود الوحدة
            item.unitName || 'قطعة',                          // C - اسم الوحدة
            item.quantity || '1',
            item.unitName || 'قطعة',
            this.formatCurrency(item.unitPrice),
            this.formatCurrency(item.totalValue),
            this.formatCurrency(item.taxAmount || '0'),       // G - الضريبة
            this.formatCurrency(item.vatAmount),
            this.formatCurrency(item.totalWithVat)            // I - الإجمالي
          ]);
        });
      } else {
        // Add summary row if no detailed items
        detailRows.push([
          'إجمالي قيمة الفاتورة',
          'EA',
          'فاتورة',
          '1',
          this.formatCurrency(invoice.totalAmount),
          this.formatCurrency(invoice.invoiceValue || invoice.totalAmount),
          '0',
          this.formatCurrency(invoice.vatAmount),
          this.formatCurrency(invoice.totalAmount)
        ]);
      }
      
      // Add totals row
      const totalValue = invoice.invoiceValue || invoice.totalAmount || '0';
      const totalVat = invoice.vatAmount || '0';
      const grandTotal = invoice.totalAmount || '0';
      const totalTax = '0'; // Default tax amount
      
      detailRows.push([
        '', '', '', '', 'الإجمالي:', 
        this.formatCurrency(totalValue),
        this.formatCurrency(totalTax),
        this.formatCurrency(totalVat),
        this.formatCurrency(grandTotal)
      ]);
      
      // Add back link to main sheet
      detailRows.push(['', '', '', '', '', '', '', '', '']);
      detailRows.push(['العودة للقائمة الرئيسية', '', '', '', '', '', '', '', '']);
      
      const ws = XLSX.utils.aoa_to_sheet(detailRows);
      
      // Add back link to main sheet
      const backLinkCell = XLSX.utils.encode_cell({ r: detailRows.length - 1, c: 0 });
      if (ws[backLinkCell]) {
        ws[backLinkCell].l = { 
          Target: "#'فواتير مصلحة الضرائب'!A1", 
          Tooltip: 'العودة للقائمة الرئيسية' 
        };
        ws[backLinkCell].s = {
          font: { color: { rgb: "0000FF" }, underline: true, bold: true }
        };
      }
      
      // Format the details sheet
      this.formatDetailsWorksheet(ws, detailHeaders);
      
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
  }
  
  formatCurrency(value) {
    if (!value || value === '0' || value === '') return '';
    
    // Remove any existing formatting and convert to number
    const numValue = parseFloat(value.toString().replace(/[,٬]/g, ''));
    if (isNaN(numValue)) return value;
    
    // Format with commas for thousands
    return numValue.toLocaleString('en-US', { 
      minimumFractionDigits: 2, 
      maximumFractionDigits: 2 
    });
  }
  
  formatMainWorksheet(ws, headers, dataLength) {
    // Set column widths based on content
    const colWidths = [
      { wch: 8 },   // A - تسلسل
      { wch: 10 },  // B - عرض
      { wch: 15 },  // C - نوع الفاتورة
      { wch: 8 },   // D - إصدار
      { wch: 15 },  // E - الحالة
      { wch: 18 },  // F - تاريخ الإصدار
      { wch: 18 },  // G - تاريخ التقديم
      { wch: 12 },  // H - عملة الفاتورة
      { wch: 15 },  // I - قيمة الفاتورة
      { wch: 18 },  // J - ضريبة القيمة المضافة
      { wch: 20 },  // K - الخصم تحت حساب الضريبة
      { wch: 15 },  // L - إجمالي الفاتورة
      { wch: 20 },  // M - الرقم الداخلي
      { wch: 30 }   // N - الرقم الإلكتروني
    ];
    
    ws['!cols'] = colWidths;
    
    // Style the header row
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      
      ws[cellAddress].s = {
        font: { bold: true, color: { rgb: "FFFFFF" }, size: 12 },
        fill: { fgColor: { rgb: "1F4E79" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        }
      };
    }
    
    // Style data rows with alternating colors
    for (let row = 1; row <= range.e.r; row++) {
      const isEvenRow = row % 2 === 0;
      const fillColor = isEvenRow ? "F8F9FA" : "FFFFFF";
      
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;
        
        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        
        ws[cellAddress].s.fill = { fgColor: { rgb: fillColor } };
        ws[cellAddress].s.alignment = { horizontal: "center", vertical: "center" };
        ws[cellAddress].s.border = {
          top: { style: "thin", color: { rgb: "E0E0E0" } },
          bottom: { style: "thin", color: { rgb: "E0E0E0" } },
          left: { style: "thin", color: { rgb: "E0E0E0" } },
          right: { style: "thin", color: { rgb: "E0E0E0" } }
        };
      }
    }
  }
  
  formatDetailsWorksheet(ws, headers) {
    // Set column widths for details sheet
    const colWidths = [
      { wch: 35 }, // A - اسم الصنف (widest column)
      { wch: 12 }, // B - كود الوحدة
      { wch: 15 }, // C - اسم الوحدة
      { wch: 10 }, // الكمية
      { wch: 12 }, // السعر
      { wch: 12 }, // القيمة
      { wch: 12 }, // الضريبة
      { wch: 15 }, // ضريبة القيمة المضافة
      { wch: 12 }  // الإجمالي
    ];
    
    ws['!cols'] = colWidths;
    
    // Style the header row (row 8, where detail headers start)
    const range = XLSX.utils.decode_range(ws['!ref']);
    const headerRow = 7; // 0-indexed, so row 8
    
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
      if (!ws[cellAddress]) continue;
      
      ws[cellAddress].s = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "2196F3" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        }
      };
    }
    
    // Add alternating row colors for data rows (starting from row 9)
    for (let row = 8; row <= range.e.r - 2; row++) { // Exclude totals and back link rows
      const isEvenRow = (row - 8) % 2 === 0;
      const fillColor = isEvenRow ? "F8F9FA" : "FFFFFF";
      
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;
        
        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        
        ws[cellAddress].s.fill = { fgColor: { rgb: fillColor } };
        ws[cellAddress].s.alignment = { horizontal: "center", vertical: "center" };
        ws[cellAddress].s.border = {
          top: { style: "thin", color: { rgb: "E0E0E0" } },
          bottom: { style: "thin", color: { rgb: "E0E0E0" } },
          left: { style: "thin", color: { rgb: "E0E0E0" } },
          right: { style: "thin", color: { rgb: "E0E0E0" } }
        };
      }
    }
  }
  
  createStatisticsSheet(wb, data, options) {
    const stats = this.calculateStatistics(data);
    
    const statsData = [
      ['إحصائيات الفواتير', ''],
      ['', ''],
      ['إجمالي عدد الفواتير', data.length],
      ['إجمالي قيمة الفواتير', this.formatCurrency(stats.totalValue) + ' EGP'],
      ['إجمالي ضريبة القيمة المضافة', this.formatCurrency(stats.totalVAT) + ' EGP'],
      ['متوسط قيمة الفاتورة', this.formatCurrency(stats.averageValue) + ' EGP'],
      ['', ''],
      ['إحصائيات حسب الحالة', ''],
      ...Object.entries(stats.statusCounts).map(([status, count]) => [status, count]),
      ['', ''],
      ['إحصائيات حسب النوع', ''],
      ...Object.entries(stats.typeCounts).map(([type, count]) => [type, count]),
      ['', ''],
      ['تاريخ التصدير', new Date().toLocaleString('ar-EG')],
      ['نوع التصدير', options.downloadAll ? 'جميع الصفحات' : 'الصفحة الحالية'],
      ['عدد الحقول المصدرة', Object.values(options).filter(Boolean).length]
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(statsData);
    
    // Format statistics sheet
    ws['!cols'] = [{ wch: 30 }, { wch: 20 }];
    
    XLSX.utils.book_append_sheet(wb, ws, 'الإحصائيات');
  }
  
  calculateStatistics(data) {
    const stats = {
      totalValue: 0,
      totalVAT: 0,
      averageValue: 0,
      statusCounts: {},
      typeCounts: {}
    };
    
    data.forEach(invoice => {
      // Calculate totals
      const value = parseFloat(invoice.totalAmount?.replace(/[,٬]/g, '') || 0);
      const vat = parseFloat(invoice.vatAmount?.replace(/[,٬]/g, '') || 0);
      
      stats.totalValue += value;
      stats.totalVAT += vat;
      
      // Count statuses
      const status = invoice.status || 'غير محدد';
      stats.statusCounts[status] = (stats.statusCounts[status] || 0) + 1;
      
      // Count types
      const type = invoice.documentType || 'غير محدد';
      stats.typeCounts[type] = (stats.typeCounts[type] || 0) + 1;
    });
    
    stats.averageValue = data.length > 0 ? stats.totalValue / data.length : 0;
    
    return stats;
  }
  
  generateJSONFile(data, options) {
    const jsonData = {
      exportDate: new Date().toISOString(),
      totalCount: data.length,
      exportType: options.downloadAll ? 'all_pages' : 'current_page',
      totalPages: this.totalPages,
      currentPage: this.currentPage,
      selectedFields: Object.keys(options).filter(key => options[key]),
      options: options,
      statistics: this.calculateStatistics(data),
      invoices: data
    };
    
    const blob = new Blob([JSON.stringify(jsonData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    const timestamp = new Date().toISOString().split('T')[0];
    const pageInfo = options.downloadAll ? 'AllPages' : `Page${this.currentPage}`;
    a.download = `ETA_Invoices_${pageInfo}_${timestamp}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  
  showStatus(message, type = '') {
    this.elements.status.textContent = message;
    this.elements.status.className = `status ${type}`;
    
    if (type === 'loading') {
      this.elements.status.innerHTML = `${message} <span class="loading-spinner"></span>`;
    }
    
    if (type === 'success' || type === 'error') {
      setTimeout(() => {
        if (!this.isProcessing) {
          this.elements.status.textContent = '';
          this.elements.status.className = 'status';
        }
      }, 3000);
    }
  }
  
  disableButtons() {
    this.elements.excelBtn.disabled = true;
    this.elements.jsonBtn.disabled = true;
    this.elements.pdfBtn.disabled = true;
  }
  
  enableButtons() {
    this.elements.excelBtn.disabled = false;
    this.elements.jsonBtn.disabled = false;
    this.elements.pdfBtn.disabled = false;
  }
}

// Initialize the exporter when popup loads
document.addEventListener('DOMContentLoaded', () => {
  new ETAInvoiceExporter();
});