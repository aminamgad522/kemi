// Content script for ETA Invoice Exporter with improved performance and accurate data extraction
class ETAContentScript {
  constructor() {
    this.invoiceData = [];
    this.allPagesData = [];
    this.totalCount = 0;
    this.currentPage = 1;
    this.totalPages = 1;
    this.isProcessingAllPages = false;
    this.progressCallback = null;
    this.domObserver = null;
    this.init();
  }
  
  init() {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => this.scanForInvoices());
    } else {
      this.scanForInvoices();
    }
    
    this.setupMutationObserver();
  }
  
  setupMutationObserver() {
    this.observer = new MutationObserver((mutations) => {
      let shouldRescan = false;
      
      mutations.forEach((mutation) => {
        if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              if (node.classList?.contains('ms-DetailsRow') || 
                  node.querySelector?.('.ms-DetailsRow') ||
                  node.classList?.contains('ms-List-cell') ||
                  node.classList?.contains('eta-pageNumber')) {
                shouldRescan = true;
              }
            }
          });
        }
      });
      
      if (shouldRescan && !this.isProcessingAllPages) {
        clearTimeout(this.rescanTimeout);
        this.rescanTimeout = setTimeout(() => this.scanForInvoices(), 500);
      }
    });
    
    this.observer.observe(document.body, {
      childList: true,
      subtree: true
    });
  }
  
  scanForInvoices() {
    try {
      this.invoiceData = [];
      
      // Extract pagination info first
      this.extractPaginationInfo();
      
      // Find invoice rows using improved selectors - only visible, actual invoice rows
      const rows = this.getVisibleInvoiceRows();
      console.log(`ETA Exporter: Found ${rows.length} visible invoice rows on page ${this.currentPage}`);
      
      rows.forEach((row, index) => {
        const invoiceData = this.extractDataFromRow(row, index + 1);
        if (this.isValidInvoiceData(invoiceData)) {
          this.invoiceData.push(invoiceData);
        }
      });
      
      console.log(`ETA Exporter: Extracted ${this.invoiceData.length} valid invoices from page ${this.currentPage}`);
      
    } catch (error) {
      console.error('ETA Exporter: Error scanning for invoices:', error);
    }
  }
  
  getVisibleInvoiceRows() {
    // Get all potential invoice rows
    const allRows = document.querySelectorAll('.ms-DetailsRow[role="row"]');
    const visibleRows = [];
    
    allRows.forEach(row => {
      // Check if row is visible and contains actual invoice data
      if (this.isRowVisible(row) && this.hasInvoiceData(row)) {
        visibleRows.push(row);
      }
    });
    
    return visibleRows;
  }
  
  isRowVisible(row) {
    // Check if the row is actually visible in the DOM
    const rect = row.getBoundingClientRect();
    const style = window.getComputedStyle(row);
    
    return (
      rect.width > 0 && 
      rect.height > 0 &&
      style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      style.opacity !== '0'
    );
  }
  
  hasInvoiceData(row) {
    // Check if row contains actual invoice data (not header or empty row)
    const cells = row.querySelectorAll('.ms-DetailsRow-cell');
    if (cells.length === 0) return false;
    
    // Look for electronic number or internal number in the first cell
    const firstCell = cells[0];
    const electronicLink = firstCell?.querySelector('.internalId-link a.griCellTitle');
    const internalNumber = firstCell?.querySelector('.griCellSubTitle');
    
    return !!(electronicLink?.textContent?.trim() || internalNumber?.textContent?.trim());
  }
  
  extractPaginationInfo() {
    try {
      // Extract total count from pagination - improved selector
      const totalLabel = document.querySelector('.eta-pagination-totalrecordCount-label, [class*="pagination"] [class*="total"], [class*="record"] [class*="count"]');
      if (totalLabel) {
        const match = totalLabel.textContent.match(/النتائج:\s*(\d+)|(\d+)\s*نتيجة|Total:\s*(\d+)/);
        if (match) {
          this.totalCount = parseInt(match[1] || match[2] || match[3]);
        }
      }
      
      // Extract current page - improved selector
      const currentPageBtn = document.querySelector('.eta-pageNumber.is-checked, [class*="page"][class*="current"], [class*="active"][class*="page"]');
      if (currentPageBtn) {
        const pageLabel = currentPageBtn.querySelector('.ms-Button-label, [class*="label"], [class*="text"]');
        if (pageLabel) {
          this.currentPage = parseInt(pageLabel.textContent) || 1;
        }
      }
      
      // Calculate total pages based on actual visible rows per page
      const visibleRows = this.getVisibleInvoiceRows();
      const itemsPerPage = Math.max(visibleRows.length, 10); // Default to 10 if no rows found
      this.totalPages = Math.ceil(this.totalCount / itemsPerPage);
      
      console.log(`ETA Exporter: Page ${this.currentPage} of ${this.totalPages}, Total: ${this.totalCount} invoices, Current page: ${visibleRows.length} rows`);
      
    } catch (error) {
      console.warn('ETA Exporter: Error extracting pagination info:', error);
    }
  }
  
  extractDataFromRow(row, index) {
    // Extract data exactly as shown in the HTML structure
    const invoice = {
      index: index,
      pageNumber: this.currentPage,
      
      // Main invoice data from the table structure
      electronicNumber: '',
      internalNumber: '',
      issueDate: '',
      issueTime: '',
      documentType: '',
      documentVersion: '',
      totalAmount: '',
      currency: 'EGP',
      sellerName: '',
      sellerTaxNumber: '',
      buyerName: '',
      buyerTaxNumber: '',
      submissionId: '',
      status: '',
      
      // Additional calculated fields
      invoiceValue: '',
      vatAmount: '',
      totalInvoice: '',
      
      // Default values for missing fields
      taxDiscount: '0',
      sellerAddress: '',
      buyerAddress: '',
      purchaseOrderRef: '',
      purchaseOrderDesc: '',
      salesOrderRef: '',
      electronicSignature: 'موقع إلكترونياً',
      foodDrugGuide: '',
      externalLink: '',
      
      // Add details property for line items
      details: []
    };
    
    try {
      // Get all cells in the row - exact structure from HTML
      const cells = row.querySelectorAll('.ms-DetailsRow-cell');
      
      if (cells.length === 0) {
        console.warn(`No cells found in row ${index}`);
        return invoice;
      }
      
      // Cell 0: Electronic Number and Internal Number (data-automation-key="uuid")
      const uuidCell = cells[0];
      if (uuidCell && uuidCell.getAttribute('data-automation-key') === 'uuid') {
        // Extract electronic number from the link
        const electronicLink = uuidCell.querySelector('.internalId-link a.griCellTitle');
        if (electronicLink) {
          invoice.electronicNumber = electronicLink.textContent?.trim() || '';
        }
        
        // Extract internal number from subtitle
        const internalNumberElement = uuidCell.querySelector('.griCellSubTitle');
        if (internalNumberElement) {
          invoice.internalNumber = internalNumberElement.textContent?.trim() || '';
        }
      }
      
      // Cell 1: Date and Time (data-automation-key="dateTimeReceived")
      const dateCell = cells[1];
      if (dateCell && dateCell.getAttribute('data-automation-key') === 'dateTimeReceived') {
        const dateElement = dateCell.querySelector('.griCellTitleGray');
        const timeElement = dateCell.querySelector('.griCellSubTitle');
        
        if (dateElement) {
          invoice.issueDate = dateElement.textContent?.trim() || '';
        }
        if (timeElement) {
          invoice.issueTime = timeElement.textContent?.trim() || '';
        }
      }
      
      // Cell 2: Document Type and Version (data-automation-key="typeName")
      const typeCell = cells[2];
      if (typeCell && typeCell.getAttribute('data-automation-key') === 'typeName') {
        const typeElement = typeCell.querySelector('.griCellTitleGray');
        const versionElement = typeCell.querySelector('.griCellSubTitle');
        
        if (typeElement) {
          invoice.documentType = typeElement.textContent?.trim() || '';
        }
        if (versionElement) {
          invoice.documentVersion = versionElement.textContent?.trim() || '';
        }
      }
      
      // Cell 3: Total Amount (data-automation-key="total")
      const totalCell = cells[3];
      if (totalCell && totalCell.getAttribute('data-automation-key') === 'total') {
        const totalElement = totalCell.querySelector('.griCellTitleGray');
        if (totalElement) {
          invoice.totalAmount = totalElement.textContent?.trim() || '';
          invoice.totalInvoice = invoice.totalAmount;
          invoice.invoiceValue = invoice.totalAmount;
        }
      }
      
      // Cell 4: Seller/Issuer Information (data-automation-key="issuerName")
      const issuerCell = cells[4];
      if (issuerCell && issuerCell.getAttribute('data-automation-key') === 'issuerName') {
        const sellerNameElement = issuerCell.querySelector('.griCellTitleGray');
        const sellerTaxElement = issuerCell.querySelector('.griCellSubTitle');
        
        if (sellerNameElement) {
          invoice.sellerName = sellerNameElement.textContent?.trim() || '';
        }
        if (sellerTaxElement) {
          invoice.sellerTaxNumber = sellerTaxElement.textContent?.trim() || '';
        }
      }
      
      // Cell 5: Buyer/Receiver Information (data-automation-key="receiverName")
      const receiverCell = cells[5];
      if (receiverCell && receiverCell.getAttribute('data-automation-key') === 'receiverName') {
        const buyerNameElement = receiverCell.querySelector('.griCellTitleGray');
        const buyerTaxElement = receiverCell.querySelector('.griCellSubTitle');
        
        if (buyerNameElement) {
          invoice.buyerName = buyerNameElement.textContent?.trim() || '';
        }
        if (buyerTaxElement) {
          invoice.buyerTaxNumber = buyerTaxElement.textContent?.trim() || '';
        }
      }
      
      // Cell 6: Submission ID (data-automation-key="submission")
      const submissionCell = cells[6];
      if (submissionCell && submissionCell.getAttribute('data-automation-key') === 'submission') {
        const submissionLink = submissionCell.querySelector('a.submissionId-link');
        if (submissionLink) {
          invoice.submissionId = submissionLink.textContent?.trim() || '';
          invoice.purchaseOrderRef = invoice.submissionId;
        }
      }
      
      // Cell 7: Status (data-automation-key="status")
      const statusCell = cells[7];
      if (statusCell && statusCell.getAttribute('data-automation-key') === 'status') {
        // Handle different status types exactly as in HTML
        const validRejectedDiv = statusCell.querySelector('.horizontal.valid-rejected');
        if (validRejectedDiv) {
          // Complex status: Valid → Rejected
          const validStatus = validRejectedDiv.querySelector('.status-Valid');
          const rejectedStatus = validRejectedDiv.querySelector('.status-Rejected');
          if (validStatus && rejectedStatus) {
            invoice.status = `${validStatus.textContent?.trim()} → ${rejectedStatus.textContent?.trim()}`;
          }
        } else {
          // Simple status
          const textStatus = statusCell.querySelector('.textStatus');
          if (textStatus) {
            invoice.status = textStatus.textContent?.trim() || '';
          } else {
            // Fallback to any text content
            invoice.status = statusCell.textContent?.trim() || '';
          }
        }
      }
      
      // Calculate VAT amount (assuming 14% VAT rate common in Egypt)
      if (invoice.totalAmount) {
        const totalValue = parseFloat(invoice.totalAmount.replace(/[,٬]/g, ''));
        if (!isNaN(totalValue)) {
          // Calculate VAT as 14% of the net amount
          invoice.vatAmount = (totalValue * 0.14 / 1.14).toFixed(2);
          invoice.invoiceValue = (totalValue - parseFloat(invoice.vatAmount)).toFixed(2);
        }
      }
      
      // Set default addresses if names are available
      if (invoice.sellerName && !invoice.sellerAddress) {
        invoice.sellerAddress = 'غير محدد';
      }
      if (invoice.buyerName && !invoice.buyerAddress) {
        invoice.buyerAddress = 'غير محدد';
      }
      
      // Generate external link for sharing
      if (invoice.electronicNumber) {
        invoice.externalLink = this.generateExternalLink(invoice);
      }
      
    } catch (error) {
      console.warn(`ETA Exporter: Error extracting data from row ${index}:`, error);
    }
    
    return invoice;
  }
  
  generateExternalLink(invoice) {
    if (!invoice.electronicNumber) {
      return '';
    }
    
    // Generate share link based on the submission ID or electronic number
    let shareId = '';
    if (invoice.submissionId && invoice.submissionId.length > 10) {
      shareId = invoice.submissionId;
    } else {
      // Generate a placeholder shareId based on electronic number
      shareId = invoice.electronicNumber.replace(/[^A-Z0-9]/g, '').substring(0, 26);
    }
    
    return `https://invoicing.eta.gov.eg/documents/${invoice.electronicNumber}/share/${shareId}`;
  }
  
  isValidInvoiceData(invoice) {
    return !!(invoice.electronicNumber || invoice.internalNumber || invoice.totalAmount);
  }
  
  async getAllPagesData(options = {}) {
    try {
      this.isProcessingAllPages = true;
      this.allPagesData = [];
      
      // أولاً، فحص الصفحة الحالية للحصول على معلومات التصفح
      this.scanForInvoices();
      
      console.log(`ETA Exporter: بدء تحميل جميع الصفحات. الصفحة الحالية: ${this.currentPage}, إجمالي الصفحات: ${this.totalPages}`);
      
      if (this.totalPages <= 1) {
        // صفحة واحدة فقط، إرجاع البيانات الحالية
        this.allPagesData = [...this.invoiceData];
        console.log(`ETA Exporter: صفحة واحدة فقط، تم جمع ${this.allPagesData.length} فاتورة`);
        return {
          success: true,
          data: this.allPagesData,
          totalProcessed: this.allPagesData.length
        };
      }
      
      // إضافة بيانات الصفحة الحالية أولاً
      if (this.invoiceData.length > 0) {
        this.allPagesData.push(...this.invoiceData);
        console.log(`ETA Exporter: تم إضافة بيانات الصفحة الحالية ${this.currentPage}: ${this.invoiceData.length} فاتورة`);
      }
      
      // معالجة الصفحات المتبقية
      for (let page = 1; page <= this.totalPages; page++) {
        // تخطي الصفحة الحالية لأن لدينا بياناتها بالفعل
        if (page === this.currentPage) {
          console.log(`ETA Exporter: تخطي الصفحة الحالية ${page}`);
          continue;
        }
        
        try {
          // تحديث التقدم
          if (this.progressCallback) {
            this.progressCallback({
              currentPage: page,
              totalPages: this.totalPages,
              message: `جاري الانتقال للصفحة ${page} من ${this.totalPages}...`,
              percentage: ((page - 1) / this.totalPages) * 100
            });
          }
          
          console.log(`ETA Exporter: محاولة الانتقال للصفحة ${page}`);
          
          // الانتقال للصفحة
          const navigated = await this.navigateToPage(page);
          if (!navigated) {
            console.warn(`ETA Exporter: فشل في الانتقال للصفحة ${page}`);
            continue;
          }
          
          console.log(`ETA Exporter: تم الانتقال بنجاح للصفحة ${page}`);
          
          // انتظار تحميل الصفحة
          await this.waitForPageLoadOptimized(page);
          
          // تحديث التقدم
          if (this.progressCallback) {
            this.progressCallback({
              currentPage: page,
              totalPages: this.totalPages,
              message: `جاري استخراج بيانات الصفحة ${page}...`,
              percentage: (page / this.totalPages) * 100
            });
          }
          
          // فحص الفواتير في هذه الصفحة
          this.scanForInvoices();
          
          // إضافة بيانات الصفحة الحالية لجميع بيانات الصفحات
          if (this.invoiceData.length > 0) {
            this.allPagesData.push(...this.invoiceData);
            console.log(`ETA Exporter: تم معالجة الصفحة ${page}، تم جمع ${this.invoiceData.length} فاتورة`);
          } else {
            console.warn(`ETA Exporter: لم يتم العثور على فواتير في الصفحة ${page}`);
          }
          
          // تأخير صغير قبل الصفحة التالية
          await this.delay(800);
          
        } catch (error) {
          console.error(`ETA Exporter: خطأ في معالجة الصفحة ${page}:`, error);
          // المتابعة مع الصفحة التالية حتى لو فشلت الصفحة الحالية
        }
      }
      
      console.log(`ETA Exporter: انتهاء تحميل جميع الصفحات. إجمالي الفواتير المجمعة: ${this.allPagesData.length}`);
      
      return {
        success: true,
        data: this.allPagesData,
        totalProcessed: this.allPagesData.length
      };
      
    } catch (error) {
      console.error('ETA Exporter: خطأ في الحصول على بيانات جميع الصفحات:', error);
      return { 
        success: false, 
        data: this.allPagesData,
        error: error.message 
      };
    } finally {
      this.isProcessingAllPages = false;
    }
  }
  
  async navigateToPage(pageNumber) {
    try {
      // إذا كنا بالفعل في الصفحة المطلوبة، فلا حاجة للتنقل
      if (this.currentPage === pageNumber) {
        console.log(`ETA Exporter: نحن بالفعل في الصفحة ${pageNumber}`);
        return true;
      }
      
      console.log(`ETA Exporter: البحث عن زر الصفحة ${pageNumber}`);
      
      // البحث عن زر رقم الصفحة باستخدام محددات متعددة
      const pageSelectors = [
        '.eta-pageNumber',
        '[class*="pageNumber"]',
        '[class*="page-number"]',
        '.ms-Button[aria-label*="Page"]',
        'button[aria-label*="صفحة"]'
      ];
      
      let pageButtons = [];
      for (const selector of pageSelectors) {
        pageButtons = document.querySelectorAll(selector);
        if (pageButtons.length > 0) {
          console.log(`ETA Exporter: تم العثور على ${pageButtons.length} أزرار صفحات باستخدام المحدد: ${selector}`);
          break;
        }
      }
      
      // البحث عن زر الصفحة المحدد
      for (const btn of pageButtons) {
        const label = btn.querySelector('.ms-Button-label, .ms-Button-textContainer, [class*="label"]');
        const buttonText = label ? label.textContent : btn.textContent;
        
        if (buttonText && parseInt(buttonText.trim()) === pageNumber) {
          console.log(`ETA Exporter: تم العثور على زر الصفحة ${pageNumber}, النقر عليه`);
          btn.click();
          await this.delay(1000); // انتظار قصير للسماح بالنقر
          
          // التحقق من أننا في الصفحة الصحيحة
          await this.waitForPageLoadOptimized(pageNumber);
          this.extractPaginationInfo();
          
          if (this.currentPage === pageNumber) {
            console.log(`ETA Exporter: تم الانتقال بنجاح للصفحة ${pageNumber}`);
            return true;
          } else {
            console.warn(`ETA Exporter: فشل في التحقق من الانتقال للصفحة ${pageNumber}. الصفحة الحالية: ${this.currentPage}`);
          }
        }
      }
      
      console.log(`ETA Exporter: لم يتم العثور على زر مباشر للصفحة ${pageNumber}, محاولة التنقل التدريجي`);
      
      // محاولة طريقة التنقل البديلة باستخدام أزرار التالي/السابق
      if (pageNumber > this.currentPage) {
        // التنقل للأمام
        for (let i = this.currentPage; i < pageNumber; i++) {
          console.log(`ETA Exporter: الانتقال للأمام من الصفحة ${i} إلى ${i + 1}`);
          const success = await this.navigateToNextPage();
          if (!success) {
            console.warn(`ETA Exporter: فشل في الانتقال للصفحة التالية من ${i}`);
            break;
          }
          await this.waitForPageLoadOptimized(i + 1);
          this.extractPaginationInfo();
          console.log(`ETA Exporter: الصفحة الحالية بعد التنقل: ${this.currentPage}`);
        }
      } else if (pageNumber < this.currentPage) {
        // التنقل للخلف
        for (let i = this.currentPage; i > pageNumber; i--) {
          console.log(`ETA Exporter: الانتقال للخلف من الصفحة ${i} إلى ${i - 1}`);
          const success = await this.navigateToPreviousPage();
          if (!success) {
            console.warn(`ETA Exporter: فشل في الانتقال للصفحة السابقة من ${i}`);
            break;
          }
          await this.waitForPageLoadOptimized(i - 1);
          this.extractPaginationInfo();
          console.log(`ETA Exporter: الصفحة الحالية بعد التنقل: ${this.currentPage}`);
        }
      }
      
      const success = this.currentPage === pageNumber;
      console.log(`ETA Exporter: نتيجة التنقل للصفحة ${pageNumber}: ${success ? 'نجح' : 'فشل'}`);
      return success;
    } catch (error) {
      console.error(`ETA Exporter: خطأ في الانتقال للصفحة ${pageNumber}:`, error);
      return false;
    }
  }
  
  async navigateToPreviousPage() {
    try {
      // البحث عن زر الصفحة السابقة
      const prevSelectors = [
        '[data-icon-name="ChevronLeft"]',
        '[data-icon-name="Previous"]',
        '[aria-label*="Previous"]',
        '[aria-label*="السابق"]',
        '.ms-Button[title*="Previous"]',
        '.ms-Button[title*="السابق"]'
      ];
      
      for (const selector of prevSelectors) {
        const prevButton = document.querySelector(selector)?.closest('button');
        if (prevButton && !prevButton.disabled && !prevButton.classList.contains('is-disabled')) {
          console.log(`ETA Exporter: النقر على زر الصفحة السابقة`);
          prevButton.click();
          await this.delay(500);
          return true;
        }
      }
      
      console.warn(`ETA Exporter: لم يتم العثور على زر الصفحة السابقة`);
      return false;
    } catch (error) {
      console.error('ETA Exporter: خطأ في الانتقال للصفحة السابقة:', error);
      return false;
    }
  }
  
  async navigateToNextPage() {
    try {
      // البحث عن زر الصفحة التالية
      const nextSelectors = [
        '[data-icon-name="ChevronRight"]',
        '[data-icon-name="Next"]',
        '[aria-label*="Next"]',
        '[aria-label*="التالي"]',
        '.ms-Button[title*="Next"]',
        '.ms-Button[title*="التالي"]'
      ];
      
      for (const selector of nextSelectors) {
        const nextButton = document.querySelector(selector)?.closest('button');
        if (nextButton && !nextButton.disabled && !nextButton.classList.contains('is-disabled')) {
          console.log(`ETA Exporter: النقر على زر الصفحة التالية`);
          nextButton.click();
          await this.delay(500);
          return true;
        }
      }
      
      console.warn(`ETA Exporter: لم يتم العثور على زر الصفحة التالية`);
      return false;
    } catch (error) {
      console.error('ETA Exporter: خطأ في الانتقال للصفحة التالية:', error);
      return false;
    }
  }
  
  async waitForPageLoadOptimized(expectedPage = null) {
    // انتظار تحميل الصفحة المحسن باستخدام فحوصات جاهزية DOM
    
    console.log(`ETA Exporter: انتظار تحميل الصفحة${expectedPage ? ` ${expectedPage}` : ''}...`);
    
    // 1. انتظار اختفاء مؤشرات التحميل
    await this.waitForCondition(() => {
      const loadingIndicators = document.querySelectorAll('.LoadingIndicator, .ms-Spinner, [class*="loading"], [class*="spinner"], [class*="Loading"]');
      return loadingIndicators.length === 0 || 
             Array.from(loadingIndicators).every(el => el.style.display === 'none' || !el.offsetParent);
    }, 10000);
    
    // 2. انتظار ظهور صفوف الفواتير واستقرارها
    await this.waitForCondition(() => {
      const rows = this.getVisibleInvoiceRows();
      return rows.length > 0;
    }, 10000);
    
    // 3. إذا كان لدينا رقم صفحة متوقع، تحقق من أننا في الصفحة الصحيحة
    if (expectedPage) {
      await this.waitForCondition(() => {
        this.extractPaginationInfo();
        return this.currentPage === expectedPage;
      }, 5000);
    }
    
    // 4. انتظار استقرار DOM (عدم إضافة صفوف جديدة)
    await this.waitForDOMStability(500, 3);
    
    // 5. انتظار إضافي صغير لضمان تحميل جميع البيانات
    await this.delay(500);
    
    console.log(`ETA Exporter: انتهاء انتظار تحميل الصفحة`);
  }
  
  async waitForDOMStability(checkInterval = 500, requiredStableChecks = 3) {
    let stableChecks = 0;
    let lastRowCount = 0;
    
    while (stableChecks < requiredStableChecks) {
      const currentRowCount = this.getVisibleInvoiceRows().length;
      
      if (currentRowCount === lastRowCount && currentRowCount > 0) {
        stableChecks++;
      } else {
        stableChecks = 0;
      }
      
      lastRowCount = currentRowCount;
      await this.delay(checkInterval);
    }
  }
  
  async waitForCondition(condition, timeout = 5000) {
    const startTime = Date.now();
    
    while (Date.now() - startTime < timeout) {
      if (condition()) {
        return true;
      }
      await this.delay(100);
    }
    
    return false;
  }
  
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  
  setProgressCallback(callback) {
    this.progressCallback = callback;
  }
  
  async getInvoiceDetails(invoiceId) {
    try {
      // Extract invoice details from the details view
      const details = await this.extractInvoiceDetailsFromPage(invoiceId);
      
      return {
        success: true,
        data: details
      };
    } catch (error) {
      console.error('Error getting invoice details:', error);
      return { 
        success: false, 
        data: [],
        error: error.message 
      };
    }
  }
  
  async extractInvoiceDetailsFromPage(invoiceId) {
    // This method extracts line items from the invoice details page
    const details = [];
    
    try {
      // Look for the details table in the invoice view
      const detailsTable = document.querySelector('.ms-DetailsList, [data-automationid="DetailsList"]');
      
      if (detailsTable) {
        const rows = detailsTable.querySelectorAll('.ms-DetailsRow[role="row"]');
        
        rows.forEach((row, index) => {
          const cells = row.querySelectorAll('.ms-DetailsRow-cell');
          
          if (cells.length >= 9) { // Based on the 9 columns in the new layout
            const item = {
              description: this.extractCellText(cells[0]) || '',    // A - اسم الصنف
              unitCode: this.extractCellText(cells[1]) || 'EA',     // B - كود الوحدة
              unitName: this.extractCellText(cells[2]) || 'قطعة',   // C - اسم الوحدة
              quantity: this.extractCellText(cells[3]) || '1',      // D - الكمية
              unitPrice: this.extractCellText(cells[4]) || '0',     // E - السعر
              totalValue: this.extractCellText(cells[5]) || '0',    // F - القيمة
              taxAmount: this.extractCellText(cells[6]) || '0',     // G - الضريبة
              vatAmount: this.extractCellText(cells[7]) || '0',     // H - ضريبة القيمة المضافة
              totalWithVat: this.extractCellText(cells[8]) || '0'   // I - الإجمالي
            };
            
            // Only add valid items (skip header rows)
            if (item.description && item.description !== 'اسم الصنف' && item.description.trim() !== '') {
              details.push(item);
            }
          }
        });
      }
      
      // If no detailed items found, create a summary item
      if (details.length === 0) {
        const invoice = this.invoiceData.find(inv => inv.electronicNumber === invoiceId);
        if (invoice) {
          details.push({
            description: 'إجمالي الفاتورة',
            unitCode: 'EA',
            unitName: 'قطعة',
            quantity: '1',
            unitPrice: invoice.totalAmount || '0',
            totalValue: invoice.invoiceValue || invoice.totalAmount || '0',
            taxAmount: '0',
            vatAmount: invoice.vatAmount || '0',
            totalWithVat: invoice.totalAmount || '0'
          });
        }
      }
      
    } catch (error) {
      console.error('Error extracting invoice details:', error);
    }
    
    return details;
  }
  
  extractCellText(cell) {
    if (!cell) return '';
    
    // Try different selectors for cell content
    const textElement = cell.querySelector('.griCellTitle, .griCellTitleGray, .ms-DetailsRow-cellContent') || cell;
    return textElement.textContent?.trim() || '';
  }
  
  getInvoiceData() {
    return {
      invoices: this.invoiceData,
      totalCount: this.totalCount,
      currentPage: this.currentPage,
      totalPages: this.totalPages
    };
  }
  
  cleanup() {
    if (this.observer) {
      this.observer.disconnect();
    }
    if (this.rescanTimeout) {
      clearTimeout(this.rescanTimeout);
    }
  }
}

// Initialize content script
const etaContentScript = new ETAContentScript();

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  switch (request.action) {
    case 'getInvoiceData':
      sendResponse({
        success: true,
        data: etaContentScript.getInvoiceData()
      });
      break;
      
    case 'getInvoiceDetails':
      etaContentScript.getInvoiceDetails(request.invoiceId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true; // Indicates async response
      
    case 'getAllPagesData':
      // Set up progress callback
      if (request.options && request.options.progressCallback) {
        etaContentScript.setProgressCallback((progress) => {
          chrome.runtime.sendMessage({
            action: 'progressUpdate',
            progress: progress
          });
        });
      }
      
      etaContentScript.getAllPagesData(request.options)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true; // Indicates async response
      
    case 'rescanPage':
      etaContentScript.scanForInvoices();
      sendResponse({
        success: true,
        data: etaContentScript.getInvoiceData()
      });
      break;
      
    default:
      sendResponse({ success: false, error: 'Unknown action' });
  }
  
  return true;
});

// Cleanup on page unload
window.addEventListener('beforeunload', () => {
  etaContentScript.cleanup();
});