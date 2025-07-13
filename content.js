// Enhanced Content script for ETA Invoice Exporter with improved performance and accurate data extraction
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
      
      // Find invoice rows using improved selectors
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
    const allRows = document.querySelectorAll('.ms-DetailsRow[role="row"]');
    const visibleRows = [];
    
    allRows.forEach(row => {
      if (this.isRowVisible(row) && this.hasInvoiceData(row)) {
        visibleRows.push(row);
      }
    });
    
    return visibleRows;
  }
  
  isRowVisible(row) {
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
    const cells = row.querySelectorAll('.ms-DetailsRow-cell');
    if (cells.length === 0) return false;
    
    const firstCell = cells[0];
    const electronicLink = firstCell?.querySelector('.internalId-link a.griCellTitle');
    const internalNumber = firstCell?.querySelector('.griCellSubTitle');
    
    return !!(electronicLink?.textContent?.trim() || internalNumber?.textContent?.trim());
  }
  
  extractPaginationInfo() {
    try {
      // Extract total count from pagination
      const totalLabel = document.querySelector('.eta-pagination-totalrecordCount-label, [class*="pagination"] [class*="total"], [class*="record"] [class*="count"]');
      if (totalLabel) {
        const match = totalLabel.textContent.match(/النتائج:\s*(\d+)|(\d+)\s*نتيجة|Total:\s*(\d+)/);
        if (match) {
          this.totalCount = parseInt(match[1] || match[2] || match[3]);
        }
      }
      
      // Extract current page
      const currentPageBtn = document.querySelector('.eta-pageNumber.is-checked, [class*="page"][class*="current"], [class*="active"][class*="page"]');
      if (currentPageBtn) {
        const pageLabel = currentPageBtn.querySelector('.ms-Button-label, [class*="label"], [class*="text"]');
        if (pageLabel) {
          this.currentPage = parseInt(pageLabel.textContent) || 1;
        }
      }
      
      // Calculate total pages
      const visibleRows = this.getVisibleInvoiceRows();
      const itemsPerPage = Math.max(visibleRows.length, 10);
      this.totalPages = Math.ceil(this.totalCount / itemsPerPage);
      
      console.log(`ETA Exporter: Page ${this.currentPage} of ${this.totalPages}, Total: ${this.totalCount} invoices`);
      
    } catch (error) {
      console.warn('ETA Exporter: Error extracting pagination info:', error);
    }
  }
  
  extractDataFromRow(row, index) {
    const invoice = {
      index: index,
      pageNumber: this.currentPage,
      
      // Main invoice data matching the Excel format exactly
      serialNumber: index,
      viewButton: 'عرض',
      documentType: '',
      documentVersion: '',
      status: '',
      issueDate: '',
      submissionDate: '',
      invoiceCurrency: 'EGP',
      invoiceValue: '',
      vatAmount: '',
      taxDiscount: '0',
      totalInvoice: '',
      internalNumber: '',
      electronicNumber: '',
      sellerTaxNumber: '',
      sellerName: '',
      sellerAddress: '',
      buyerTaxNumber: '',
      buyerName: '',
      buyerAddress: '',
      purchaseOrderRef: '',
      purchaseOrderDesc: '',
      salesOrderRef: '',
      electronicSignature: 'موقع إلكترونياً',
      foodDrugGuide: '',
      externalLink: '',
      
      // Additional fields for compatibility
      issueTime: '',
      totalAmount: '',
      currency: 'EGP',
      submissionId: '',
      details: []
    };
    
    try {
      const cells = row.querySelectorAll('.ms-DetailsRow-cell');
      
      if (cells.length === 0) {
        console.warn(`No cells found in row ${index}`);
        return invoice;
      }
      
      // Cell 0: Electronic Number and Internal Number (data-automation-key="uuid")
      const uuidCell = cells[0];
      if (uuidCell && uuidCell.getAttribute('data-automation-key') === 'uuid') {
        const electronicLink = uuidCell.querySelector('.internalId-link a.griCellTitle');
        if (electronicLink) {
          invoice.electronicNumber = electronicLink.textContent?.trim() || '';
        }
        
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
          invoice.submissionDate = invoice.issueDate; // Default to same date
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
          const totalText = totalElement.textContent?.trim() || '';
          invoice.totalAmount = totalText;
          invoice.totalInvoice = totalText;
          
          // Calculate VAT and invoice value (assuming 14% VAT rate)
          const totalValue = this.parseAmount(totalText);
          if (totalValue > 0) {
            const vatRate = 0.14;
            invoice.vatAmount = this.formatAmount((totalValue * vatRate) / (1 + vatRate));
            invoice.invoiceValue = this.formatAmount(totalValue - this.parseAmount(invoice.vatAmount));
          }
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
        
        // Set default address
        if (invoice.sellerName && !invoice.sellerAddress) {
          invoice.sellerAddress = 'غير محدد';
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
        
        // Set default address
        if (invoice.buyerName && !invoice.buyerAddress) {
          invoice.buyerAddress = 'غير محدد';
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
        const validRejectedDiv = statusCell.querySelector('.horizontal.valid-rejected');
        if (validRejectedDiv) {
          const validStatus = validRejectedDiv.querySelector('.status-Valid');
          const rejectedStatus = validRejectedDiv.querySelector('.status-Rejected');
          if (validStatus && rejectedStatus) {
            invoice.status = `${validStatus.textContent?.trim()} → ${rejectedStatus.textContent?.trim()}`;
          }
        } else {
          const textStatus = statusCell.querySelector('.textStatus');
          if (textStatus) {
            invoice.status = textStatus.textContent?.trim() || '';
          } else {
            invoice.status = statusCell.textContent?.trim() || '';
          }
        }
      }
      
      // Generate external link
      if (invoice.electronicNumber) {
        invoice.externalLink = this.generateExternalLink(invoice);
      }
      
    } catch (error) {
      console.warn(`ETA Exporter: Error extracting data from row ${index}:`, error);
    }
    
    return invoice;
  }
  
  parseAmount(amountText) {
    if (!amountText) return 0;
    const cleanText = amountText.replace(/[,٬\s]/g, '').replace(/[^\d.]/g, '');
    return parseFloat(cleanText) || 0;
  }
  
  formatAmount(amount) {
    if (!amount || amount === 0) return '0';
    return amount.toLocaleString('en-US', { 
      minimumFractionDigits: 2, 
      maximumFractionDigits: 2 
    });
  }
  
  generateExternalLink(invoice) {
    if (!invoice.electronicNumber) return '';
    
    let shareId = '';
    if (invoice.submissionId && invoice.submissionId.length > 10) {
      shareId = invoice.submissionId;
    } else {
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
      
      console.log(`ETA Exporter: Starting to load all pages. Current: ${this.currentPage}, Total: ${this.totalPages}`);
      
      // First scan current page
      this.scanForInvoices();
      
      if (this.totalPages <= 1) {
        this.allPagesData = [...this.invoiceData];
        console.log(`ETA Exporter: Only one page, collected ${this.allPagesData.length} invoices`);
        return {
          success: true,
          data: this.allPagesData,
          totalProcessed: this.allPagesData.length
        };
      }
      
      // Collect data from all pages using optimized approach
      const allPagePromises = [];
      
      // Process pages in parallel batches for better performance
      const batchSize = 3;
      for (let startPage = 1; startPage <= this.totalPages; startPage += batchSize) {
        const endPage = Math.min(startPage + batchSize - 1, this.totalPages);
        const batchPromise = this.processPagesInBatch(startPage, endPage);
        allPagePromises.push(batchPromise);
      }
      
      // Wait for all batches to complete
      const batchResults = await Promise.all(allPagePromises);
      
      // Flatten results
      batchResults.forEach(batchData => {
        this.allPagesData.push(...batchData);
      });
      
      console.log(`ETA Exporter: Completed loading all pages. Total invoices: ${this.allPagesData.length}`);
      
      return {
        success: true,
        data: this.allPagesData,
        totalProcessed: this.allPagesData.length
      };
      
    } catch (error) {
      console.error('ETA Exporter: Error getting all pages data:', error);
      return { 
        success: false, 
        data: this.allPagesData,
        error: error.message 
      };
    } finally {
      this.isProcessingAllPages = false;
    }
  }
  
  async processPagesInBatch(startPage, endPage) {
    const batchData = [];
    
    for (let page = startPage; page <= endPage; page++) {
      try {
        if (this.progressCallback) {
          this.progressCallback({
            currentPage: page,
            totalPages: this.totalPages,
            message: `جاري معالجة الصفحة ${page} من ${this.totalPages}...`,
            percentage: (page / this.totalPages) * 100
          });
        }
        
        // Navigate to page if not current
        if (page !== this.currentPage) {
          const navigated = await this.navigateToPage(page);
          if (!navigated) {
            console.warn(`Failed to navigate to page ${page}`);
            continue;
          }
          
          await this.waitForPageLoad(page);
        }
        
        // Scan invoices on this page
        this.scanForInvoices();
        
        if (this.invoiceData.length > 0) {
          batchData.push(...this.invoiceData);
          console.log(`ETA Exporter: Processed page ${page}, collected ${this.invoiceData.length} invoices`);
        }
        
        // Small delay between pages
        await this.delay(300);
        
      } catch (error) {
        console.error(`Error processing page ${page}:`, error);
      }
    }
    
    return batchData;
  }
  
  async navigateToPage(pageNumber) {
    try {
      if (this.currentPage === pageNumber) {
        return true;
      }
      
      // Find page button
      const pageButtons = document.querySelectorAll('.eta-pageNumber, [class*="pageNumber"], .ms-Button[aria-label*="Page"]');
      
      for (const btn of pageButtons) {
        const label = btn.querySelector('.ms-Button-label, [class*="label"]');
        const buttonText = label ? label.textContent : btn.textContent;
        
        if (buttonText && parseInt(buttonText.trim()) === pageNumber) {
          btn.click();
          await this.delay(500);
          return true;
        }
      }
      
      // Try navigation with next/previous buttons
      if (pageNumber > this.currentPage) {
        for (let i = this.currentPage; i < pageNumber; i++) {
          const success = await this.navigateToNextPage();
          if (!success) break;
          await this.delay(300);
        }
      } else if (pageNumber < this.currentPage) {
        for (let i = this.currentPage; i > pageNumber; i--) {
          const success = await this.navigateToPreviousPage();
          if (!success) break;
          await this.delay(300);
        }
      }
      
      return this.currentPage === pageNumber;
    } catch (error) {
      console.error(`Error navigating to page ${pageNumber}:`, error);
      return false;
    }
  }
  
  async navigateToNextPage() {
    const nextSelectors = [
      '[data-icon-name="ChevronRight"]',
      '[data-icon-name="Next"]',
      '[aria-label*="Next"]',
      '[aria-label*="التالي"]'
    ];
    
    for (const selector of nextSelectors) {
      const nextButton = document.querySelector(selector)?.closest('button');
      if (nextButton && !nextButton.disabled) {
        nextButton.click();
        await this.delay(300);
        return true;
      }
    }
    return false;
  }
  
  async navigateToPreviousPage() {
    const prevSelectors = [
      '[data-icon-name="ChevronLeft"]',
      '[data-icon-name="Previous"]',
      '[aria-label*="Previous"]',
      '[aria-label*="السابق"]'
    ];
    
    for (const selector of prevSelectors) {
      const prevButton = document.querySelector(selector)?.closest('button');
      if (prevButton && !prevButton.disabled) {
        prevButton.click();
        await this.delay(300);
        return true;
      }
    }
    return false;
  }
  
  async waitForPageLoad(expectedPage = null) {
    // Wait for loading indicators to disappear
    await this.waitForCondition(() => {
      const loadingIndicators = document.querySelectorAll('.LoadingIndicator, .ms-Spinner, [class*="loading"]');
      return loadingIndicators.length === 0 || 
             Array.from(loadingIndicators).every(el => !el.offsetParent);
    }, 5000);
    
    // Wait for invoice rows to appear
    await this.waitForCondition(() => {
      const rows = this.getVisibleInvoiceRows();
      return rows.length > 0;
    }, 5000);
    
    // Wait for page number to update if expected
    if (expectedPage) {
      await this.waitForCondition(() => {
        this.extractPaginationInfo();
        return this.currentPage === expectedPage;
      }, 3000);
    }
    
    // Wait for DOM stability
    await this.delay(300);
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
    const details = [];
    
    try {
      const detailsTable = document.querySelector('.ms-DetailsList, [data-automationid="DetailsList"]');
      
      if (detailsTable) {
        const rows = detailsTable.querySelectorAll('.ms-DetailsRow[role="row"]');
        
        rows.forEach((row, index) => {
          const cells = row.querySelectorAll('.ms-DetailsRow-cell');
          
          if (cells.length >= 9) {
            const item = {
              itemCode: this.extractCellText(cells[0]) || '',
              description: this.extractCellText(cells[1]) || '',
              unitCode: this.extractCellText(cells[2]) || 'EA',
              unitName: this.extractCellText(cells[3]) || 'قطعة',
              quantity: this.extractCellText(cells[4]) || '1',
              unitPrice: this.extractCellText(cells[5]) || '0',
              totalValue: this.extractCellText(cells[6]) || '0',
              taxAmount: this.extractCellText(cells[7]) || '0',
              vatAmount: this.extractCellText(cells[8]) || '0'
            };
            
            if (item.description && item.description !== 'اسم الصنف' && item.description.trim() !== '') {
              details.push(item);
            }
          }
        });
      }
      
      if (details.length === 0) {
        const invoice = this.invoiceData.find(inv => inv.electronicNumber === invoiceId);
        if (invoice) {
          details.push({
            itemCode: invoice.electronicNumber,
            description: 'إجمالي الفاتورة',
            unitCode: 'EA',
            unitName: 'فاتورة',
            quantity: '1',
            unitPrice: invoice.totalAmount || '0',
            totalValue: invoice.invoiceValue || invoice.totalAmount || '0',
            taxAmount: '0',
            vatAmount: invoice.vatAmount || '0'
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
      return true;
      
    case 'getAllPagesData':
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
      return true;
      
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