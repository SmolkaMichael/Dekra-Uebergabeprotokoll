// DEKRA Gutachten-App - Professionelles Layout

// App State
const appState = {
    formData: {},
    attachedFiles: [],
    templates: [],
    currentZoom: 100
};

// Wait for DOM to be fully loaded
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initAll);
} else {
    initAll();
}

function initAll() {
    console.log('Initializing DEKRA App...');
    
    // Initialize components
    setupEventListeners();
    setDefaultDates();
    initializeAutoSave();
    
    // Initial preview update
    updateLivePreview();
    
    console.log('DEKRA App initialized successfully');
}

// Setup Event Listeners
function setupEventListeners() {
    // Form input listeners
    const form = document.getElementById('gutachtenForm');
    if (form) {
        const inputs = form.querySelectorAll('input, textarea, select');
        
        inputs.forEach(input => {
            input.addEventListener('input', handleFormInput);
            input.addEventListener('change', handleFormInput);
        });
    }

    // Button listeners - check if elements exist before adding listeners
    
    const exportPdfBtn = document.getElementById('exportPdf');
    if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportToPDF);
    
    const exportWordBtn = document.getElementById('exportWord');
    if (exportWordBtn) exportWordBtn.addEventListener('click', exportToWord);
    
    // Template export button listener
    const exportTemplateBtn = document.getElementById('exportTemplate');
    if (exportTemplateBtn) exportTemplateBtn.addEventListener('click', exportFromTemplate);
    
    // Word merge button listener
    const mergeWordBtn = document.getElementById('mergeWord');
    if (mergeWordBtn) mergeWordBtn.addEventListener('click', () => {
        window.open('https://products.groupdocs.app/merger/word', '_blank');
    });
    
    // File upload listener
    const fileUpload = document.getElementById('fileUpload');
    if (fileUpload) fileUpload.addEventListener('change', handleFileUpload);
    
    
    
    // Keyboard shortcuts
    document.addEventListener('keydown', handleKeyboardShortcuts);
}

// Handle form input changes
function handleFormInput(event) {
    const field = event.target;
    const fieldName = field.name;
    const fieldValue = field.value;
    
    // Update app state
    appState.formData[fieldName] = fieldValue;
    
    // Update live preview immediately
    updateLivePreview();
    
    // Save to localStorage
    saveToLocalStorage();
}

// Update Live Preview with MATTHIAS TEMPLATE using Microsoft Office Online
async function updateLivePreview() {
    const previewContent = document.getElementById('previewContent');
    if (!previewContent) {
        console.error('Preview content element not found');
        return;
    }
    
    try {
        const formData = collectFormData();
        
        // Use Microsoft Office Online Viewer for EXACT Word display
        // This shows the document EXACTLY like in Microsoft Word!
        const templateUrl = 'https://raw.githubusercontent.com/SmolkaMichael/Dekra-Uebergabeprotokoll/main/MatthiasVorlage.docx';
        const viewerUrl = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(templateUrl)}`;
        
        previewContent.innerHTML = `
            <iframe 
                src="${viewerUrl}"
                width="100%" 
                height="100%" 
                frameborder="0"
                style="min-height: 800px; border: none;">
                Dies ist eine Office Online-Vorschau.
            </iframe>
        `;
        
        
    } catch (error) {
        console.error('Error updating preview with template:', error);
        // Pass formData to the fallback function
        const formData = collectFormData();
        updateHtmlPreview(formData);
    }
}

// Fallback HTML preview (old method)
function updateHtmlPreview(formData) {
    const previewContent = document.getElementById('previewContent');
    if (!previewContent) return;
    
    try {
        // Apply the exact DEKRA template styles with proper padding
        previewContent.style.cssText = `
            font-family: Arial, sans-serif;
            font-size: 16pt;
            line-height: 1.8;
            background: white;
            position: relative;
            width: 100%;
            min-height: 400mm;
            padding: 30mm 25mm 40mm 30mm;
            display: flex;
            flex-direction: column;
            box-sizing: border-box;
        `;
        
        // Create the exact professional layout
        previewContent.innerHTML = `
        <style>
            .preview-page * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            
            .preview-page {
                font-family: Arial, sans-serif;
                font-size: 16pt;
                line-height: 1.6;
            }
            
            /* KOPFZEILE */
            .header {
                text-align: left;
                padding-bottom: 10px;
                border-bottom: 2px solid black;
                margin-bottom: 15px;
                position: relative;
            }
            
            .header .company-name {
                font-weight: bold;
                font-size: 22pt;
                margin-bottom: 5px;
            }
            
            .header .department {
                font-size: 18pt;
                margin-bottom: 3px;
            }
            
            .header .contact {
                font-size: 16pt;
                color: #333;
                display: flex;
                justify-content: space-between;
                align-items: center;
            }
            
            .page-indicator {
                font-size: 16pt;
                margin-left: auto;
            }
            
            /* HAUPTINHALT */
            .content {
                flex: 1;
                margin-top: 30px;
            }
            
            .recipient-info {
                display: flex;
                justify-content: space-between;
                margin-bottom: 40px;
            }
            
            .recipient-address {
                font-size: 16pt;
                line-height: 1.5;
                max-width: 300px;
            }
            
            .document-info {
                text-align: left;
                font-size: 16pt;
                line-height: 1.8;
                display: grid;
                grid-template-columns: auto auto;
                gap: 15px;
                margin-left: auto;
                white-space: nowrap;
            }
            
            .document-info .label {
                font-weight: normal;
                padding-right: 15px;
                text-align: left;
            }
            
            .document-info .value {
                font-weight: normal;
                white-space: nowrap;
                text-align: left;
            }
            
            .document-info .gutachten-nr {
                font-weight: bold;
            }
            
            .gutachten-title {
                text-align: center;
                font-size: 24pt;
                font-weight: bold;
                letter-spacing: 15px;
                margin: 50px 0;
                padding: 15px 0;
            }
            
            .gutachten-details {
                margin-top: 30px;
            }
            
            .detail-row {
                display: grid;
                grid-template-columns: 180px 20px 1fr;
                margin-bottom: 12px;
                align-items: baseline;
            }
            
            .detail-label {
                font-weight: normal;
                padding-right: 10px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            
            .detail-colon {
                text-align: center;
            }
            
            .detail-value {
                padding-left: 10px;
                word-wrap: break-word;
            }
            
            /* FUSSZEILE */
            .footer {
                margin-top: auto;
                border-top: 1px solid #666;
                padding-top: 10px;
                font-size: 7pt;
                line-height: 1.3;
                color: #444;
            }
            
            .footer-content {
                display: grid;
                grid-template-columns: 1fr 1.2fr 1fr;
                gap: 30px;
            }
            
            .footer-column h4 {
                font-size: 7.5pt;
                margin-bottom: 3px;
                font-weight: bold;
            }
            
            .footer-column {
                font-size: 6.5pt;
            }
            
            .footer-column p {
                margin: 2px 0;
            }
        </style>
        
        <!-- KOPFZEILE -->
        <header class="header">
            <div class="company-name">DEKRA Automobil GmbH</div>
            <div class="department">Niederlassung Düsseldorf Fachbereich Polizei- und Gerichtsgutachten</div>
            <div class="contact">
                <span>Höherweg 111, 40233 Düsseldorf, Telefon 0211/2300-0, Fax 0211/2300-222</span>
                <span class="page-indicator">Blatt 1</span>
            </div>
        </header>
        
        <!-- HAUPTINHALT -->
        <main class="content">
            <div class="recipient-info">
                <div class="recipient-address">
                    ${formData.empfaengerName || 'KPB Neuss'}<br>
                    ${formData.empfaengerStrasse || 'Holbeinstraße 4'}<br>
                    ${formData.empfaengerPLZ || '40667'} ${formData.empfaengerOrt || 'Meerbusch'}
                </div>
                <div class="document-info">
                    <span class="label">Gutachten-Nummer</span>
                    <span class="value gutachten-nr">${formData.gutachtenNummer || '302/53884 25-4322909943'}</span>
                    
                    <span class="label">vom</span>
                    <span class="value">${formatDate(formData.datum) || '02.09.2025'}</span>
                    
                    <span class="label">Kunden-Nummer</span>
                    <span class="value">${formData.kundenNummer || '212500140'}</span>
                </div>
            </div>
            
            <h1 class="gutachten-title">GUTACHTEN</h1>
            
            <div class="gutachten-details">
                <div class="detail-row">
                    <span class="detail-label">Aktenzeichen</span>
                    <span class="detail-colon">:</span>
                    <span class="detail-value">${formData.aktenzeichen || ''}</span>
                </div>
                
                <div class="detail-row">
                    <span class="detail-label">Beteiligte / Sache</span>
                    <span class="detail-colon">:</span>
                    <span class="detail-value">${formData.beteiligte || ''}</span>
                </div>
                
                <div class="detail-row">
                    <span class="detail-label">Auftraggeber</span>
                    <span class="detail-colon">:</span>
                    <span class="detail-value">${formData.auftraggeber || ''}</span>
                </div>
                
                <div class="detail-row">
                    <span class="detail-label">Auftrag vom</span>
                    <span class="detail-colon">:</span>
                    <span class="detail-value">${formatDate(formData.auftragVom) ? formatDate(formData.auftragVom) + ' (Akteneingang)' : ''}</span>
                </div>
                
                <div class="detail-row">
                    <span class="detail-label">Besichtigungsdatum</span>
                    <span class="detail-colon">:</span>
                    <span class="detail-value">${formData.besichtigungsdatum ? formatDateLong(formData.besichtigungsdatum) : ''}</span>
                </div>
                
                <div class="detail-row">
                    <span class="detail-label">Besichtigungsort</span>
                    <span class="detail-colon">:</span>
                    <span class="detail-value">${formData.besichtigungsort || ''}</span>
                </div>
                
                <div class="detail-row">
                    <span class="detail-label">Sachverständiger</span>
                    <span class="detail-colon">:</span>
                    <span class="detail-value">${formData.sachverstaendiger || ''}</span>
                </div>
            </div>
        </main>
        
        <!-- FUSSZEILE -->
        <footer class="footer">
            <div class="footer-content">
                <div class="footer-column">
                    <h4>DEKRA Automobil GmbH</h4>
                    <p>Handwerkstraße 15</p>
                    <p>D-70565 Stuttgart</p>
                    <p>Telefon (07 11) 78 61-0</p>
                    <p>Telefax (07 11) 78 61-22 40</p>
                    <p>www.dekra.com</p>
                </div>
                
                <div class="footer-column">
                    <p>Sitz Stuttgart, Amtsgericht Stuttgart,</p>
                    <p>HRB-Nr. 21039</p>
                    <br>
                    <h4>Bankverbindungen:</h4>
                    <p>Commerzbank AG</p>
                    <p>IBAN DE84 6008 0000 0901 0051 00 / BIC DRESDEFF600</p>
                    <p>BW-Bank</p>
                    <p>IBAN DE74 6005 0101 0002 0195 25 / BIC SOLADEST</p>
                </div>
                
                <div class="footer-column">
                    <h4>Vorsitzender des Aufsichtsrates:</h4>
                    <p>Stefan Kölbl</p>
                    <br>
                    <h4>Geschäftsführer:</h4>
                    <p>Guido Kutschera (Vorsitzender)</p>
                    <p>Friedemann Bausch</p>
                    <p>Jann Fehlauer</p>
                </div>
            </div>
        </footer>
    `;
        
        const pdfPreview = document.getElementById('pdfPreview');
        const oldAttachments = pdfPreview.querySelectorAll('.attachment-page');
        oldAttachments.forEach(page => page.remove());
        
        // Add attachment pages to preview
        if (appState.attachedFiles.length > 0) {
            appState.attachedFiles.forEach((file, index) => {
                const attachmentPage = document.createElement('div');
                attachmentPage.className = 'preview-page attachment-page';
                attachmentPage.style.cssText = `
                    background: white;
                    width: 210mm;
                    min-height: 400mm;
                    padding: 30mm 25mm 40mm 30mm;
                    box-shadow: 0 0 20px rgba(0, 0, 0, 0.5);
                    transform-origin: top center;
                    font-family: Arial, sans-serif;
                    font-size: 16pt;
                    line-height: 1.8;
                    color: black;
                    position: relative;
                    display: flex;
                    flex-direction: column;
                    box-sizing: border-box;
                    transform: scale(0.6);
                    margin-bottom: 50px;
                `;
                
                if (file.type === 'application/pdf' || file.name.endsWith('.pdf')) {
                    attachmentPage.innerHTML = `
                        <h2 style="text-align: center; font-size: 24pt; margin-bottom: 30px;">Anhang ${index + 1}</h2>
                        <div style="text-align: center; padding: 50px; background: #f0f0f0; border: 2px dashed #999;">
                            <p style="font-size: 18pt; font-weight: bold; margin-bottom: 20px;">📄 PDF-Dokument</p>
                            <p style="font-size: 16pt;">${file.name}</p>
                            <p style="font-size: 14pt; color: #666; margin-top: 10px;">Größe: ${formatFileSize(file.size)}</p>
                            <p style="font-size: 12pt; color: #999; margin-top: 20px;">Wird beim Export automatisch eingefügt</p>
                        </div>
                    `;
                } else if (file.type.startsWith('image/')) {
                    // For images, create a preview
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        attachmentPage.innerHTML = `
                            <h2 style="text-align: center; font-size: 24pt; margin-bottom: 30px;">Anhang ${index + 1}</h2>
                            <div style="text-align: center; flex: 1; display: flex; flex-direction: column; justify-content: center;">
                                <img src="${e.target.result}" style="max-width: 100%; max-height: 600px; object-fit: contain; margin: 0 auto;">
                                <p style="font-size: 14pt; margin-top: 20px;">${file.name}</p>
                                <p style="font-size: 12pt; color: #666;">Größe: ${formatFileSize(file.size)}</p>
                            </div>
                        `;
                    };
                    reader.readAsDataURL(file);
                } else if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
                           file.type === 'application/msword' || 
                           file.name.endsWith('.docx') || 
                           file.name.endsWith('.doc')) {
                    attachmentPage.innerHTML = `
                        <h2 style="text-align: center; font-size: 24pt; margin-bottom: 30px;">Anhang ${index + 1}</h2>
                        <div style="text-align: center; padding: 50px; background: #e3f2fd; border: 2px dashed #2196F3;">
                            <p style="font-size: 18pt; font-weight: bold; margin-bottom: 20px;">📝 Word-Dokument</p>
                            <p style="font-size: 16pt;">${file.name}</p>
                            <p style="font-size: 14pt; color: #666; margin-top: 10px;">Größe: ${formatFileSize(file.size)}</p>
                            <p style="font-size: 12pt; color: #1976D2; margin-top: 20px;">Wird beim Word-Export angehängt</p>
                        </div>
                    `;
                } else {
                    attachmentPage.innerHTML = `
                        <h2 style="text-align: center; font-size: 24pt; margin-bottom: 30px;">Anhang ${index + 1}</h2>
                        <div style="text-align: center; padding: 50px; background: #f5f5f5; border: 2px dashed #999;">
                            <p style="font-size: 18pt; font-weight: bold; margin-bottom: 20px;">📎 Dokument</p>
                            <p style="font-size: 16pt;">${file.name}</p>
                            <p style="font-size: 14pt; color: #666; margin-top: 10px;">Typ: ${file.type || 'Unbekannt'}</p>
                            <p style="font-size: 14pt; color: #666;">Größe: ${formatFileSize(file.size)}</p>
                        </div>
                    `;
                }
                
                pdfPreview.appendChild(attachmentPage);
            });
        }
    } catch (error) {
        console.error('Error in updateHtmlPreview:', error);
    }
}

// Export using MatthiasVorlage.docx template
async function exportFromTemplate() {
    showProgress('Erstelle Word mit Matthias Vorlage...');
    
    try {
        const formData = collectFormData();
        
        // Load the template file
        const templateResponse = await fetch('./MatthiasVorlage.docx');
        if (!templateResponse.ok) {
            throw new Error('MatthiasVorlage.docx konnte nicht geladen werden');
        }
        
        const templateArrayBuffer = await templateResponse.arrayBuffer();
        const zip = new PizZip(templateArrayBuffer);
        const doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });
        
        // Prepare data for template (convert to template format)
        const templateData = {
            gutachtenNummer: formData.gutachtenNummer || '',
            datum: formatDateLong(formData.datum),
            auftraggeber: formData.auftraggeber || '',
            kundenNummer: formData.kundenNummer || '',
            empfaengerName: formData.empfaengerName || '',
            empfaengerStrasse: formData.empfaengerStrasse || '',
            empfaengerPLZ: formData.empfaengerPLZ || '',
            empfaengerOrt: formData.empfaengerOrt || '',
            aktenzeichen: formData.aktenzeichen || '',
            beteiligte: formData.beteiligte || '',
            auftragVom: formatDate(formData.auftragVom),
            besichtigungsdatum: formatDate(formData.besichtigungsdatum),
            besichtigungsort: formData.besichtigungsort || '',
            sachverstaendiger: formData.sachverstaendiger || ''
        };
        
        console.log('Template data:', templateData);
        
        // Fill template with data
        doc.setData(templateData);
        
        try {
            doc.render();
        } catch (error) {
            console.error('Template render error:', error);
            // If template has no placeholders, continue anyway
            alert('Hinweis: Template hat möglicherweise keine Platzhalter. Dokument wird trotzdem erstellt.');
        }
        
        // Generate and download
        const buf = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        
        const url = URL.createObjectURL(buf);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Matthias_Gutachten_${formData.gutachtenNummer || 'draft'}.docx`;
        a.click();
        URL.revokeObjectURL(url);
        
    } catch (error) {
        console.error('Fehler beim Template-Export:', error);
        alert(`Fehler beim Erstellen mit Matthias Vorlage: ${error.message}`);
    } finally {
        hideProgress();
    }
}

// Collect form data
function collectFormData() {
    const data = {};
    const form = document.getElementById('gutachtenForm');
    
    if (!form) return data;
    
    const formData = new FormData(form);
    for (let [key, value] of formData.entries()) {
        data[key] = value;
    }
    
    return data;
}

// Format date to German format
function formatDate(dateString) {
    if (!dateString) return '';
    try {
        const date = new Date(dateString);
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}.${month}.${year}`;
    } catch (error) {
        return dateString;
    }
}

// Format date to long German format with weekday
function formatDateLong(dateString) {
    if (!dateString) return '';
    try {
        const date = new Date(dateString);
        const weekdays = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        const weekday = weekdays[date.getDay()];
        return `${weekday}, den ${day}.${month}.${year}`;
    } catch (error) {
        return dateString;
    }
}

// Format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// Handle file upload
function handleFileUpload(event) {
    const files = Array.from(event.target.files);
    const fileList = document.getElementById('fileList');
    
    files.forEach(file => {
        // Check file size (50 MB limit)
        if (file.size > 50 * 1024 * 1024) {
            alert(`Die Datei ${file.name} ist zu groß. Maximale Größe: 50 MB`);
            return;
        }
        
        // Check file type (erweitert für mehr Dateitypen)
        const allowedTypes = [
            'application/pdf',
            'application/msword', 
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/vnd.ms-word',
            'text/plain',
            'image/jpeg',
            'image/png'
        ];
        if (!allowedTypes.includes(file.type) && !file.name.endsWith('.docx') && !file.name.endsWith('.doc')) {
            alert(`Dateityp nicht unterstützt: ${file.name}`);
            return;
        }
        
        // Add to state
        appState.attachedFiles.push(file);
        
        // Create file item element
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <span class="file-name">${file.name}</span>
            <span class="file-size">${formatFileSize(file.size)}</span>
            <button class="remove-file" data-filename="${file.name}">Entfernen</button>
        `;
        
        fileList.appendChild(fileItem);
        
        // Add remove listener
        fileItem.querySelector('.remove-file').addEventListener('click', function() {
            removeFile(file.name);
            fileItem.remove();
        });
    });
    
    updateLivePreview();
}

// Remove file
function removeFile(filename) {
    appState.attachedFiles = appState.attachedFiles.filter(f => f.name !== filename);
    updateLivePreview();
}


// Export to PDF
async function exportToPDF() {
    showProgress('Exportiere PDF...');
    
    try {
        const formData = collectFormData();
        const { PDFDocument, rgb, StandardFonts } = PDFLib;
        
        // Create new PDF document
        const pdfDoc = await PDFDocument.create();
        
        // Add first page with form data (A4 size)
        const page = pdfDoc.addPage([595, 842]); // A4 in points
        const { width, height } = page.getSize();
        
        // Embed fonts
        const helvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const helveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        
        // Draw header
        let yPos = height - 50;
        
        // Company name
        page.drawText('DEKRA Automobil GmbH', {
            x: 50,
            y: yPos,
            size: 14,
            font: helveticaBold
        });
        yPos -= 20;
        
        // Department
        page.drawText('Niederlassung Düsseldorf Fachbereich Polizei- und Gerichtsgutachten', {
            x: 50,
            y: yPos,
            size: 11,
            font: helvetica
        });
        yPos -= 15;
        
        // Contact with Blatt 1
        page.drawText('Höherweg 111, 40233 Düsseldorf, Telefon 0211/2300-0, Fax 0211/2300-222', {
            x: 50,
            y: yPos,
            size: 10,
            font: helvetica,
            color: rgb(0.2, 0.2, 0.2)
        });
        
        // Blatt 1 on same line as Höherweg
        page.drawText('Blatt 1', {
            x: width - 100,
            y: yPos,
            size: 10,
            font: helvetica
        });
        yPos -= 10;
        
        // Draw header line
        page.drawLine({
            start: { x: 50, y: yPos },
            end: { x: width - 50, y: yPos },
            thickness: 2,
            color: rgb(0, 0, 0)
        });
        yPos -= 40;
        
        // Set same Y position for address and document info
        const infoStartY = yPos - 40;
        
        // Recipient address (left side)
        const addressLines = [
            formData.empfaengerName || 'KPB Neuss',
            formData.empfaengerStrasse || 'Holbeinstraße 4',
            `${formData.empfaengerPLZ || '40667'} ${formData.empfaengerOrt || 'Meerbusch'}`
        ];
        
        let addressY = infoStartY;
        for (const line of addressLines) {
            page.drawText(line, {
                x: 50,
                y: addressY,
                size: 10,
                font: helvetica
            });
            addressY -= 15;
        }
        
        // Document info (right side) - aligned with address
        let infoY = infoStartY;
        page.drawText('Gutachten-Nummer', { x: 320, y: infoY, size: 10, font: helvetica });
        page.drawText(formData.gutachtenNummer || '302/53884 25-4322909943', { 
            x: 450, y: infoY, size: 10, font: helveticaBold 
        });
        infoY -= 15;
        
        page.drawText('vom', { x: 320, y: infoY, size: 10, font: helvetica });
        page.drawText(formatDate(formData.datum) || '02.09.2025', { 
            x: 450, y: infoY, size: 10, font: helvetica 
        });
        infoY -= 15;
        
        page.drawText('Kunden-Nummer', { x: 320, y: infoY, size: 10, font: helvetica });
        page.drawText(formData.kundenNummer || '212500140', { 
            x: 450, y: infoY, size: 10, font: helvetica 
        });
        
        yPos = addressY - 20;
        
        // Title
        yPos = height - 300;
        const titleText = 'GUTACHTEN';
        const titleWidth = helveticaBold.widthOfTextAtSize(titleText, 16);
        page.drawText(titleText, {
            x: (width - titleWidth) / 2,
            y: yPos,
            size: 16,
            font: helveticaBold
        });
        yPos -= 60;
        
        // Details
        const details = [
            { label: 'Aktenzeichen', value: formData.aktenzeichen },
            { label: 'Beteiligte / Sache', value: formData.beteiligte },
            { label: 'Auftraggeber', value: formData.auftraggeber },
            { label: 'Auftrag vom', value: formatDate(formData.auftragVom) ? formatDate(formData.auftragVom) + ' (Akteneingang)' : '' },
            { label: 'Besichtigungsdatum', value: formData.besichtigungsdatum ? formatDateLong(formData.besichtigungsdatum) : '' },
            { label: 'Besichtigungsort', value: formData.besichtigungsort },
            { label: 'Sachverständiger', value: formData.sachverstaendiger }
        ];
        
        for (const detail of details) {
            page.drawText(detail.label, {
                x: 50,
                y: yPos,
                size: 11,
                font: helvetica
            });
            
            page.drawText(':', {
                x: 200,
                y: yPos,
                size: 11,
                font: helvetica
            });
            
            if (detail.value) {
                page.drawText(detail.value, {
                    x: 220,
                    y: yPos,
                    size: 11,
                    font: helvetica
                });
            }
            
            yPos -= 25;
        }
        
        // Footer - Complete footer with all information
        const footerY = 120;
        page.drawLine({
            start: { x: 50, y: footerY },
            end: { x: width - 50, y: footerY },
            thickness: 1,
            color: rgb(0.4, 0.4, 0.4)
        });
        
        // Footer Column 1 - Company info (smaller font)
        let col1X = 50;
        let footerTextY = footerY - 12;
        
        page.drawText('DEKRA Automobil GmbH', {
            x: col1X,
            y: footerTextY,
            size: 7,
            font: helveticaBold
        });
        
        page.drawText('Handwerkstraße 15', {
            x: col1X,
            y: footerTextY - 8,
            size: 6,
            font: helvetica
        });
        
        page.drawText('D-70565 Stuttgart', {
            x: col1X,
            y: footerTextY - 16,
            size: 6,
            font: helvetica
        });
        
        page.drawText('Telefon (07 11) 78 61-0', {
            x: col1X,
            y: footerTextY - 24,
            size: 6,
            font: helvetica
        });
        
        page.drawText('Telefax (07 11) 78 61-22 40', {
            x: col1X,
            y: footerTextY - 32,
            size: 6,
            font: helvetica
        });
        
        page.drawText('www.dekra.com', {
            x: col1X,
            y: footerTextY - 40,
            size: 6,
            font: helvetica
        });
        
        // Footer Column 2 - Legal info
        let col2X = 220;
        
        page.drawText('Sitz Stuttgart, Amtsgericht Stuttgart,', {
            x: col2X,
            y: footerTextY,
            size: 6,
            font: helvetica
        });
        
        page.drawText('HRB-Nr. 21039', {
            x: col2X,
            y: footerTextY - 8,
            size: 6,
            font: helvetica
        });
        
        page.drawText('Bankverbindungen:', {
            x: col2X,
            y: footerTextY - 20,
            size: 6,
            font: helveticaBold
        });
        
        page.drawText('Commerzbank AG', {
            x: col2X,
            y: footerTextY - 28,
            size: 6,
            font: helvetica
        });
        
        page.drawText('IBAN DE84 6008 0000 0901 0051 00', {
            x: col2X,
            y: footerTextY - 36,
            size: 5,
            font: helvetica
        });
        
        page.drawText('BIC DRESDEFF600', {
            x: col2X,
            y: footerTextY - 42,
            size: 5,
            font: helvetica
        });
        
        // Footer Column 3 - Management
        let col3X = 430;
        
        page.drawText('Vorsitzender des Aufsichtsrates:', {
            x: col3X,
            y: footerTextY,
            size: 6,
            font: helveticaBold
        });
        
        page.drawText('Stefan Kölbl', {
            x: col3X,
            y: footerTextY - 8,
            size: 6,
            font: helvetica
        });
        
        page.drawText('Geschäftsführer:', {
            x: col3X,
            y: footerTextY - 20,
            size: 6,
            font: helveticaBold
        });
        
        page.drawText('Guido Kutschera (Vorsitzender)', {
            x: col3X,
            y: footerTextY - 28,
            size: 6,
            font: helvetica
        });
        
        page.drawText('Friedemann Bausch', {
            x: col3X,
            y: footerTextY - 36,
            size: 6,
            font: helvetica
        });
        
        page.drawText('Jann Fehlauer', {
            x: col3X,
            y: footerTextY - 44,
            size: 6,
            font: helvetica
        });
        
        // Process attachments - add them AFTER the first page (ignore Word documents)
        if (appState.attachedFiles.length > 0) {
            for (const file of appState.attachedFiles) {
                try {
                    // Skip Word documents in PDF export
                    if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
                        file.type === 'application/msword' || 
                        file.name.endsWith('.docx') || 
                        file.name.endsWith('.doc')) {
                        continue; // Skip Word files
                    }
                    
                    if (file.type === 'application/pdf' || file.name.endsWith('.pdf')) {
                        // PDF files - copy pages directly
                        const attachmentBytes = await file.arrayBuffer();
                        const attachmentDoc = await PDFDocument.load(attachmentBytes);
                        const pages = await pdfDoc.copyPages(attachmentDoc, attachmentDoc.getPageIndices());
                        pages.forEach(page => pdfDoc.addPage(page));
                    } else if (file.type.startsWith('image/')) {
                        // Image files - add as new page
                        const imageBytes = await file.arrayBuffer();
                        let embedImage;
                        
                        if (file.type === 'image/png') {
                            embedImage = await pdfDoc.embedPng(imageBytes);
                        } else if (file.type === 'image/jpeg' || file.type === 'image/jpg') {
                            embedImage = await pdfDoc.embedJpg(imageBytes);
                        }
                        
                        if (embedImage) {
                            const imagePage = pdfDoc.addPage([595, 842]);
                            const { width, height } = embedImage.scale(0.5);
                            imagePage.drawImage(embedImage, {
                                x: (595 - width) / 2,
                                y: (842 - height) / 2,
                                width: width,
                                height: height
                            });
                        }
                    }
                } catch (error) {
                    console.error(`Fehler beim Verarbeiten von ${file.name}:`, error);
                }
            }
        }
        
        // Save and download
        const pdfBytes = await pdfDoc.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = `Gutachten_${formData.gutachtenNummer || 'draft'}.pdf`;
        a.click();
        URL.revokeObjectURL(url);
        
    } catch (error) {
        console.error('Fehler beim PDF-Export:', error);
        alert('Fehler beim Erstellen der PDF. Bitte versuchen Sie es erneut.');
    } finally {
        hideProgress();
    }
}

// Export to Word - Exact format like preview
async function exportToWord() {
    showProgress('Exportiere Word-Dokument...');
    
    try {
        const formData = collectFormData();
        const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, Header, Footer } = docx;
        
        // Create document with exact formatting
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1134, // 2cm
                            right: 1134,
                            bottom: 1134,
                            left: 1134
                        }
                    }
                },
                footers: {
                    default: new Footer({
                        children: [
                            new Paragraph({ text: "_____________________________________________________________________________", alignment: AlignmentType.CENTER }),
                            new Table({
                                width: { size: 100, type: WidthType.PERCENTAGE },
                                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
                                rows: [
                                    new TableRow({
                                        children: [
                                            new TableCell({
                                                children: [
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "DEKRA Automobil GmbH", size: 16, bold: true })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Handwerkstraße 15", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "D-70565 Stuttgart", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Telefon (07 11) 78 61-0", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Telefax (07 11) 78 61-22 40", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "www.dekra.com", size: 14 })
                                                        ]
                                                    })
                                                ],
                                                width: { size: 33, type: WidthType.PERCENTAGE },
                                                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                            }),
                                            new TableCell({
                                                children: [
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Sitz Stuttgart, Amtsgericht Stuttgart,", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "HRB-Nr. 21039", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ text: "" }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Bankverbindungen:", size: 14, bold: true })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Commerzbank AG", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "IBAN DE84 6008 0000 0901 0051 00", size: 12 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "BIC DRESDEFF600", size: 12 })
                                                        ]
                                                    })
                                                ],
                                                width: { size: 34, type: WidthType.PERCENTAGE },
                                                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                            }),
                                            new TableCell({
                                                children: [
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Vorsitzender des Aufsichtsrates:", size: 14, bold: true })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Stefan Kölbl", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ text: "" }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Geschäftsführer:", size: 14, bold: true })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Guido Kutschera (Vorsitzender)", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Friedemann Bausch", size: 14 })
                                                        ]
                                                    }),
                                                    new Paragraph({ 
                                                        children: [
                                                            new TextRun({ text: "Jann Fehlauer", size: 14 })
                                                        ]
                                                    })
                                                ],
                                                width: { size: 33, type: WidthType.PERCENTAGE },
                                                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                            })
                                        ]
                                    })
                                ]
                            })
                        ]
                    })
                },
                children: [
                    // Header table with DEKRA info and "Blatt 1"
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: "DEKRA Automobil GmbH",
                                                        bold: true,
                                                        size: 28,
                                                        color: "000000"
                                                    })
                                                ]
                                            })
                                        ],
                                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({ text: "" })
                                        ],
                                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                    })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: "Niederlassung Düsseldorf Fachbereich Polizei- und Gerichtsgutachten",
                                                        size: 22,
                                                        color: "000000"
                                                    })
                                                ]
                                            })
                                        ],
                                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({ text: "" })
                                        ],
                                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                    })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: "Höherweg 111, 40233 Düsseldorf, Telefon 0211/2300-0, Fax 0211/2300-222",
                                                        size: 20,
                                                        color: "333333"
                                                    })
                                                ]
                                            })
                                        ],
                                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: "Blatt 1",
                                                        size: 20,
                                                        color: "000000"
                                                    })
                                                ],
                                                alignment: AlignmentType.RIGHT
                                            })
                                        ],
                                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                    })
                                ]
                            })
                        ]
                    }),
                    
                    new Paragraph({
                        text: "",
                        spacing: { after: 200 },
                        border: {
                            bottom: {
                                color: "000000",
                                space: 1,
                                style: BorderStyle.SINGLE,
                                size: 6
                            }
                        }
                    }),
                    
                    // Spacing
                    new Paragraph({ text: "" }),
                    
                    // Create table for recipient and document info
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
                        rows: [
                            new TableRow({
                                children: [
                                    // Left column - Recipient
                                    new TableCell({
                                        children: [
                                            new Paragraph({ 
                                                children: [
                                                    new TextRun({ text: formData.empfaengerName || "KPB Neuss", size: 20 })
                                                ]
                                            }),
                                            new Paragraph({ 
                                                children: [
                                                    new TextRun({ text: formData.empfaengerStrasse || "Holbeinstraße 4", size: 20 })
                                                ]
                                            }),
                                            new Paragraph({ 
                                                children: [
                                                    new TextRun({ text: `${formData.empfaengerPLZ || '40667'} ${formData.empfaengerOrt || 'Meerbusch'}`, size: 20 })
                                                ]
                                            })
                                        ],
                                        width: { size: 50, type: WidthType.PERCENTAGE },
                                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                    }),
                                    // Right column - Document info
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: "Gutachten-Nummer    ", size: 20 }),
                                                    new TextRun({ text: formData.gutachtenNummer || "302/53884 25-4322909943", bold: true, size: 20 })
                                                ]
                                            }),
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: "vom                 ", size: 20 }),
                                                    new TextRun({ text: formatDate(formData.datum) || "02.09.2025", size: 20 })
                                                ]
                                            }),
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: "Kunden-Nummer       ", size: 20 }),
                                                    new TextRun({ text: formData.kundenNummer || "212500140", size: 20 })
                                                ]
                                            })
                                        ],
                                        width: { size: 50, type: WidthType.PERCENTAGE },
                                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
                                    })
                                ]
                            })
                        ]
                    }),
                    
                    // Spacing
                    new Paragraph({ text: "", spacing: { before: 600, after: 600 } }),
                    
                    // Title with letter spacing effect
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "G U T A C H T E N",
                                bold: true,
                                size: 32,
                                color: "000000"
                            })
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 600, after: 600 }
                    }),
                    
                    // Spacing
                    new Paragraph({ text: "", spacing: { before: 400 } }),
                    
                    // Details with colons
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
                        rows: [
                            createDetailRow("Aktenzeichen", formData.aktenzeichen || ""),
                            createDetailRow("Beteiligte / Sache", formData.beteiligte || ""),
                            createDetailRow("Auftraggeber", formData.auftraggeber || ""),
                            createDetailRow("Auftrag vom", formatDate(formData.auftragVom) ? formatDate(formData.auftragVom) + " (Akteneingang)" : ""),
                            createDetailRow("Besichtigungsdatum", formData.besichtigungsdatum ? formatDateLong(formData.besichtigungsdatum) : ""),
                            createDetailRow("Besichtigungsort", formData.besichtigungsort || ""),
                            createDetailRow("Sachverständiger", formData.sachverstaendiger || "")
                        ]
                    })
                ]
            }]
        });
        
        // KEINE Anhänge im Word-Export - nur die ausgefüllte Vorlage!
        // Der folgende Code ist komplett deaktiviert:
        /*
        // Add attachment pages if there are attachments
        if (appState.attachedFiles.length > 0) {
            // Add separator page
            doc.addSection({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "ANHÄNGE",
                                bold: true,
                                size: 48
                            })
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 1200, after: 600 }
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Folgende Dokumente sind diesem Gutachten beigefügt:",
                                size: 28
                            })
                        ],
                        spacing: { after: 600 }
                    })
                ]
            });
            
            // Process each attachment - DEAKTIVIERT
            for (const file of appState.attachedFiles) {
                // For Word documents, try to read and append content
                if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
                    file.type === 'application/msword' || 
                    file.name.endsWith('.docx') || 
                    file.name.endsWith('.doc')) {
                    
                    try {
                        // Read the Word file content using mammoth.js
                        const arrayBuffer = await file.arrayBuffer();
                        
                        // Convert Word document to HTML using mammoth
                        const result = await mammoth.convertToHtml({arrayBuffer: arrayBuffer});
                        const htmlContent = result.value;
                        
                        // Parse HTML to extract text content
                        const tempDiv = document.createElement('div');
                        tempDiv.innerHTML = htmlContent;
                        
                        // Extract paragraphs and create Word paragraphs
                        const paragraphs = [];
                        
                        // Add title for attachment
                        paragraphs.push(
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: "",
                                        break: 1 // Page break
                                    })
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `--- ${file.name} ---`,
                                        bold: true,
                                        size: 28
                                    })
                                ],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 400 }
                            })
                        );
                        
                        // Process all paragraphs from the HTML
                        const htmlParagraphs = tempDiv.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li');
                        htmlParagraphs.forEach(element => {
                            const text = element.textContent.trim();
                            if (text) {
                                // Check if it's a heading
                                const isHeading = element.tagName.startsWith('H');
                                const isList = element.tagName === 'LI';
                                
                                paragraphs.push(
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: isList ? `• ${text}` : text,
                                                bold: isHeading,
                                                size: isHeading ? 28 : 22
                                            })
                                        ],
                                        spacing: { after: isHeading ? 400 : 200 }
                                    })
                                );
                            }
                        });
                        
                        // If no paragraphs found, add the raw text
                        if (paragraphs.length === 2) {
                            const fullText = tempDiv.textContent.trim();
                            if (fullText) {
                                // Split by newlines and add as paragraphs
                                const lines = fullText.split('\n').filter(line => line.trim());
                                lines.forEach(line => {
                                    paragraphs.push(
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: line,
                                                    size: 22
                                                })
                                            ],
                                            spacing: { after: 200 }
                                        })
                                    );
                                });
                            }
                        }
                        
                        // Add all paragraphs to the document
                        doc.addSection({
                            children: paragraphs
                        });
                    } catch (error) {
                        console.error('Error reading Word document with mammoth:', error);
                        // Fallback if mammoth.js fails - add placeholder
                        doc.addSection({
                            children: [
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: "",
                                            break: 1 // Page break
                                        })
                                    ]
                                }),
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `--- ${file.name} ---`,
                                            bold: true,
                                            size: 28
                                        })
                                    ],
                                    alignment: AlignmentType.CENTER,
                                    spacing: { after: 400 }
                                }),
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `[Fehler beim Lesen des Word-Dokuments]`,
                                            italics: true,
                                            size: 22,
                                            color: "FF0000"
                                        })
                                    ],
                                    alignment: AlignmentType.CENTER,
                                    spacing: { after: 400 }
                                }),
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `Das Dokument "${file.name}" konnte nicht automatisch gelesen werden.`,
                                            size: 22
                                        })
                                    ],
                                    spacing: { after: 200 }
                                }),
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `Dateigröße: ${formatFileSize(file.size)}`,
                                            size: 20
                                        })
                                    ],
                                    spacing: { after: 200 }
                                }),
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `Bitte fügen Sie den Inhalt manuell ein.`,
                                            size: 22,
                                            italics: true
                                        })
                                    ]
                                })
                            ]
                        });
                    }
                } else {
                    // For other file types, add info page
                    doc.addSection({
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Anhang: ${file.name}`,
                                        bold: true,
                                        size: 32
                                    })
                                ],
                                spacing: { after: 400 }
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Dateigröße: ${formatFileSize(file.size)}`,
                                        size: 24
                                    })
                                ],
                                spacing: { after: 200 }
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `Dateityp: ${file.type || 'Unbekannt'}`,
                                        size: 24
                                    })
                                ],
                                spacing: { after: 600 }
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: file.type === 'application/pdf' ? 
                                              "PDF-Dokument - Wird beim PDF-Export automatisch angehängt." :
                                              file.type.startsWith('image/') ?
                                              "Bilddatei - Wird beim PDF-Export als neue Seite eingefügt." :
                                              "Dokument angehängt.",
                                        size: 22,
                                        italics: true
                                    })
                                ]
                            })
                        ]
                    });
                }
            }
        }
        */
        
        // Generate and download
        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Gutachten_${formData.gutachtenNummer || 'draft'}.docx`;
        a.click();
        URL.revokeObjectURL(url);
        
    } catch (error) {
        console.error('Fehler beim Word-Export:', error);
        alert('Fehler beim Erstellen des Word-Dokuments. Bitte versuchen Sie es erneut.');
    } finally {
        hideProgress();
    }
}

// Helper function to create detail rows for Word table
function createDetailRow(label, value) {
    const { TableRow, TableCell, Paragraph, TextRun, WidthType, BorderStyle } = docx;
    
    return new TableRow({
        children: [
            new TableCell({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: label, size: 22 })
                        ]
                    })
                ],
                width: { size: 30, type: WidthType.PERCENTAGE },
                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
            }),
            new TableCell({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: ":", size: 22 })
                        ]
                    })
                ],
                width: { size: 5, type: WidthType.PERCENTAGE },
                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
            }),
            new TableCell({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: value, size: 22 })
                        ]
                    })
                ],
                width: { size: 65, type: WidthType.PERCENTAGE },
                borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }
            })
        ]
    });
}


// Fill form with data
function fillFormWithData(data) {
    if (!data || typeof data !== 'object') return;
    
    try {
        for (const [key, value] of Object.entries(data)) {
            const field = document.querySelector(`[name="${key}"]`);
            if (field) {
                field.value = value || '';
                appState.formData[key] = value || '';
            }
        }
    } catch (error) {
        console.error('Error filling form:', error);
    }
}

// Auto-save functionality
function initializeAutoSave() {
    setInterval(() => {
        saveToLocalStorage();
    }, 30000); // Save every 30 seconds
}

function saveToLocalStorage() {
    try {
        const formData = collectFormData();
        localStorage.setItem('dekra_current_form', JSON.stringify(formData));
    } catch (error) {
        console.error('Error saving to localStorage:', error);
    }
}

// Set default dates
function setDefaultDates() {
    try {
        const today = new Date().toISOString().split('T')[0];
        const datumField = document.getElementById('datum');
        const auftragField = document.getElementById('auftragVom');
        const besichtigungField = document.getElementById('besichtigungsdatum');
        
        if (datumField) datumField.value = today;
        if (auftragField) auftragField.value = today;
        if (besichtigungField) besichtigungField.value = today;
    } catch (error) {
        console.error('Error setting default dates:', error);
    }
}

// Keyboard shortcuts
function handleKeyboardShortcuts(event) {
    // Ctrl+S: Save
    if (event.ctrlKey && event.key === 's') {
        event.preventDefault();
        saveToLocalStorage();
        showNotification('Gespeichert');
    }
    
}

// Progress indicator
function showProgress(message) {
    const indicator = document.getElementById('progressIndicator');
    if (!indicator) return;
    
    const progressText = indicator.querySelector('.progress-text');
    const progressFill = indicator.querySelector('.progress-fill');
    
    if (progressText) progressText.textContent = message;
    if (progressFill) progressFill.style.width = '0%';
    indicator.classList.remove('hidden');
    
    // Animate progress
    let progress = 0;
    const interval = setInterval(() => {
        progress += 10;
        if (progressFill) progressFill.style.width = `${progress}%`;
        if (progress >= 90) {
            clearInterval(interval);
        }
    }, 100);
}

function hideProgress() {
    const indicator = document.getElementById('progressIndicator');
    if (!indicator) return;
    
    const progressFill = indicator.querySelector('.progress-fill');
    if (progressFill) progressFill.style.width = '100%';
    
    setTimeout(() => {
        indicator.classList.add('hidden');
    }, 500);
}

// Show notification
function showNotification(message) {
    // Create notification element
    const notification = document.createElement('div');
    notification.className = 'notification';
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: #28A745;
        color: white;
        padding: 15px 25px;
        border-radius: 4px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        z-index: 3000;
        animation: slideIn 0.3s ease;
    `;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.remove();
    }, 3000);
}

