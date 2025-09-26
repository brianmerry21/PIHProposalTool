// Main JavaScript for PIH Pricing Proposal Converter

document.addEventListener('DOMContentLoaded', function() {
    // Set up form submission handling to show loading animation
    const uploadForm = document.getElementById('upload-form');
    const loadingOverlay = document.getElementById('loading-overlay');

    if (uploadForm) {
        uploadForm.addEventListener('submit', function(e) {
            // Check if form is valid before showing the loading overlay
            if (uploadForm.checkValidity()) {
                // Show loading overlay with VLM animation
                loadingOverlay.style.display = 'flex';
                
                // Update loading message with processing stages
                const loadingMessage = document.getElementById('loading-message');
                
                // Simulate progress messages (optional)
                const messages = [
                    "Uploading PDF...",
                    "Extracting text from PDF...",
                    "Analyzing document structure...",
                    "Extracting customer information...",
                    "Processing line items...",
                    "Calculating margins...",
                    "Creating Excel document...",
                    "Generating Word proposal..."
                ];
                
                let messageIndex = 0;
                const messageInterval = setInterval(function() {
                    if (messageIndex < messages.length) {
                        loadingMessage.innerHTML = `
                            <p>${messages[messageIndex]}</p>
                            <p class="small text-muted">Processing page ${Math.floor(Math.random() * 10) + 1}</p>
                        `;
                        messageIndex++;
                    } else {
                        clearInterval(messageInterval);
                        loadingMessage.innerHTML = `
                            <p>Finalizing documents...</p>
                            <p class="small text-muted">Almost done!</p>
                        `;
                    }
                }, 3000);
            }
        });
    }
    
    // If we're on the review page, set up the same loading animation for the processing form
    const processForm = document.getElementById('process-form');
    if (processForm) {
        processForm.addEventListener('submit', function(e) {
            loadingOverlay.style.display = 'flex';
            const loadingMessage = document.getElementById('loading-message');
            loadingMessage.innerHTML = `
                <p>Generating final documents...</p>
                <p class="small text-muted">This may take a few moments</p>
            `;
        });
    }
    
    // Set up toggle switches for custom info
    const customInfoToggle = document.getElementById('customInfoToggle');
    const customInfoFields = document.getElementById('customInfoFields');
    
    if (customInfoToggle && customInfoFields) {
        customInfoToggle.addEventListener('change', function() {
            customInfoFields.style.display = this.checked ? 'block' : 'none';
        });
    }
    
    // Initialize tooltips if Bootstrap is available
    if (typeof bootstrap !== 'undefined' && bootstrap.Tooltip) {
        const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
        tooltipTriggerList.map(function (tooltipTriggerEl) {
            return new bootstrap.Tooltip(tooltipTriggerEl);
        });
    }
});

// Function to hide the loading overlay (can be called explicitly)
function hideLoadingOverlay() {
    document.getElementById('loading-overlay').style.display = 'none';
}