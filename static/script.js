document.addEventListener('DOMContentLoaded', () => {
    // --- Custom Cursor Logic (Trailing) ---
    const cursorDot = document.querySelector('[data-cursor-dot]');
    const cursorOutline = document.querySelector('[data-cursor-outline]');

    // Position storage
    let mouseX = 0;
    let mouseY = 0;
    let outlineX = 0;
    let outlineY = 0;

    // Track mouse
    document.addEventListener('mousemove', (e) => {
        mouseX = e.clientX;
        mouseY = e.clientY;

        // Instant update for dot
        cursorDot.style.left = `${mouseX}px`;
        cursorDot.style.top = `${mouseY}px`;

        // CSS Variable for spotlight
        document.body.style.setProperty('--mouse-x', `${mouseX}px`);
        document.body.style.setProperty('--mouse-y', `${mouseY}px`);

        // Optional: Interactive hover state for cursor
        const target = e.target;
        if (target.matches('button, a, input, .drop-zone, .icon-btn, .nav-btn, .close-icon, h2, i')) {
            cursorOutline.style.width = '60px';
            cursorOutline.style.height = '60px';
            cursorOutline.style.borderColor = 'rgba(94, 106, 210, 0.8)';
            cursorOutline.style.backgroundColor = 'rgba(94, 106, 210, 0.05)';
        } else {
            cursorOutline.style.width = '40px';
            cursorOutline.style.height = '40px';
            cursorOutline.style.borderColor = 'rgba(255, 255, 255, 0.5)';
            cursorOutline.style.backgroundColor = 'transparent';
        }
    });

    // Smooth loop for outline
    function animateCursor() {
        // Linear Interpolation (Lerp) for smooth lag
        // 0.3 is faster (less laggy) than 0.15
        outlineX += (mouseX - outlineX) * 0.3;
        outlineY += (mouseY - outlineY) * 0.3;

        cursorOutline.style.left = `${outlineX}px`;
        cursorOutline.style.top = `${outlineY}px`;

        requestAnimationFrame(animateCursor);
    }

    // Start animation loop
    animateCursor();

    // --- Elements ---
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const uploadForm = document.getElementById('uploadForm');

    // States
    const initialState = document.getElementById('initialState');
    const processingState = document.getElementById('processingState');
    const successState = document.getElementById('successState');
    const errorState = document.getElementById('errorState');

    // Output
    const fileNameDisplay = document.getElementById('fileName');
    const downloadLink = document.getElementById('downloadLink');
    const resetBtn = document.getElementById('resetBtn');
    const errorText = document.getElementById('errorText');

    // Modal
    const helpBtn = document.getElementById('helpBtn');
    const closeHelp = document.getElementById('closeHelp');
    const helpModal = document.getElementById('helpModal');

    // --- Modal Logic ---
    helpBtn.addEventListener('click', () => {
        helpModal.classList.add('active');
    });

    closeHelp.addEventListener('click', () => {
        helpModal.classList.remove('active');
    });

    helpModal.addEventListener('click', (e) => {
        if (e.target === helpModal) { // Click on backdrop
            helpModal.classList.remove('active');
        }
    });

    // Close on Escape Key
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && helpModal.classList.contains('active')) {
            helpModal.classList.remove('active');
        }
    });


    // --- Upload Interaction Logic ---

    // Click to browse (only if not processing/success)
    dropZone.addEventListener('click', () => {
        if (getComputedStyle(initialState).display !== 'none') {
            fileInput.click();
        }
    });

    // Reset
    resetBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        resetUI();
    });

    function resetUI() {
        uploadForm.reset();
        showState(initialState);
        dropZone.classList.remove('drag-over');
        fileNameDisplay.textContent = '';
        fileNameDisplay.style.opacity = '0'; // reset fade in
    }

    function showState(stateElement) {
        [initialState, processingState, successState, errorState].forEach(el => el.style.display = 'none');
        stateElement.style.display = 'flex'; // or block for text elements, but flex for these containers
    }

    // Drag & Drop Handling
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('drag-over'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('drag-over'), false);
    });

    dropZone.addEventListener('drop', handleDrop, false);
    fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

    function handleDrop(e) {
        const dt = e.dataTransfer;
        handleFiles(dt.files);
    }

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            if (isValidFile(file)) {
                fileNameDisplay.textContent = file.name;
                processFile(file);
            } else {
                showError("Please upload a valid Excel file (.xlsx)");
            }
        }
    }

    function isValidFile(file) {
        return file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
    }

    function showError(msg) {
        showState(errorState);
        errorText.textContent = msg;
    }

    // --- Backend Processing ---

    function processFile(file) {
        showState(processingState);

        const formData = new FormData();
        formData.append('file', file);
        formData.append('tolerance', '1.0');

        fetch('/upload', {
            method: 'POST',
            body: formData
        })
            .then(response => response.json())
            .then(data => {
                if (data.error) {
                    showError(data.error);
                } else {
                    showSuccess(data.download_url);
                }
            })
            .catch(error => {
                console.error('Error:', error);
                showError("Server unable to match. Please check file format.");
            });
    }

    function showSuccess(url) {
        showState(successState);
        downloadLink.href = url;
    }
});
