// ===================================
// Global State
// ===================================
let participants = [];
let drawResults = [];

// ===================================
// DOM Elements
// ===================================
const fileInput = document.getElementById('fileInput');
const participantInfo = document.getElementById('participantInfo');
const participantCount = document.getElementById('participantCount');
const drawBtn = document.getElementById('drawBtn');
const redrawBtn = document.getElementById('redrawBtn');
const resultsSection = document.getElementById('resultsSection');
const resultTableBody = document.getElementById('resultTableBody');
const messageBox = document.getElementById('messageBox');

// ===================================
// Event Listeners
// ===================================
fileInput.addEventListener('change', handleFileUpload);
drawBtn.addEventListener('click', handleDraw);
redrawBtn.addEventListener('click', handleRedraw);

// ===================================
// File Upload Handler
// ===================================
function handleFileUpload(event) {
    const file = event.target.files[0];

    if (!file) {
        return;
    }

    // Validate file type
    const validTypes = [
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];

    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showMessage('請上傳 Excel 檔案 (.xlsx 或 .xls)', 'error');
        fileInput.value = '';
        return;
    }

    // Read the file
    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Get first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            if (jsonData.length === 0) {
                showMessage('Excel 檔案為空，請檢查檔案內容', 'error');
                return;
            }

            // Parse and validate data
            parseExcelData(jsonData);

        } catch (error) {
            console.error('Excel parsing error:', error);
            showMessage('Excel 檔案格式錯誤，請檢查檔案內容', 'error');
        } finally {
            // Clear the file input value to allow re-uploading the same file
            event.target.value = '';
        }
    };

    reader.onerror = function () {
        showMessage('檔案讀取失敗，請重試', 'error');
        // Clear the file input value
        event.target.value = '';
    };

    reader.readAsArrayBuffer(file);
}

// ===================================
// Excel Data Parser (Async with performance optimization)
// ===================================
async function parseExcelData(jsonData) {
    // Show loading state
    showMessage('正在處理資料...', 'success');

    // Step 1: Process data asynchronously to avoid blocking UI
    const { parsedParticipants, invalidRows } = await processParticipantsAsync(jsonData);
    console.log('parsedParticipants:', parsedParticipants)

    // Step 2: Check for invalid rows
    if (invalidRows.length > 0) {
        const names = invalidRows.join('、');
        showMessage(`${names} 資料不完整，因此不進行配對`, 'warning');
    }

    // Step 3: Check for duplicates (runs after Step 1 completes)
    const { cleanedData, duplicates } = await removeDuplicatesAsync(parsedParticipants);

    if (duplicates.length > 0) {
        const dupNames = duplicates.join('、');
        showMessage(`偵測到重複資料: ${dupNames}，已自動去除重複`, 'warning');
    }

    // Step 4: Update global state (runs after Step 3 completes)
    participants = cleanedData;

    // Step 5: Update UI (runs after Step 4 completes)
    updateUIAfterParsing(invalidRows, duplicates);
}

// ===================================
// Process Participants Asynchronously
// ===================================
function processParticipantsAsync(jsonData) {
    return new Promise((resolve) => {
        // Use setTimeout to avoid blocking the UI thread
            const parsedParticipants = [];
            const invalidRows = [];

            // Optimize: Use for loop instead of forEach for better performance
            for (let index = 0; index < jsonData.length; index++) {
                const row = jsonData[index];

                // Extract fields - handle different possible column names
                const participant = {
                    id: index + 1,
                    timestamp: row['時間戳記'] || '',
                    igAccount: row['您的 IG 帳號'] || '',
                    name: row['姓名'] || '',
                    phone: row['電話'] || '',
                    store: row['7-11 店號+店名'] || ''
                };

                // Validate required fields (timestamp is optional)
                const isValid = participant.igAccount &&
                    participant.name &&
                    participant.phone &&
                    participant.store;

                if (isValid) {
                    // Clean up data inline for better performance
                    parsedParticipants.push({
                        id: participant.id,
                        timestamp: participant.timestamp,
                        igAccount: String(participant.igAccount).trim(),
                        name: String(participant.name).trim(),
                        phone: String(participant.phone).trim(),
                        store: String(participant.store).trim()
                    });
                } else {
                    const identifier = participant.name || `第 ${index + 1} 筆資料`;
                    invalidRows.push(identifier);
                }
            }
            resolve({ parsedParticipants, invalidRows });
    });
}

// ===================================
// Update UI After Parsing
// ===================================
function updateUIAfterParsing(invalidRows, duplicates) {
    // Check if no valid participants
    if (participants.length === 0) {
        showMessage('無有效參與者資料，請檢查 Excel 檔案', 'error');
        resetUI();
        return;
    }

    // Show participant info
    participantCount.textContent = `已匯入 ${participants.length} 位參與者`;
    participantInfo.style.display = 'inline-flex';

    // Check minimum participants
    if (participants.length < 2) {
        showMessage('至少需要 2 人才能進行抽籤', 'error');
        drawBtn.disabled = true;
    } else if (participants.length === 2) {
        showMessage('建議至少需要 3 人才能不重複抽籤', 'warning');
        drawBtn.disabled = false;
    } else {
        drawBtn.disabled = false;
        if (invalidRows.length === 0 && duplicates.length === 0) {
            showMessage(`成功匯入 ${participants.length} 位參與者！`, 'success');
        }
    }

    // Hide results if previously shown
    resultsSection.style.display = 'none';
    redrawBtn.style.display = 'none';
}

// ===================================
// Remove Duplicate Participants (Async)
// ===================================
function removeDuplicatesAsync(data) {
    return new Promise((resolve) => {
        // Use setTimeout to avoid blocking the UI thread
        setTimeout(() => {
            const seen = new Map();
            const cleanedData = [];
            const duplicates = [];

            // Optimize: Use for loop instead of forEach for better performance
            for (let i = 0; i < data.length; i++) {
                const participant = data[i];

                // Create unique key based on name + IG + phone
                const key = `${participant.name}_${participant.igAccount}_${participant.phone}`;

                if (!seen.has(key)) {
                    seen.set(key, true);
                    cleanedData.push(participant);
                } else {
                    duplicates.push(participant.name);
                }
            }

            resolve({ cleanedData, duplicates });
        }, 0);
    });
}

// ===================================
// Draw Handler
// ===================================
function handleDraw() {
    if (participants.length < 2) {
        showMessage('至少需要 2 人才能進行抽籤', 'error');
        return;
    }

    // Perform the draw
    const result = generatePairs(participants);

    if (result.success) {
        drawResults = result.pairs;
        displayResults(drawResults);
        redrawBtn.style.display = 'inline-flex';
        showMessage('抽籤成功！', 'success');
    } else {
        showMessage('抽籤失敗，請重試', 'error');
        console.error('Draw failed after maximum attempts');
    }
}

// ===================================
// Redraw Handler
// ===================================
function handleRedraw() {
    handleDraw();
}

// ===================================
// Generate Pairs Algorithm
// ===================================
function generatePairs(participantList) {
    const maxAttempts = 1000;
    let attempts = 0;

    while (attempts < maxAttempts) {
        attempts++;

        // Shuffle participants
        const shuffled = shuffleArray([...participantList]);

        // Create pairs
        const pairs = participantList.map((giver, index) => ({
            giver: giver,
            receiver: shuffled[index]
        }));

        // Validate pairs
        if (validatePairs(pairs)) {
            return {
                success: true,
                pairs: pairs
            };
        }
    }

    return {
        success: false,
        pairs: []
    };
}

// ===================================
// Fisher-Yates Shuffle Algorithm
// ===================================
function shuffleArray(array) {
    const shuffled = [...array];

    for (let i = shuffled.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }

    return shuffled;
}

// ===================================
// Validate Pairs
// ===================================
function validatePairs(pairs) {
    // Check 1: No one should receive gift from themselves
    for (let pair of pairs) {
        if (pair.giver.id === pair.receiver.id) {
            return false;
        }
    }

    // Check 2: No mutual pairs (A→B and B→A)
    if (checkMutualPairs(pairs)) {
        return false;
    }

    return true;
}

// ===================================
// Check for Mutual Pairs
// ===================================
function checkMutualPairs(pairs) {
    for (let i = 0; i < pairs.length; i++) {
        for (let j = i + 1; j < pairs.length; j++) {
            const pair1 = pairs[i];
            const pair2 = pairs[j];

            // Check if A→B and B→A
            if (pair1.giver.id === pair2.receiver.id &&
                pair1.receiver.id === pair2.giver.id) {
                return true; // Found mutual pair
            }
        }
    }
    return false; // No mutual pairs found
}

// ===================================
// Display Results
// ===================================
function displayResults(pairs) {
    // Clear previous results
    resultTableBody.innerHTML = '';

    // Create table rows
    pairs.forEach((pair, index) => {
        const row = document.createElement('tr');

        row.innerHTML = `
            <td>${pair.giver.name}</td>
            <td>${pair.giver.igAccount}</td>
            <td>${formatPhone(pair.giver.phone)}</td>
            <td>${pair.giver.store}</td>
            <td class="arrow-cell">→</td>
            <td>${pair.receiver.name}</td>
            <td>${pair.receiver.igAccount}</td>
            <td>${formatPhone(pair.receiver.phone)}</td>
            <td>${pair.receiver.store}</td>
        `;

        resultTableBody.appendChild(row);
    });

    // Show results section
    resultsSection.style.display = 'block';

    // Scroll to results
    setTimeout(() => {
        resultsSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }, 100);
}

// ===================================
// Format Phone Number
// ===================================
function formatPhone(phone) {
    // Convert to string and remove any non-digit characters
    const cleaned = String(phone).replace(/\D/g, '');

    // Format as 0912-345-678 or similar
    if (cleaned.length === 10) {
        return `${cleaned.slice(0, 4)}-${cleaned.slice(4, 7)}-${cleaned.slice(7)}`;
    }

    return phone; // Return original if format doesn't match
}

// ===================================
// Show Message
// ===================================
function showMessage(text, type = 'success') {
    const messageContent = messageBox.querySelector('.message-content');
    const messageIcon = messageBox.querySelector('.message-icon');
    const messageText = messageBox.querySelector('.message-text');
    const closeButton = messageBox.querySelector('.message-close');

    // Set icon based on type
    const icons = {
        success: '✅',
        warning: '⚠️',
        error: '❌'
    };

    messageIcon.textContent = icons[type] || icons.success;
    messageText.textContent = text;

    // Set message box class
    messageBox.className = `message-box ${type}`;
    messageBox.style.display = 'block';

    // Close button handler
    closeButton.onclick = function () {
        messageBox.style.display = 'none';
    };

    // Auto-hide success messages after 5 seconds
    if (type === 'success') {
        setTimeout(() => {
            if (messageBox.style.display !== 'none') {
                messageBox.style.display = 'none';
            }
        }, 5000);
    }
}

// ===================================
// Reset UI
// ===================================
function resetUI() {
    participantInfo.style.display = 'none';
    drawBtn.disabled = true;
    redrawBtn.style.display = 'none';
    resultsSection.style.display = 'none';
    fileInput.value = '';
}

// ===================================
// Initialize
// ===================================
console.log('Christmas Gift Exchange Draw System initialized!');
console.log('Made with ❤️ using HTML, CSS, and JavaScript');
