// Global variables
let workbook = null;
let currentSheet = null;
let extractedData = [];
let allMonths = [];
let holidays = new Set(); // Store holidays as 'YYYY-MM-DD' strings

// Firebase variables
let database = null;
let isFirebaseReady = false;

// Firebase configuration
const firebaseConfig = {
  apiKey: "AIzaSyAFuBQuONd5MOd0xYzQTMruhMIvfWWVquk",
  authDomain: "prefect-s-attendance.firebaseapp.com",
  databaseURL: "https://prefect-s-attendance-default-rtdb.firebaseio.com",
  projectId: "prefect-s-attendance",
  storageBucket: "prefect-s-attendance.firebasestorage.app",
  messagingSenderId: "386092614514",
  appId: "1:386092614514:web:4a1a8d8bb6b3160e4253eb",
  measurementId: "G-2WFVTFGHNN"
};

// Storage keys
const STORAGE_KEYS = {
    EXTRACTED_DATA: 'fingerprint_extracted_data',
    ALL_MONTHS: 'fingerprint_all_months',
    HOLIDAYS: 'fingerprint_holidays',
    LAST_UPDATED: 'fingerprint_last_updated'
};

// Wait for DOM to load before initializing
document.addEventListener('DOMContentLoaded', async function() {
    initializeFirebase();
    await initializeApp();
});

// Initialize Firebase
async function initializeFirebase() {
    try {
        if (typeof firebase !== 'undefined') {
            console.log('Initializing Firebase...');
            firebase.initializeApp(firebaseConfig);
            database = firebase.database();
            isFirebaseReady = true;
            console.log('Firebase initialized successfully');
            
            // Test connection
            await testFirebaseConnection();
        } else {
            console.warn('Firebase SDK not loaded');
        }
    } catch (error) {
        console.error('Firebase initialization failed:', error);
        isFirebaseReady = false;
    }
}

// Test Firebase connection
async function testFirebaseConnection() {
    if (!isFirebaseReady) return false;
    
    try {
        await database.ref('test/connection').set({
            timestamp: firebase.database.ServerValue.TIMESTAMP,
            status: 'connected'
        });
        console.log('Firebase connection test successful');
        showFirebaseStatus('✅ Firebase Connected', 'success');
        return true;
    } catch (error) {
        console.error('Firebase connection test failed:', error);
        showFirebaseStatus('❌ Firebase Connection Failed', 'error');
        return false;
    }
}

// Show Firebase status
function showFirebaseStatus(message, type = 'info') {
    console.log(`Firebase Status: ${message}`);
}

async function initializeApp() {
    // DOM elements
    const fileInput = document.getElementById('fileInput');
    const fileName = document.getElementById('fileName');
    const loading = document.getElementById('loading');
    const error = document.getElementById('error');
    
    // Event listeners
    if (fileInput) {
        fileInput.addEventListener('change', handleFileSelect);
    }
    
    // Load existing data from storage
    await loadDataFromStorage();
}

// Load data from storage (Firebase first, then localStorage fallback)
async function loadDataFromStorage() {
    if (isFirebaseReady) {
        return await loadDataFromFirebase();
    } else {
        return loadDataFromLocalStorage();
    }
}

// Load data from Firebase
async function loadDataFromFirebase() {
    try {
        showLoading(true);
        console.log('Loading data from Firebase...');
        
        // Load main attendance data
        const attendanceSnapshot = await database.ref('fingerprintData/attendance').once('value');
        if (attendanceSnapshot.exists()) {
            const data = attendanceSnapshot.val();
            extractedData = data.extractedData || [];
            
            // Convert ISO string timestamps back to Date objects and validate data structure
            extractedData.forEach(employee => {
                if (employee.attendanceData) {
                    employee.attendanceData.forEach(record => {
                        // Ensure morning property exists
                        if (!record.morning) {
                            record.morning = {};
                        }
                        
                        // Convert date strings back to Date objects
                        if (record.fullDate && typeof record.fullDate === 'string') {
                            record.fullDate = new Date(record.fullDate);
                        }
                        if (record.startDate && typeof record.startDate === 'string') {
                            record.startDate = new Date(record.startDate);
                        }
                        if (record.endDate && typeof record.endDate === 'string') {
                            record.endDate = new Date(record.endDate);
                        }
                    });
                }
            });
        }
        
        // Load months data
        const monthsSnapshot = await database.ref('fingerprintData/months').once('value');
        if (monthsSnapshot.exists()) {
            const data = monthsSnapshot.val();
            allMonths = data.allMonths || [];
        }
        
        // Load holidays data
        const holidaysSnapshot = await database.ref('fingerprintData/holidays').once('value');
        if (holidaysSnapshot.exists()) {
            const data = holidaysSnapshot.val();
            holidays = new Set(data.holidays || []);
        }
        
        // Display data if available
        if (extractedData.length > 0) {
            displayDataWithFilters();
            showDataStatus();
        }
        
        showLoading(false);
        console.log('Data loaded from Firebase successfully');
        
    } catch (error) {
        console.error('Error loading data from Firebase:', error);
        showLoading(false);
        showError('Error loading data from Firebase. Trying local storage...');
        // Fallback to localStorage
        loadDataFromLocalStorage();
    }
}

// Save data (Firebase first, localStorage backup)
async function saveDataFromStorage() {
    if (isFirebaseReady) {
        return await saveDataToFirebase();
    } else {
        return saveDataToLocalStorage();
    }
}

// Save data to Firebase
async function saveDataToFirebase() {
    try {
        console.log('Saving data to Firebase...');
        
        // Prepare data for Firebase (convert Date objects to ISO strings)
        const dataToSave = extractedData.map(employee => ({
            ...employee,
            attendanceData: employee.attendanceData.map(record => ({
                ...record,
                fullDate: record.fullDate ? record.fullDate.toISOString() : null,
                startDate: record.startDate ? record.startDate.toISOString() : null,
                endDate: record.endDate ? record.endDate.toISOString() : null
            }))
        }));
        
        // Save to Firebase Realtime Database
        const updates = {};
        const timestamp = firebase.database.ServerValue.TIMESTAMP;
        
        // Save attendance data
        updates['fingerprintData/attendance'] = {
            extractedData: dataToSave,
            lastUpdated: timestamp
        };
        
        // Save months data
        updates['fingerprintData/months'] = {
            allMonths: allMonths,
            lastUpdated: timestamp
        };
        
        // Save holidays data
        updates['fingerprintData/holidays'] = {
            holidays: [...holidays],
            lastUpdated: timestamp
        };
        
        await database.ref().update(updates);
        console.log('Data saved to Firebase successfully');
        
        // Also save to localStorage as backup
        saveDataToLocalStorage();
        
    } catch (error) {
        console.error('Error saving data to Firebase:', error);
        showError('Warning: Could not save data to Firebase. Saved locally instead.');
        // Fallback to localStorage
        saveDataToLocalStorage();
    }
}

// Fallback: Save data to localStorage
function saveDataToLocalStorage() {
    try {
        localStorage.setItem(STORAGE_KEYS.EXTRACTED_DATA, JSON.stringify(extractedData));
        localStorage.setItem(STORAGE_KEYS.ALL_MONTHS, JSON.stringify(allMonths));
        localStorage.setItem(STORAGE_KEYS.HOLIDAYS, JSON.stringify([...holidays]));
        localStorage.setItem(STORAGE_KEYS.LAST_UPDATED, new Date().toISOString());
        
        console.log('Data saved to localStorage successfully');
    } catch (error) {
        console.error('Error saving data to localStorage:', error);
        showError('Warning: Could not save data to browser storage.');
    }
}

// Show data status
async function showDataStatus() {
    try {
        let lastUpdated = null;
        let source = 'Unknown';
        
        if (isFirebaseReady) {
            // Get last updated from Firebase
            const attendanceSnapshot = await database.ref('fingerprintData/attendance').once('value');
            if (attendanceSnapshot.exists()) {
                const data = attendanceSnapshot.val();
                if (data.lastUpdated) {
                    lastUpdated = new Date(data.lastUpdated);
                    source = 'Firebase';
                }
            }
        }
        
        // Fallback to localStorage
        if (!lastUpdated) {
            const localLastUpdated = localStorage.getItem(STORAGE_KEYS.LAST_UPDATED);
            if (localLastUpdated) {
                lastUpdated = new Date(localLastUpdated);
                source = 'Local Storage';
            }
        }
        
        const fileName = document.getElementById('fileName');
        if (fileName && lastUpdated) {
            const storageIcon = source === 'Firebase' ? 'bi-cloud-check' : 'bi-hdd';
            const storageColor = source === 'Firebase' ? '#10b981' : '#f59e0b';
            
            fileName.innerHTML = `
                <i class="bi ${storageIcon}" style="color: ${storageColor}"></i> 
                ${source} Data (${extractedData.length} records) - 
                Last updated: ${lastUpdated.toLocaleDateString()} ${lastUpdated.toLocaleTimeString()}
                <button onclick="clearAllData()" class="btn btn-sm btn-outline-danger ms-2 mobile-hide" style="background: transparent; border: 1px solid #ef4444; color: #ef4444; padding: 0.25rem 0.5rem; border-radius: 6px; font-size: 0.75rem;">
                    <i class="bi bi-trash"></i> Clear All Data
                </button>
                ${source === 'Firebase' ? 
                    `<button onclick="syncToFirebase()" class="btn btn-sm btn-outline-success ms-2 mobile-hide" style="background: transparent; border: 1px solid #10b981; color: #10b981; padding: 0.25rem 0.5rem; border-radius: 6px; font-size: 0.75rem;">
                        <i class="bi bi-cloud-upload"></i> Sync
                    </button>` : ''
                }
            `;
            fileName.style.display = 'block';
        }
    } catch (error) {
        console.error('Error showing data status:', error);
    }
}

// Legacy function compatibility
function saveDataToStorage() {
    console.log('Using legacy saveDataToStorage, switching to Firebase version...');
    return saveDataFromStorage();
}

// Load data from localStorage (fallback)
function loadDataFromStorage() {
    return loadDataFromLocalStorage();
}

// Handle file upload from input change
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (file) {
        processFile(file);
    }
}

// File selection handler
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

// Process the selected Excel file
function processFile(file) {
    // Validate file type
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel' // .xls
    ];
    
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Please select a valid Excel file (.xlsx or .xls)');
        return;
    }
    
    // Show loading state
    showLoading(true);
    showFileName(file.name);
    hideError();
    
    // Read and process the file
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = e.target.result;
            workbook = XLSX.read(data, { type: 'binary' });
            
            // Automatically extract and display data
            setTimeout(async () => {
                await mergeNewDataWithExisting();
                showLoading(false);
            }, 500);
            
        } catch (error) {
            console.error('Error reading Excel file:', error);
            showError('Error reading Excel file. Please make sure it\'s a valid Excel file.');
            showLoading(false);
        }
    };
    
    reader.onerror = function() {
        showError('Error reading file');
        showLoading(false);
    };
    
    reader.readAsBinaryString(file);
}

// Merge new data with existing data
async function mergeNewDataWithExisting() {
    // Extract new data from the uploaded file
    const newData = [];
    workbook.SheetNames.forEach(sheetName => {
        const sheetData = extractDataFromFingerprintData(sheetName);
        newData.push(...sheetData);
    });
    
    if (newData.length === 0) {
        showError('No valid fingerprint data found in the uploaded file.');
        return;
    }
    
    // If no existing data, use new data as is
    if (extractedData.length === 0) {
        extractedData = newData;
    } else {
        // Merge with existing data
        mergeEmployeeData(newData);
    }
    
    // Update months list
    updateMonthsList();
    
    // Save to storage
    await saveDataFromStorage();
    
    // Display updated data
    displayDataWithFilters();
    showDataStatus();
}

// Merge employee data (combine attendance records for same employees)
function mergeEmployeeData(newData) {
    newData.forEach(newEmployee => {
        // Find existing employee by name and employee ID
        const existingIndex = extractedData.findIndex(existing => 
            existing.name.toLowerCase() === newEmployee.name.toLowerCase() &&
            existing.employeeId === newEmployee.employeeId
        );
        
        if (existingIndex !== -1) {
            // Employee exists, merge attendance data
            const existing = extractedData[existingIndex];
            
            // Add new attendance records that don't already exist
            newEmployee.attendanceData.forEach(newRecord => {
                const recordExists = existing.attendanceData.some(existingRecord => 
                    existingRecord.date === newRecord.date &&
                    existingRecord.month === newRecord.month &&
                    existingRecord.year === newRecord.year
                );
                
                if (!recordExists) {
                    existing.attendanceData.push(newRecord);
                }
            });
            
            // Sort attendance data by date
            existing.attendanceData.sort((a, b) => {
                if (a.fullDate && b.fullDate) {
                    return a.fullDate.getTime() - b.fullDate.getTime();
                }
                return 0;
            });
            
            // Update employee info if needed
            if (newEmployee.department && !existing.department) {
                existing.department = newEmployee.department;
            }
            
        } else {
            // New employee, add to extracted data
            extractedData.push(newEmployee);
        }
    });
}

// Update months list from all attendance data
function updateMonthsList() {
    const monthsFromAttendance = new Set();
    extractedData.forEach(employee => {
        employee.attendanceData.forEach(record => {
            if (record.month && record.year) {
                monthsFromAttendance.add(`${record.year}-${String(record.month).padStart(2, '0')}`);
            }
        });
        // Also include employee record month if available
        if (employee.month && employee.year) {
            monthsFromAttendance.add(`${employee.year}-${String(employee.month).padStart(2, '0')}`);
        }
    });
    
    allMonths = [...monthsFromAttendance].sort();
}

// Clear all saved data
function clearAllData() {
    if (confirm('Are you sure you want to clear all saved data? This action cannot be undone.')) {
        // Clear localStorage
        Object.values(STORAGE_KEYS).forEach(key => {
            localStorage.removeItem(key);
        });
        
        // Clear memory
        extractedData = [];
        allMonths = [];
        holidays = new Set();
        
        // Clear UI
        const summaryDiv = document.getElementById('namesSummary');
        if (summaryDiv) {
            summaryDiv.remove();
        }
        
        const fileName = document.getElementById('fileName');
        if (fileName) {
            fileName.style.display = 'none';
        }
        
        showError('All data has been cleared.');
        
        // Hide error after 3 seconds
        setTimeout(() => {
            hideError();
        }, 3000);
    }
}

// Hide error message
function hideError() {
    const error = document.getElementById('error');
    if (error) {
        error.style.display = 'none';
    }
}

// Show/hide loading state
function showLoading(show) {
    const loading = document.getElementById('loading');
    if (loading) {
        loading.style.display = show ? 'flex' : 'none';
    }
}

// Show file name
function showFileName(name) {
    const fileName = document.getElementById('fileName');
    if (fileName) {
        fileName.textContent = `Selected: ${name}`;
        fileName.style.display = 'block';
    }
}

// Show error message
function showError(message) {
    const error = document.getElementById('error');
    if (error) {
        error.textContent = message;
        error.style.display = 'block';
    }
}

// Load data from localStorage
function loadDataFromStorage() {
    try {
        // Load extracted data
        const savedData = localStorage.getItem(STORAGE_KEYS.EXTRACTED_DATA);
        if (savedData) {
            extractedData = JSON.parse(savedData);
            // Convert date strings back to Date objects
            extractedData.forEach(employee => {
                employee.attendanceData.forEach(record => {
                    if (record.fullDate && typeof record.fullDate === 'string') {
                        record.fullDate = new Date(record.fullDate);
                    }
                    if (record.startDate && typeof record.startDate === 'string') {
                        record.startDate = new Date(record.startDate);
                    }
                    if (record.endDate && typeof record.endDate === 'string') {
                        record.endDate = new Date(record.endDate);
                    }
                });
            });
        }
        
        // Load months
        const savedMonths = localStorage.getItem(STORAGE_KEYS.ALL_MONTHS);
        if (savedMonths) {
            allMonths = JSON.parse(savedMonths);
        }
        
        // Load holidays
        const savedHolidays = localStorage.getItem(STORAGE_KEYS.HOLIDAYS);
        if (savedHolidays) {
            holidays = new Set(JSON.parse(savedHolidays));
        }
        
        // Display data if available
        if (extractedData.length > 0) {
            displayDataWithFilters();
            showDataStatus();
        }
        
    } catch (error) {
        console.error('Error loading data from storage:', error);
        showError('Error loading saved data. Starting fresh.');
    }
}

// Save data to localStorage
function saveDataToStorage() {
    try {
        localStorage.setItem(STORAGE_KEYS.EXTRACTED_DATA, JSON.stringify(extractedData));
        localStorage.setItem(STORAGE_KEYS.ALL_MONTHS, JSON.stringify(allMonths));
        localStorage.setItem(STORAGE_KEYS.HOLIDAYS, JSON.stringify([...holidays]));
        localStorage.setItem(STORAGE_KEYS.LAST_UPDATED, new Date().toISOString());
        
        console.log('Data saved successfully');
    } catch (error) {
        console.error('Error saving data to storage:', error);
        showError('Warning: Could not save data to browser storage.');
    }
}

// Show data status
function showDataStatus() {
    const lastUpdated = localStorage.getItem(STORAGE_KEYS.LAST_UPDATED);
    if (lastUpdated) {
        const date = new Date(lastUpdated);
        const fileName = document.getElementById('fileName');
        if (fileName) {
            fileName.innerHTML = `
                <i class="bi bi-database"></i> 
                Saved Data Available (${extractedData.length} records) - 
                Last updated: ${date.toLocaleDateString()} ${date.toLocaleTimeString()}
                <button onclick="clearAllData()" class="btn btn-sm btn-outline-danger ms-2 mobile-hide" style="background: transparent; border: 1px solid #ef4444; color: #ef4444; padding: 0.25rem 0.5rem; border-radius: 6px; font-size: 0.75rem;">
                    <i class="bi bi-trash"></i> Clear All Data
                </button>
            `;
            fileName.style.display = 'block';
        }
    }
}

// Remove old functions that are no longer needed

// Clear all data
function clearData() {
    workbook = null;
    currentSheet = null;
    extractedData = [];
    allMonths = [];
    
    const fileInput = document.getElementById('fileInput');
    if (fileInput) {
        fileInput.value = '';
    }
    
    // Hide UI elements
    hideError();
    showLoading(false);
    
    const fileName = document.getElementById('fileName');
    if (fileName) {
        fileName.style.display = 'none';
    }
    
    // Remove results
    const summaryDiv = document.getElementById('namesSummary');
    if (summaryDiv) {
        summaryDiv.remove();
    }
}

// Extract names, dates, and times from fingerprint machine data
function extractDataFromFingerprintData(sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const extractedRecords = [];
    
    for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        let currentRecord = null;
        
        // Look for rows that contain "Name:" pattern
        for (let j = 0; j < row.length; j++) {
            const cell = row[j];
            if (typeof cell === 'string' && cell.includes('Name:')) {
                // Extract name
                const nameMatch = cell.match(/Name:([^ID]*)/);
                let name = '';
                if (nameMatch) {
                    name = nameMatch[1].trim();
                    
                    // Clean up the name - remove any duplicate parts
                    // If name contains the same word repeated, keep only one
                    const words = name.split(/\s+/);
                    const uniqueWords = [];
                    for (const word of words) {
                        if (word && !uniqueWords.includes(word.toLowerCase())) {
                            uniqueWords.push(word);
                        }
                    }
                    name = uniqueWords.join(' ');
                    
                    // Also check columns 4-7 (0-indexed) for additional name data
                    for (let k = 4; k <= 7; k++) {
                        if (row[k] && typeof row[k] === 'string' && !row[k].includes('ID:') && !row[k].includes('Name:')) {
                            const additionalName = row[k].trim();
                            // Only add if it's not already in the name
                            if (additionalName && !name.toLowerCase().includes(additionalName.toLowerCase())) {
                                name += ' ' + additionalName;
                            }
                        }
                    }
                }
                
                // Extract date range
                let dateRange = '';
                const dateMatch = cell.match(/Date:([^\s]*)/);
                if (dateMatch) {
                    dateRange = dateMatch[1].trim();
                }
                
                // Extract ID
                let employeeId = '';
                const idMatch = cell.match(/ID:([^\s]*)/);
                if (idMatch) {
                    employeeId = idMatch[1].trim();
                }
                
                // Extract department
                let department = '';
                const deptMatch = cell.match(/Dept:([^Name]*)/);
                if (deptMatch) {
                    department = deptMatch[1].trim();
                }
                
                if (name && name !== '') {
                    currentRecord = {
                        name: name.replace(/\s+/g, ' ').trim(),
                        employeeId: employeeId,
                        department: department,
                        dateRange: dateRange,
                        startDate: null,
                        endDate: null,
                        month: null,
                        year: null,
                        row: i + 1,
                        sheet: sheetName,
                        attendanceData: []
                    };
                    
                    // Parse date range (format: 25.06.01~25.06.30)
                    if (dateRange) {
                        const dates = dateRange.split('~');
                        if (dates.length === 2) {
                            try {
                                // Parse start date (25.06.01 format - YY.MM.DD)
                                const startParts = dates[0].split('.');
                                if (startParts.length === 3) {
                                    const year = 2000 + parseInt(startParts[0]);
                                    const month = parseInt(startParts[1]);
                                    const day = parseInt(startParts[2]);
                                    currentRecord.startDate = new Date(year, month - 1, day);
                                    currentRecord.month = month;
                                    currentRecord.year = year;
                                }
                                
                                // Parse end date
                                const endParts = dates[1].split('.');
                                if (endParts.length === 3) {
                                    const year = 2000 + parseInt(endParts[0]);
                                    const month = parseInt(endParts[1]);
                                    const day = parseInt(endParts[2]);
                                    currentRecord.endDate = new Date(year, month - 1, day);
                                }
                            } catch (e) {
                                console.warn('Date parsing error:', e);
                            }
                        }
                    }
                    
                    // Extract attendance table data
                    currentRecord.attendanceData = extractAttendanceTable(jsonData, i + 3, currentRecord.year); // Skip 3 rows to get to attendance table
                    
                    extractedRecords.push(currentRecord);
                }
                break;
            }
        }
    }
    
    return extractedRecords;
}

// Extract attendance table data for an individual
function extractAttendanceTable(jsonData, startRow, year) {
    const attendanceRecords = [];
    
    // Look for the attendance table starting from startRow
    for (let i = startRow; i < Math.min(startRow + 50, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row || row.length === 0) continue;
        
        // Stop if we hit another employee's data
        const rowText = row.join(' ');
        if (rowText.includes('Name:') || rowText.includes('Dept:')) {
            break;
        }
        
        // Look for date pattern (MM.DD format)
        const dateCell = row[0];
        if (typeof dateCell === 'string' && dateCell.match(/^\d{2}\.\d{2}$/)) {
            const [month, day] = dateCell.split('.').map(Number);
            const date = new Date(year || 2025, month - 1, day); // month - 1 because JS months are 0-indexed
            const dayOfWeek = row[1];
            
            // Extract times from different shifts
            const attendanceRecord = {
                date: dateCell,
                fullDate: date,
                dayOfWeek: dayOfWeek,
                month: month, // Store the actual month number (1-12)
                year: year || 2025,
                morning: {
                    in: extractTimeFromCell(row[2]),
                    out: extractTimeFromCell(row[3])
                },
                afternoon: {
                    in: extractTimeFromCell(row[4]),
                    out: extractTimeFromCell(row[5])
                },
                evening: {
                    in: extractTimeFromCell(row[6]),
                    out: extractTimeFromCell(row[7])
                }
            };
            
            // Check if there's a second set of columns (right side of the table)
            if (row.length > 8) {
                const rightDate = row[8];
                if (typeof rightDate === 'string' && rightDate.match(/^\d{2}\.\d{2}$/)) {
                    const [rightMonth, rightDay] = rightDate.split('.').map(Number);
                    const rightFullDate = new Date(year || 2025, rightMonth - 1, rightDay);
                    const rightDayOfWeek = row[9];
                    
                    const rightRecord = {
                        date: rightDate,
                        fullDate: rightFullDate,
                        dayOfWeek: rightDayOfWeek,
                        month: rightMonth, // Store the actual month number (1-12)
                        year: year || 2025,
                        morning: {
                            in: extractTimeFromCell(row[10]),
                            out: extractTimeFromCell(row[11])
                        },
                        afternoon: {
                            in: extractTimeFromCell(row[12]),
                            out: extractTimeFromCell(row[13])
                        },
                        evening: {
                            in: extractTimeFromCell(row[14]),
                            out: extractTimeFromCell(row[15])
                        }
                    };
                    
                    attendanceRecords.push(rightRecord);
                }
            }
            
            attendanceRecords.push(attendanceRecord);
        }
    }
    
    return attendanceRecords;
}

// Extract time from cell (handles formats like "06:42", "07:07*", etc.)
function extractTimeFromCell(cell) {
    if (!cell || cell === '') return null;
    
    const cellStr = cell.toString().trim();
    // Match time pattern and remove any trailing characters like *
    const timeMatch = cellStr.match(/^(\d{1,2}:\d{2})/);
    return timeMatch ? timeMatch[1] : null;
}

// Display extracted data with filtering
function displayExtractedNames() {
    // This function is now replaced by mergeNewDataWithExisting
    // but kept for compatibility
    mergeNewDataWithExisting();
}

// Display data with month filtering interface
function displayDataWithFilters(selectedMonth = 'all') {
    // Filter data by selected month
    let filteredData = extractedData;
    if (selectedMonth !== 'all') {
        const [year, month] = selectedMonth.split('-');
        filteredData = extractedData.filter(record => {
            // Check if employee has attendance data for the selected month
            const hasAttendanceInMonth = record.attendanceData.some(attendance => 
                attendance.year == year && attendance.month == month
            );
            // Also check the employee record month (fallback)
            const employeeRecordMatch = record.year == year && record.month == month;
            
            return hasAttendanceInMonth || employeeRecordMatch;
        });
        
        // Also filter the attendance data within each employee record
        filteredData = filteredData.map(record => ({
            ...record,
            attendanceData: record.attendanceData.filter(attendance => 
                attendance.year == year && attendance.month == month
            )
        }));
    }
    
    // Create month selector
    const monthOptions = allMonths.map(month => {
        const [year, monthNum] = month.split('-');
        const monthName = new Date(year, monthNum - 1).toLocaleString('default', { month: 'long', year: 'numeric' });
        return `<option value="${month}" ${selectedMonth === month ? 'selected' : ''}>${monthName}</option>`;
    }).join('');
    
    // Create the summary view
    const summaryHtml = `
        <div class="main-card">
            <div class="main-header">
                <div class="header-content">
                    <div class="header-left">
                        <h5 class="main-title">
                            <i class="bi bi-person-lines-fill me-2"></i>Fingerprint Data Analysis
                        </h5>
                    </div>
                    <div class="header-right">
                        <div class="search-container">
                            <input type="text" id="prefectSearch" class="search-input" placeholder="Search prefects...">
                            <i class="bi bi-search search-icon"></i>
                        </div>
                        <label for="monthFilter" class="filter-label">Filter by Month:</label>
                        <select id="monthFilter" class="filter-select" onchange="filterByMonth(this.value)">
                            <option value="all">All Months</option>
                            ${monthOptions}
                        </select>
                    </div>
                </div>
            </div>
            <div class="main-body">
                <div class="stats-grid" data-animate="fadeInUp" data-delay="0">
                    <div class="stat-item" data-animate="fadeInUp" data-delay="0">
                        <div class="stat-icon bg-primary">
                            <i class="bi bi-files"></i>
                        </div>
                        <div class="stat-content">
                            <h3 class="stat-number text-primary">${filteredData.length}</h3>
                            <p class="stat-label">Total Records</p>
                        </div>
                    </div>
                    <div class="stat-item" data-animate="fadeInUp" data-delay="100">
                        <div class="stat-icon bg-success">
                            <i class="bi bi-people"></i>
                        </div>
                        <div class="stat-content">
                            <h3 class="stat-number text-success">${[...new Set(filteredData.map(r => r.name))].length}</h3>
                            <p class="stat-label">Unique Prefects</p>
                        </div>
                    </div>
                    <div class="stat-item" data-animate="fadeInUp" data-delay="200">
                        <div class="stat-icon bg-info">
                            <i class="bi bi-calendar-week"></i>
                        </div>
                        <div class="stat-content">
                            <h3 class="stat-number text-info">${(() => {
                                if (selectedMonth === 'all') {
                                    // Calculate total working days across all months
                                    const allWorkingDays = new Set();
                                    filteredData.forEach(employee => {
                                        employee.attendanceData.forEach(record => {
                                            const dateString = record.fullDate ? record.fullDate.toISOString().split('T')[0] : null;
                                            const isWeekend = record.dayOfWeek === 'SAT' || record.dayOfWeek === 'SUN';
                                            const isHolidayDate = dateString && isHoliday(dateString);
                                            if (!isWeekend && !isHolidayDate && dateString) {
                                                allWorkingDays.add(dateString);
                                            }
                                        });
                                    });
                                    return allWorkingDays.size;
                                } else {
                                    // Calculate working days for selected month
                                    const [year, month] = selectedMonth.split('-');
                                    const daysInMonth = new Date(year, month, 0).getDate();
                                    let workingDays = 0;
                                    
                                    for (let day = 1; day <= daysInMonth; day++) {
                                        const checkDate = new Date(year, month - 1, day);
                                        const dayOfWeek = checkDate.getDay(); // 0 = Sunday, 6 = Saturday
                                        const dateString = checkDate.toISOString().split('T')[0];
                                        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
                                        const isHolidayDate = isHoliday(dateString);
                                        
                                        if (!isWeekend && !isHolidayDate) {
                                            workingDays++;
                                        }
                                    }
                                    return workingDays;
                                }
                            })()}</h3>
                            <p class="stat-label">Working Days</p>
                        </div>
                    </div>
                </div>
                
                <div class="table-section" data-animate="fadeInUp" data-delay="300">
                    <div class="section-header">
                        <h6 class="section-title">Prefect Records ${selectedMonth !== 'all' ? `- ${new Date(selectedMonth + '-01').toLocaleString('default', { month: 'long', year: 'numeric' })}` : ''}</h6>
                        <div class="action-buttons">
                            <button class="btn-action btn-info mobile-hide" onclick="showHolidayCalendar('${selectedMonth}')">
                                <i class="bi bi-calendar-event"></i> Manage Holidays
                            </button>
                            <button class="btn-action btn-warning" onclick="showAnalysis('${selectedMonth}')">
                                <i class="bi bi-bar-chart"></i> Analysis
                            </button>
                            <button class="btn-action btn-success" onclick="downloadDetailedCSV('${selectedMonth}')">
                                <i class="bi bi-download"></i> Download Report
                            </button>
                            <button class="btn-action btn-outline mobile-hide" onclick="downloadSummaryCSV('${selectedMonth}')">
                                <i class="bi bi-file-earmark-spreadsheet"></i> Summary
                            </button>
                            <button class="btn-action btn-secondary mobile-hide" onclick="exportAllData()">
                                <i class="bi bi-box-arrow-up"></i> Export All Data
                            </button>
                        </div>
                    </div>
                    
                    <div class="table-container">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>#</th>
                                    <th>Prefect Name</th>
                                    <th class="mobile-hide">Morning Entrance Times</th>
                                    <th>Working Days</th>
                                    <th>Sheet</th>
                                    <th>Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${filteredData.map((record, index) => `
                                    <tr class="table-row-fade" style="animation-delay: ${400 + (index * 50)}ms">
                                        <td><span class="row-number">${index + 1}</span></td>
                                        <td><span class="prefect-name">${record.name}</span></td>
                                        <td class="mobile-hide">
                                            ${record.attendanceData.filter(r => r.morning && r.morning.in).length > 0 ? `
                                                <div class="time-badges">
                                                    ${record.attendanceData.filter(r => r.morning && r.morning.in).slice(0, 5).map(r => 
                                                        `<span class="time-badge">${r.morning.in}</span>`
                                                    ).join('')}
                                                    ${record.attendanceData.filter(r => r.morning && r.morning.in).length > 5 ? 
                                                        `<span class="more-badge">+${record.attendanceData.filter(r => r.morning && r.morning.in).length - 5} more</span>` : ''
                                                    }
                                                </div>
                                            ` : '<span class="no-data">No morning entries</span>'}
                                        </td>
                                        <td>
                                            <span class="days-badge">${getPresentDaysCount(record.attendanceData)} days</span>
                                            <div class="days-detail">
                                                Working: ${getWorkingDaysCount(record.attendanceData)} | Holidays: ${record.attendanceData.filter(r => {
                                                    const dateString = r.fullDate ? r.fullDate.toISOString().split('T')[0] : null;
                                                    return dateString && isHoliday(dateString);
                                                }).length}
                                            </div>
                                        </td>
                                        <td><span class="sheet-name">${record.sheet}</span></td>
                                        <td>
                                            <button class="btn-view" data-employee-name="${record.name}" data-selected-month="${selectedMonth}" title="View Details">
                                                <i class="bi bi-eye"></i>
                                            </button>
                                        </td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
        
        <style>
            /* Global Styles */
            * {
                box-sizing: border-box;
            }
            
            body {
                background: #0a0a0a;
                color: #e2e8f0;
            }
            
            /* Main Card */
            .main-card {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 20px;
                box-shadow: 0 4px 30px rgba(0,0,0,0.5);
                overflow: hidden;
                margin-top: 2rem;
                animation: slideUp 0.8s cubic-bezier(0.4, 0, 0.2, 1);
            }
            
            .main-header {
                background: linear-gradient(135deg, #065f46 0%, #022c22 100%);
                padding: 2rem;
                color: #10b981;
                border-bottom: 1px solid #2d2d2d;
            }
            
            .header-content {
                display: flex;
                justify-content: space-between;
                align-items: center;
                flex-wrap: wrap;
                gap: 1rem;
            }
            
            .main-title {
                font-size: 1.5rem;
                font-weight: 300;
                margin: 0;
                display: flex;
                align-items: center;
                color: #10b981;
            }
            
            .header-right {
                display: flex;
                align-items: center;
                gap: 1rem;
            }
            
            .header-right {
                display: flex;
                align-items: center;
                gap: 1rem;
                flex-wrap: wrap;
            }
            
            .search-container {
                position: relative;
                display: flex;
                align-items: center;
            }
            
            .search-input {
                background: #0f0f0f;
                border: 1px solid #2d2d2d;
                border-radius: 8px;
                padding: 0.5rem 2.5rem 0.5rem 1rem;
                color: #e2e8f0;
                font-size: 0.875rem;
                min-width: 200px;
                transition: all 0.3s ease;
            }
            
            .search-input:focus {
                outline: none;
                border-color: #10b981;
                box-shadow: 0 0 0 2px rgba(16, 185, 129, 0.2);
            }
            
            .search-input::placeholder {
                color: #6b7280;
            }
            
            .search-icon {
                position: absolute;
                right: 0.75rem;
                color: #6b7280;
                pointer-events: none;
            }
            
            .search-result-indicator {
                font-size: 0.75rem;
                color: #10b981;
                margin-top: 0.25rem;
                text-align: center;
                display: none;
            }
            
            .filter-label {
                font-weight: 400;
                margin: 0;
                opacity: 0.9;
                color: #a7f3d0;
            }
            
            .filter-select {
                background: rgba(16, 185, 129, 0.15);
                border: 1px solid #065f46;
                border-radius: 10px;
                padding: 0.5rem 1rem;
                color: #10b981;
                font-size: 0.875rem;
                min-width: 200px;
                backdrop-filter: blur(10px);
            }
            
            .filter-select option {
                background: #1a1a1a;
                color: #10b981;
            }
            
            .main-body {
                padding: 2rem;
                background: #0f0f0f;
            }
            
            /* Stats Grid */
            .stats-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
                gap: 1.5rem;
                margin-bottom: 2rem;
            }
            
            .stat-item {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 16px;
                padding: 1.5rem;
                box-shadow: 0 2px 20px rgba(0,0,0,0.3);
                transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                position: relative;
                overflow: hidden;
                display: flex;
                align-items: center;
                gap: 1rem;
            }
            
            .stat-item::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                height: 3px;
                background: linear-gradient(90deg, #10b981, #065f46);
                transform: scaleX(0);
                transition: transform 0.3s ease;
            }
            
            .stat-item:hover {
                transform: translateY(-4px);
                box-shadow: 0 8px 40px rgba(16, 185, 129, 0.2);
                border-color: #065f46;
            }
            
            .stat-item:hover::before {
                transform: scaleX(1);
            }
            
            .stat-icon {
                width: 60px;
                height: 60px;
                border-radius: 15px;
                display: flex;
                align-items: center;
                justify-content: center;
                flex-shrink: 0;
                background: linear-gradient(135deg, #10b981, #065f46);
            }
            
            .stat-icon i {
                font-size: 24px;
                color: white;
            }
            
            .stat-content {
                flex: 1;
            }
            
            .stat-number {
                font-size: 2.5rem;
                font-weight: 300;
                line-height: 1;
                margin-bottom: 0.5rem;
                color: #10b981;
            }
            
            .stat-label {
                font-size: 1rem;
                font-weight: 500;
                color: #a7f3d0;
                margin: 0;
            }
            
            /* Table Section */
            .table-section {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 16px;
                box-shadow: 0 2px 20px rgba(0,0,0,0.3);
                overflow: hidden;
            }
            
            .section-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 1.5rem;
                border-bottom: 1px solid #2d2d2d;
                flex-wrap: wrap;
                gap: 1rem;
                background: #0f0f0f;
            }
            
            .section-title {
                font-size: 1.125rem;
                font-weight: 500;
                color: #10b981;
                margin: 0;
            }
            
            .action-buttons {
                display: flex;
                gap: 0.5rem;
                flex-wrap: wrap;
            }
            
            .btn-action {
                border: none;
                border-radius: 8px;
                padding: 0.5rem 1rem;
                font-size: 0.875rem;
                font-weight: 500;
                display: flex;
                align-items: center;
                gap: 0.5rem;
                transition: all 0.2s ease;
                text-decoration: none;
                cursor: pointer;
            }
            
            .btn-info {
                background: linear-gradient(135deg, #10b981, #065f46);
                color: white;
                border: 1px solid #065f46;
            }
            
            .btn-warning {
                background: linear-gradient(135deg, #059669, #047857);
                color: white;
                border: 1px solid #047857;
            }
            
            .btn-success {
                background: linear-gradient(135deg, #10b981, #059669);
                color: white;
                border: 1px solid #059669;
            }
            
            .btn-outline {
                background: transparent;
                border: 1px solid #2d2d2d;
                color: #a7f3d0;
            }
            
            .btn-secondary {
                background: linear-gradient(135deg, #374151, #1f2937);
                color: #10b981;
                border: 1px solid #374151;
            }
            
            .btn-action:hover {
                transform: translateY(-1px);
                box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
            }
            
            .btn-outline:hover {
                background: #1a1a1a;
                border-color: #10b981;
                color: #10b981;
            }
            
            /* Data Table */
            .table-container {
                overflow-x: auto;
            }
            
            .data-table {
                width: 100%;
                border-collapse: collapse;
            }
            
            .data-table th {
                background: #0f0f0f;
                padding: 1rem;
                font-weight: 500;
                font-size: 0.875rem;
                color: #10b981;
                text-align: left;
                border: none;
                border-bottom: 1px solid #2d2d2d;
                position: sticky;
                top: 0;
                z-index: 10;
            }
            
            .data-table td {
                padding: 1rem;
                border: none;
                border-bottom: 1px solid #2d2d2d;
                font-size: 0.875rem;
                vertical-align: middle;
                color: #e2e8f0;
            }
            
            .data-table tr:hover {
                background-color: #0f0f0f;
            }
            
            .row-number {
                background: #374151;
                color: #10b981;
                padding: 0.25rem 0.5rem;
                border-radius: 6px;
                font-weight: 500;
                font-size: 0.75rem;
            }
            
            .prefect-name {
                font-weight: 600;
                color: #10b981;
            }
            
            .time-badges {
                display: flex;
                flex-wrap: wrap;
                gap: 0.25rem;
            }
            
            .time-badge {
                background: #065f46;
                color: #a7f3d0;
                padding: 0.25rem 0.5rem;
                border-radius: 6px;
                font-weight: 500;
                font-size: 0.75rem;
                border: 1px solid #047857;
            }
            
            .more-badge {
                background: #374151;
                color: #9ca3af;
                padding: 0.25rem 0.5rem;
                border-radius: 6px;
                font-weight: 500;
                font-size: 0.75rem;
                border: 1px solid #4b5563;
            }
            
            .no-data {
                color: #6b7280;
                font-style: italic;
            }
            
            .days-badge {
                background: #065f46;
                color: #a7f3d0;
                padding: 0.25rem 0.75rem;
                border-radius: 8px;
                font-weight: 600;
                font-size: 0.875rem;
                border: 1px solid #047857;
            }
            
            .days-detail {
                font-size: 0.75rem;
                color: #9ca3af;
                margin-top: 0.25rem;
            }
            
            .sheet-name {
                color: #9ca3af;
                font-size: 0.75rem;
            }
            
            .btn-view {
                background: transparent;
                border: 1px solid #2d2d2d;
                border-radius: 6px;
                padding: 0.5rem;
                color: #10b981;
                transition: all 0.2s ease;
                cursor: pointer;
            }
            
            .btn-view:hover {
                background: #065f46;
                border-color: #10b981;
                color: white;
            }
            
            /* Animations */
            @keyframes slideUp {
                from { 
                    opacity: 0; 
                    transform: translateY(30px); 
                }
                to { 
                    opacity: 1; 
                    transform: translateY(0); 
                }
            }
            
            @keyframes fadeIn {
                from { opacity: 0; }
                to { opacity: 1; }
            }            @keyframes fadeInUp {
                from {
                    opacity: 0;
                    transform: translateY(20px);
                }
                to {
                    opacity: 1;
                    transform: translateY(0);
                }
            }
            
            @keyframes fadeIn {
                from {
                    opacity: 0;
                    transform: translateX(-10px);
                }
                to {
                    opacity: 1;
                    transform: translateX(0);
                }
            }
            
            [data-animate="fadeInUp"] {
                animation: fadeInUp 0.6s cubic-bezier(0.4, 0, 0.2, 1) forwards;
                opacity: 0;
            }
            
            .table-row-fade {
                animation: fadeIn 0.4s cubic-bezier(0.4, 0, 0.2, 1) forwards;
                opacity: 0;
            }
            
            /* Animation delays */
            [data-delay="0"] { animation-delay: 0ms; }
            [data-delay="100"] { animation-delay: 100ms; }
            [data-delay="200"] { animation-delay: 200ms; }
            [data-delay="300"] { animation-delay: 300ms; }
            
            /* Mobile Hide Utility */
            @media (max-width: 768px) {
                .mobile-hide {
                    display: none !important;
                }
            }
            
            /* Mobile-First Responsive Design */
            
            /* Base Mobile Styles (320px+) */
            @media (max-width: 480px) {
                /* Main Layout */
                .main-card {
                    margin-top: 1rem;
                    border-radius: 16px;
                }
                
                .main-header {
                    padding: 1.5rem;
                }
                
                .main-title {
                    font-size: 1.25rem;
                }
                
                .main-body {
                    padding: 1rem;
                }
                
                /* Header adjustments */
                .header-content {
                    flex-direction: column;
                    align-items: stretch;
                    gap: 1rem;
                }
                
                .header-right {
                    flex-direction: column;
                    align-items: stretch;
                    gap: 0.5rem;
                }
                
                .filter-select {
                    min-width: 100%;
                    text-align: center;
                }
                
                /* Stats Grid - Single Column */
                .stats-grid {
                    grid-template-columns: 1fr;
                    gap: 1rem;
                    margin-bottom: 1.5rem;
                }
                
                .stat-item {
                    padding: 1.25rem;
                    border-radius: 12px;
                }
                
                .stat-icon {
                    width: 50px;
                    height: 50px;
                }
                
                .stat-number {
                    font-size: 2rem;
                }
                
                /* Section Header */
                .section-header {
                    flex-direction: column;
                    align-items: stretch;
                    padding: 1rem;
                    gap: 1rem;
                }
                
                .section-title {
                    text-align: center;
                    font-size: 1rem;
                }
                
                /* Action Buttons - Stack vertically */
                .action-buttons {
                    display: grid;
                    grid-template-columns: 1fr 1fr;
                    gap: 0.5rem;
                    width: 100%;
                }
                
                .btn-action {
                    justify-content: center;
                    padding: 0.75rem 0.5rem;
                    font-size: 0.75rem;
                    border-radius: 6px;
                }
                
                .btn-action i {
                    font-size: 0.875rem;
                }
                
                /* Table Responsive */
                .table-container {
                    margin: 0 -1rem;
                    border-radius: 0;
                }
                
                .data-table {
                    font-size: 0.75rem;
                }
                
                .data-table th,
                .data-table td {
                    padding: 0.5rem 0.25rem;
                    white-space: nowrap;
                }
                
                .data-table th:first-child,
                .data-table td:first-child {
                    padding-left: 1rem;
                }
                
                .data-table th:last-child,
                .data-table td:last-child {
                    padding-right: 1rem;
                }
                
                /* Hide less important columns on very small screens */
                .data-table th:nth-child(5),
                .data-table td:nth-child(5) {
                    display: none;
                }
                
                /* Compact badges */
                .time-badges {
                    flex-direction: column;
                    gap: 0.125rem;
                }
                
                .time-badge {
                    font-size: 0.625rem;
                    padding: 0.125rem 0.375rem;
                }
                
                .more-badge {
                    font-size: 0.625rem;
                    padding: 0.125rem 0.375rem;
                }
                
                .days-badge {
                    font-size: 0.75rem;
                    padding: 0.125rem 0.5rem;
                }
                
                .days-detail {
                    font-size: 0.625rem;
                }
                
                .btn-view {
                    padding: 0.375rem;
                }
                
                .btn-view i {
                    font-size: 0.875rem;
                }
            }
            
            /* Tablet Styles (481px - 768px) */
            @media (min-width: 481px) and (max-width: 768px) {
                .main-body {
                    padding: 1.5rem;
                }
                
                .header-content {
                    flex-direction: column;
                    align-items: stretch;
                }
                
                .stats-grid {
                    grid-template-columns: repeat(2, 1fr);
                }
                
                .action-buttons {
                    justify-content: center;
                    flex-wrap: wrap;
                }
                
                .btn-action {
                    flex: 1;
                    min-width: 120px;
                    justify-content: center;
                }
                
                .section-header {
                    flex-direction: column;
                    align-items: center;
                    text-align: center;
                }
            }
            
            /* Large Tablet/Small Desktop (769px - 1024px) */
            @media (min-width: 769px) and (max-width: 1024px) {
                .stats-grid {
                    grid-template-columns: repeat(3, 1fr);
                }
                
                .action-buttons {
                    flex-wrap: wrap;
                    justify-content: center;
                }
            }
        </style>
    `;
    
    // Insert the summary in the main container
    let summaryDiv = document.getElementById('namesSummary');
    
    if (!summaryDiv) {
        summaryDiv = document.createElement('div');
        summaryDiv.id = 'namesSummary';
        // Find results container or create one
        const resultsContainer = document.getElementById('results');
        if (resultsContainer) {
            resultsContainer.appendChild(summaryDiv);
        } else {
            document.body.appendChild(summaryDiv);
        }
    }
    
    summaryDiv.innerHTML = summaryHtml;
    
    // Add event listeners for view buttons
    setTimeout(() => {
        document.querySelectorAll('.btn-view').forEach(button => {
            button.addEventListener('click', function() {
                const employeeName = this.getAttribute('data-employee-name');
                const selectedMonth = this.getAttribute('data-selected-month');
                console.log('Opening details for:', employeeName, selectedMonth);
                showEmployeeDetails(employeeName, selectedMonth);
            });
        });
        
        // Set up search functionality
        const searchInput = document.getElementById('prefectSearch');
        if (searchInput) {
            // Remove any existing event listeners
            searchInput.removeEventListener('input', debouncedSearch);
            searchInput.removeEventListener('keyup', debouncedSearch);
            
            // Add new event listeners with debouncing
            searchInput.addEventListener('input', debouncedSearch);
            searchInput.addEventListener('keyup', debouncedSearch);
            
            // Trigger search if there's already a value
            if (searchInput.value) {
                searchPrefects();
            }
        }
    }, 100);
}

// Filter data by month
function filterByMonth(selectedMonth) {
    displayDataWithFilters(selectedMonth);
}

// Show detailed view for an employee
function showEmployeeDetails(employeeName, selectedMonth) {
    console.log('showEmployeeDetails called with:', employeeName, selectedMonth);
    console.log('extractedData length:', extractedData.length);
    
    // Find the employee directly from extractedData (don't filter the main array)
    const employee = extractedData.find(record => record.name === employeeName);
    if (!employee) {
        console.log('Employee not found:', employeeName);
        return;
    }
    
    console.log('Employee found:', employee.name, 'with', employee.attendanceData.length, 'attendance records');
    
    // Filter attendance data by month if needed
    let attendanceToShow = employee.attendanceData;
    if (selectedMonth !== 'all') {
        const [year, month] = selectedMonth.split('-');
        attendanceToShow = employee.attendanceData.filter(record => 
            record.fullDate && record.fullDate.getFullYear() == year && 
            record.fullDate.getMonth() + 1 == month
        );
        console.log('Filtered attendance for', selectedMonth, ':', attendanceToShow.length, 'records');
    } else {
        console.log('Showing all attendance records:', attendanceToShow.length);
    }
    
    // Create attendance table
    const attendanceTableHtml = attendanceToShow.length > 0 ? `
        <div class="attendance-table-wrapper">
            <table class="attendance-table">
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Day</th>
                        <th>Morning Entrance</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${attendanceToShow.map((record, index) => {
                        const isLate = record.morning.in && record.morning.in > '06:45';
                        const isWeekend = record.dayOfWeek === 'SUN' || record.dayOfWeek === 'SAT';
                        const dateString = record.fullDate ? record.fullDate.toISOString().split('T')[0] : null;
                        const isHolidayDate = dateString && isHoliday(dateString);
                        
                        let status = '';
                        let statusClass = '';
                        
                        if (isHolidayDate) {
                            status = 'HOLIDAY';
                            statusClass = 'status-holiday';
                        } else if (isWeekend) {
                            status = 'WEEKEND';
                            statusClass = 'status-weekend';
                        } else if (record.morning.in) {
                            if (isLate) {
                                status = 'LATE';
                                statusClass = 'status-late';
                            } else {
                                status = 'ON TIME';
                                statusClass = 'status-ontime';
                            }
                        }
                        
                        return `
                        <tr class="attendance-row" style="animation-delay: ${300 + (index * 50)}ms">
                            <td>
                                <span class="date-display">${record.date}</span>
                            </td>
                            <td>
                                <span class="day-display ${isWeekend ? 'weekend' : 'weekday'}">${record.dayOfWeek}</span>
                            </td>
                            <td>
                                <div class="entrance-time">
                                    ${record.morning.in ? 
                                        `<span class="time-display ${isLate ? 'late' : 'ontime'}">${record.morning.in}</span>` : 
                                        '<span class="no-time">-</span>'
                                    }
                                    ${isHolidayDate ? '<i class="bi bi-star-fill holiday-icon" title="Holiday"></i>' : ''}
                                </div>
                            </td>
                            <td>
                                ${status ? `<span class="status-badge ${statusClass}">${status}</span>` : '-'}
                            </td>
                        </tr>
                        `;
                    }).join('')}
                </tbody>
            </table>
        </div>
        
        <div class="stats-summary" data-animate="fadeInUp" data-delay="400">
            <div class="summary-card present-card">
                <div class="summary-icon">
                    <i class="bi bi-sunrise"></i>
                </div>
                <div class="summary-content">
                    <h4 class="summary-number">${getPresentDaysCount(attendanceToShow)}</h4>
                    <p class="summary-label">Present Days</p>
                    <small class="summary-note">Excl. weekends & holidays</small>
                </div>
            </div>
            
            <div class="summary-card ontime-card">
                <div class="summary-icon">
                    <i class="bi bi-check-circle"></i>
                </div>
                <div class="summary-content">
                    <h4 class="summary-number">${attendanceToShow.filter(r => {
                        const isWeekend = r.dayOfWeek === 'SUN' || r.dayOfWeek === 'SAT';
                        const dateString = r.fullDate ? r.fullDate.toISOString().split('T')[0] : null;
                        const isHolidayDate = dateString && isHoliday(dateString);
                        return r.morning.in && r.morning.in <= '06:45' && !isWeekend && !isHolidayDate;
                    }).length}</h4>
                    <p class="summary-label">On Time</p>
                    <small class="summary-note">≤ 6:45 AM</small>
                </div>
            </div>
            
            <div class="summary-card late-card">
                <div class="summary-icon">
                    <i class="bi bi-clock"></i>
                </div>
                <div class="summary-content">
                    <h4 class="summary-number">${attendanceToShow.filter(r => {
                        const isWeekend = r.dayOfWeek === 'SUN' || r.dayOfWeek === 'SAT';
                        const dateString = r.fullDate ? r.fullDate.toISOString().split('T')[0] : null;
                        const isHolidayDate = dateString && isHoliday(dateString);
                        return r.morning.in && r.morning.in > '06:45' && !isWeekend && !isHolidayDate;
                    }).length}</h4>
                    <p class="summary-label">Late Days</p>
                    <small class="summary-note">> 6:45 AM</small>
                </div>
            </div>
        </div>
        
        <style>
            /* Mobile-Responsive Attendance Table Styles */
            .attendance-table-wrapper {
                max-height: 400px;
                overflow-y: auto;
                overflow-x: auto;
                margin-bottom: 2rem;
                -webkit-overflow-scrolling: touch;
            }
            
            .attendance-table {
                width: 100%;
                border-collapse: collapse;
                min-width: 600px;
            }
            
            .attendance-table th {
                background: #0f0f0f;
                padding: 1rem;
                font-weight: 500;
                font-size: 0.875rem;
                color: #10b981;
                text-align: left;
                border: none;
                border-bottom: 1px solid #2d2d2d;
                position: sticky;
                top: 0;
                z-index: 10;
            }
            
            .attendance-table td {
                padding: 1rem;
                border: none;
                border-bottom: 1px solid #2d2d2d;
                font-size: 0.875rem;
                vertical-align: middle;
                color: #e2e8f0;
            }
            
            .attendance-row {
                animation: slideInLeft 0.4s cubic-bezier(0.4, 0, 0.2, 1) forwards;
                opacity: 0;
                transition: background-color 0.2s ease;
            }
            
            .attendance-row:hover {
                background-color: #0f0f0f;
            }
            
            .date-display {
                font-weight: 600;
                color: #10b981;
            }
            
            .day-display {
                padding: 0.25rem 0.75rem;
                border-radius: 8px;
                font-weight: 500;
                font-size: 0.75rem;
            }
            
            .day-display.weekday {
                background: #065f46;
                color: #a7f3d0;
                border: 1px solid #047857;
            }
            
            .day-display.weekend {
                background: #374151;
                color: #f59e0b;
                border: 1px solid #4b5563;
            }
            
            .entrance-time {
                display: flex;
                align-items: center;
                gap: 0.5rem;
                flex-wrap: wrap;
            }
            
            .time-display {
                padding: 0.5rem 1rem;
                border-radius: 8px;
                font-weight: 600;
                font-size: 0.875rem;
            }
            
            .time-display.ontime {
                background: #065f46;
                color: #a7f3d0;
                border: 1px solid #047857;
            }
            
            .time-display.late {
                background: #dc2626;
                color: #fef2f2;
                border: 1px solid #b91c1c;
            }
            
            .no-time {
                color: #6b7280;
                font-style: italic;
            }
            
            .holiday-icon {
                color: #f59e0b;
                font-size: 0.875rem;
            }
            
            .status-badge {
                padding: 0.25rem 0.75rem;
                border-radius: 8px;
                font-weight: 600;
                font-size: 0.75rem;
            }
            
            .status-ontime {
                background: #065f46;
                color: #a7f3d0;
                border: 1px solid #047857;
            }
            
            .status-late {
                background: #dc2626;
                color: #fef2f2;
                border: 1px solid #b91c1c;
            }
            
            .status-weekend {
                background: #374151;
                color: #f59e0b;
                border: 1px solid #4b5563;
            }
            
            .status-holiday {
                background: #374151;
                color: #9ca3af;
                border: 1px solid #4b5563;
            }
            
            /* Mobile-Responsive Stats Summary */
            .stats-summary {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 1rem;
            }
            
            .summary-card {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 12px;
                padding: 1.5rem;
                box-shadow: 0 2px 15px rgba(0,0,0,0.3);
                transition: all 0.3s ease;
                display: flex;
                align-items: center;
                gap: 1rem;
            }
            
            .summary-card:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 25px rgba(16, 185, 129, 0.2);
                border-color: #065f46;
            }
            
            .present-card {
                border-left: 4px solid #10b981;
            }
            
            .ontime-card {
                border-left: 4px solid #059669;
            }
            
            .late-card {
                border-left: 4px solid #f59e0b;
            }
            
            .summary-icon {
                width: 48px;
                height: 48px;
                border-radius: 12px;
                display: flex;
                align-items: center;
                justify-content: center;
                flex-shrink: 0;
            }
            
            .present-card .summary-icon {
                background: rgba(16, 185, 129, 0.2);
                color: #10b981;
            }
            
            .ontime-card .summary-icon {
                background: rgba(5, 150, 105, 0.2);
                color: #059669;
            }
            
            .late-card .summary-icon {
                background: rgba(245, 158, 11, 0.2);
                color: #f59e0b;
            }
            
            .summary-icon i {
                font-size: 20px;
            }
            
            .summary-content {
                flex: 1;
            }
            
            .summary-number {
                font-size: 2rem;
                font-weight: 300;
                line-height: 1;
                margin-bottom: 0.25rem;
                color: #10b981;
            }
            
            .summary-label {
                font-size: 1rem;
                font-weight: 500;
                color: #a7f3d0;
                margin-bottom: 0.25rem;
            }
            
            .summary-note {
                font-size: 0.75rem;
                color: #9ca3af;
            }
            
            /* Mobile Responsive Adjustments */
            @media (max-width: 480px) {
                /* Attendance Table Mobile */
                .attendance-table-wrapper {
                    margin: 0 -1rem 2rem -1rem;
                    border-radius: 0;
                }
                
                .attendance-table {
                    font-size: 0.75rem;
                    min-width: 500px;
                }
                
                .attendance-table th,
                .attendance-table td {
                    padding: 0.5rem 0.25rem;
                }
                
                .attendance-table th:first-child,
                .attendance-table td:first-child {
                    padding-left: 1rem;
                }
                
                .attendance-table th:last-child,
                .attendance-table td:last-child {
                    padding-right: 1rem;
                }
                
                .time-display {
                    padding: 0.375rem 0.75rem;
                    font-size: 0.75rem;
                }
                
                .day-display {
                    padding: 0.25rem 0.5rem;
                    font-size: 0.625rem;
                }
                
                .status-badge {
                    padding: 0.25rem 0.5rem;
                    font-size: 0.625rem;
                }
                
                /* Stats Summary Mobile */
                .stats-summary {
                    grid-template-columns: 1fr;
                    gap: 0.75rem;
                }
                
                .summary-card {
                    padding: 1rem;
                    border-radius: 8px;
                }
                
                .summary-icon {
                    width: 40px;
                    height: 40px;
                }
                
                .summary-icon i {
                    font-size: 16px;
                }
                
                .summary-number {
                    font-size: 1.5rem;
                }
                
                .summary-label {
                    font-size: 0.875rem;
                }
                
                .summary-note {
                    font-size: 0.625rem;
                }
            }
            
            @media (min-width: 481px) and (max-width: 768px) {
                .stats-summary {
                    grid-template-columns: repeat(2, 1fr);
                }
                
                .attendance-table {
                    font-size: 0.8125rem;
                }
            }
            
            /* Scrollbar styling */
            .attendance-table-wrapper::-webkit-scrollbar {
                width: 6px;
            }
            
            .attendance-table-wrapper::-webkit-scrollbar-track {
                background: #2d2d2d;
            }
            
            .attendance-table-wrapper::-webkit-scrollbar-thumb {
                background: #065f46;
                border-radius: 3px;
            }
            
            .attendance-table-wrapper::-webkit-scrollbar-thumb:hover {
                background: #10b981;
            }
            
            /* Animation for attendance rows */
            @keyframes slideInLeft {
                from {
                    opacity: 0;
                    transform: translateX(-20px);
                }
                to {
                    opacity: 1;
                    transform: translateX(0);
                }
            }
        </style>
    ` : '<p class="text-muted">No attendance records found for this period.</p>';
    
    const modalHtml = `
        <div class="modal fade" id="employeeModal" tabindex="-1">
            <div class="modal-dialog modal-xl">
                <div class="modal-content border-0 shadow-lg employee-modal">
                    <div class="modal-header border-0 text-white" style="background: linear-gradient(135deg, #065f46 0%, #022c22 100%); border-bottom: 1px solid #2d2d2d;">
                        <h5 class="modal-title fw-light" style="color: #10b981;">
                            <i class="bi bi-person-circle me-2"></i>${employee.name} - Attendance Details
                            ${selectedMonth !== 'all' ? ` (${new Date(selectedMonth + '-01').toLocaleString('default', { month: 'long', year: 'numeric' })})` : ''}
                        </h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body p-4" style="background: #1a1a1a; color: #e2e8f0;">
                        <div class="employee-info" data-animate="fadeInUp" data-delay="0">
                            <div class="info-card">
                                <i class="bi bi-card-text info-icon"></i>
                                <span class="info-label">Prefect ID:</span>
                                <span class="info-value">${employee.employeeId}</span>
                            </div>
                        </div>
                        
                        <h6 class="section-title" data-animate="fadeInUp" data-delay="100">
                            <i class="bi bi-sunrise me-2"></i>Morning Entrance Records
                        </h6>
                        
                        <div class="attendance-container" data-animate="fadeInUp" data-delay="200">
                            ${attendanceTableHtml}
                        </div>
                    </div>
                    <div class="modal-footer border-0" style="background: #1a1a1a; border-top: 1px solid #2d2d2d;">
                        <button type="button" class="btn btn-minimal btn-download" onclick="downloadEmployeeAttendance('${employeeName}', '${selectedMonth}')">
                            <i class="bi bi-download me-2"></i>Download Attendance
                        </button>
                        <button type="button" class="btn btn-minimal btn-secondary" data-bs-dismiss="modal">Close</button>
                    </div>
                </div>
            </div>
        </div>
        
        <style>
            /* Mobile-Responsive Employee Modal Styles */
            .employee-modal .modal-content {
                border-radius: 20px;
                overflow: hidden;
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
            }
            
            .employee-info {
                margin-bottom: 2rem;
            }
            
            .info-card {
                background: #0f0f0f;
                border: 1px solid #2d2d2d;
                border-radius: 12px;
                padding: 1rem;
                display: flex;
                align-items: center;
                gap: 1rem;
                border-left: 4px solid #667eea;
                flex-wrap: wrap;
            }
            
            .info-icon {
                font-size: 1.25rem;
                color: #667eea;
                flex-shrink: 0;
            }
            
            .info-label {
                font-weight: 600;
                color: #4a5568;
            }
            
            .info-value {
                color: #2d3748;
                font-weight: 500;
            }
            
            .section-title {
                font-size: 1.125rem;
                font-weight: 500;
                color: #2d3748;
                margin-bottom: 1rem;
                display: flex;
                align-items: center;
                flex-wrap: wrap;
            }
            
            .attendance-container {
                background: #1a1a1a;
                border-radius: 12px;
                overflow: hidden;
                box-shadow: 0 2px 20px rgba(0,0,0,0.3);
                border: 1px solid #2d2d2d;
            }
            
            /* Mobile Modal Responsive */
            @media (max-width: 480px) {
                .modal-dialog {
                    margin: 0.5rem;
                    max-width: none;
                }
                
                .employee-modal .modal-content {
                    border-radius: 16px;
                }
                
                .modal-header {
                    padding: 1rem 1.5rem;
                }
                
                .modal-title {
                    font-size: 1rem;
                }
                
                .modal-body {
                    padding: 1rem 1.5rem !important;
                }
                
                .modal-footer {
                    padding: 1rem 1.5rem;
                    flex-direction: column;
                    gap: 0.5rem;
                }
                
                .modal-footer .btn {
                    width: 100%;
                    justify-content: center;
                }
                
                .info-card {
                    padding: 0.75rem;
                    flex-direction: column;
                    align-items: flex-start;
                    gap: 0.5rem;
                }
                
                .info-icon {
                    font-size: 1rem;
                }
                
                .section-title {
                    font-size: 1rem;
                    text-align: center;
                    justify-content: center;
                }
                
                .employee-info {
                    margin-bottom: 1.5rem;
                }
            }
            
            @media (min-width: 481px) and (max-width: 768px) {
                .modal-dialog {
                    margin: 1rem;
                }
                
                .modal-footer {
                    flex-direction: row;
                    justify-content: center;
                    gap: 1rem;
                }
                
                .modal-footer .btn {
                    flex: 1;
                    max-width: 200px;
                }
            }
        </style>
    `;
    
    // Remove existing modal if any
    const existingModal = document.getElementById('employeeModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Add modal to body
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    
    // Show modal
    const modal = new bootstrap.Modal(document.getElementById('employeeModal'));
    
    // Handle modal events for accessibility
    const modalElement = document.getElementById('employeeModal');
    
    // Remove focus from any focused elements before showing modal
    if (document.activeElement) {
        document.activeElement.blur();
    }
    
    // Focus management
    modalElement.addEventListener('shown.bs.modal', function () {
        // Focus the first focusable element in the modal
        const firstFocusable = modalElement.querySelector('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
        if (firstFocusable) {
            firstFocusable.focus();
        }
    });
    
    modalElement.addEventListener('hide.bs.modal', function () {
        // Remove focus from any element inside the modal before it's hidden
        const focusedElement = modalElement.querySelector(':focus');
        if (focusedElement) {
            focusedElement.blur();
        }
        
        // Also remove focus from the active element if it's inside the modal
        if (document.activeElement && modalElement.contains(document.activeElement)) {
            document.activeElement.blur();
        }
    });
    
    modalElement.addEventListener('hidden.bs.modal', function () {
        // Clean up the modal element after it's hidden
        setTimeout(() => {
            if (modalElement && modalElement.parentNode) {
                modalElement.remove();
            }
        }, 150);
    });
    
    modal.show();
}

// Download detailed CSV report
function downloadDetailedCSV(selectedMonth) {
    // Generate comprehensive Excel analysis
    downloadComprehensiveExcelAnalysis(selectedMonth);
}

// Download comprehensive Excel analysis for all prefects
function downloadComprehensiveExcelAnalysis(selectedMonth) {
    if (extractedData.length === 0) {
        alert('No data to download');
        return;
    }

    // Create workbook
    const wb = XLSX.utils.book_new();
    
    // Get all available months from the data
    const availableMonths = new Set();
    extractedData.forEach(employee => {
        employee.attendanceData.forEach(record => {
            if (record.year && record.month) {
                const monthKey = `${record.year}-${String(record.month).padStart(2, '0')}`;
                availableMonths.add(monthKey);
            }
        });
    });
    
    const sortedMonths = Array.from(availableMonths).sort();
    
    // If specific month selected, only process that month
    const monthsToProcess = selectedMonth !== 'all' ? [selectedMonth] : sortedMonths;
    
    if (monthsToProcess.length === 0) {
        alert('No data to download');
        return;
    }
    
    // Create a worksheet for each month
    monthsToProcess.forEach(monthKey => {
        const [year, month] = monthKey.split('-');
        const analysisYear = parseInt(year);
        const analysisMonth = parseInt(month);
        const monthName = new Date(analysisYear, analysisMonth - 1).toLocaleString('default', { month: 'long', year: 'numeric' });
        const daysInMonth = new Date(analysisYear, analysisMonth, 0).getDate();
        
        // Filter data for this specific month
        const monthData = extractedData.filter(employee => {
            return employee.attendanceData.some(record => 
                record.year == year && record.month == month
            );
        }).map(employee => ({
            ...employee,
            attendanceData: employee.attendanceData.filter(record => 
                record.year == year && record.month == month
            )
        }));
        
        if (monthData.length === 0) return;
        
        // Create worksheet data
        const sheetData = [];
        
        // Header with month and year
        sheetData.push([monthName, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Summary']);
        sheetData.push(['S.No', 'Name', ...Array.from({length: daysInMonth}, (_, i) => i + 1), 'Total Present', 'Total Late', 'On Time', 'Working Days', 'Attendance %']);
        
        // Add data for each employee
        monthData.forEach((employee, index) => {
            const row = [index + 1, employee.name];
            
            // Create attendance map for quick lookup
            const attendanceMap = {};
            employee.attendanceData.forEach(record => {
                if (record.fullDate) {
                    const day = record.fullDate.getDate();
                    attendanceMap[day] = record;
                }
            });
            
            // Fill daily attendance
            let presentCount = 0;
            let lateCount = 0;
            let onTimeCount = 0;
            let workingDaysCount = 0;
            
            for (let day = 1; day <= daysInMonth; day++) {
                const record = attendanceMap[day];
                const checkDate = new Date(analysisYear, analysisMonth - 1, day);
                const dayOfWeek = checkDate.toLocaleDateString('en-US', { weekday: 'short' }).toUpperCase();
                const dateString = checkDate.toISOString().split('T')[0];
                const isHolidayDate = isHoliday(dateString);
                const isWeekend = dayOfWeek === 'SAT' || dayOfWeek === 'SUN';
                
                if (isHolidayDate) {
                    row.push('H'); // Holiday
                } else if (isWeekend) {
                    row.push('W'); // Weekend
                } else {
                    workingDaysCount++;
                    if (record && record.morning.in) {
                        const [hours, minutes] = record.morning.in.split(':').map(Number);
                        const isLate = hours > 6 || (hours === 6 && minutes > 45);
                        
                        if (isLate) {
                            row.push(`L\n${record.morning.in}`); // Late with time
                            lateCount++;
                        } else {
                            row.push(`P\n${record.morning.in}`); // Present on time with time
                            onTimeCount++;
                        }
                        presentCount++;
                    } else {
                        row.push('A'); // Absent
                    }
                }
            }
            
            // Add summary columns
            const attendanceRate = workingDaysCount > 0 ? Math.round((presentCount / workingDaysCount) * 100) : 0;
            row.push(presentCount, lateCount, onTimeCount, workingDaysCount, `${attendanceRate}%`);
            
            sheetData.push(row);
        });
        
        // Add legend
        sheetData.push([]);
        sheetData.push(['Legend:']);
        sheetData.push(['P = Present (On Time)', 'L = Late', 'A = Absent', 'H = Holiday', 'W = Weekend']);
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        
        // Set column widths - make day columns wider to accommodate time
        const colWidths = [
            { wch: 6 },  // S.No
            { wch: 20 }, // Name
            ...Array.from({length: daysInMonth}, () => ({ wch: 8 })), // Day columns (wider for time)
            { wch: 12 }, // Total Present
            { wch: 10 }, // Total Late
            { wch: 10 }, // On Time
            { wch: 12 }, // Working Days
            { wch: 12 }  // Attendance %
        ];
        ws['!cols'] = colWidths;
        
        // Add cell styling and colors
        const range = XLSX.utils.decode_range(ws['!ref']);
        
        // Style header rows
        for (let C = range.s.c; C <= range.e.c; ++C) {
            // First row (month header)
            const headerCell = XLSX.utils.encode_cell({r: 0, c: C});
            if (!ws[headerCell]) ws[headerCell] = {t: 's', v: ''};
            ws[headerCell].s = {
                fill: { fgColor: { rgb: "366092" } },
                font: { color: { rgb: "FFFFFF" }, bold: true, sz: 14 },
                alignment: { horizontal: "center", vertical: "center" }
            };
            
            // Second row (column headers)
            const colHeaderCell = XLSX.utils.encode_cell({r: 1, c: C});
            if (!ws[colHeaderCell]) ws[colHeaderCell] = {t: 's', v: ''};
            ws[colHeaderCell].s = {
                fill: { fgColor: { rgb: "D9E2F3" } },
                font: { bold: true, sz: 10 },
                alignment: { horizontal: "center", vertical: "center" },
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                }
            };
        }
        
        // Style data rows with colors based on attendance
        for (let R = 2; R < sheetData.length - 3; ++R) { // Exclude legend rows
            for (let C = 0; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({r: R, c: C});
                if (!ws[cellAddress]) continue;
                
                const cellValue = ws[cellAddress].v;
                let cellStyle = {
                    alignment: { horizontal: "center", vertical: "center", wrapText: true },
                    border: {
                        top: { style: "thin", color: { rgb: "CCCCCC" } },
                        bottom: { style: "thin", color: { rgb: "CCCCCC" } },
                        left: { style: "thin", color: { rgb: "CCCCCC" } },
                        right: { style: "thin", color: { rgb: "CCCCCC" } }
                    },
                    font: { sz: 9 }
                };
                
                // Color coding based on cell content
                if (C >= 2 && C <= daysInMonth + 1) { // Day columns
                    if (typeof cellValue === 'string') {
                        if (cellValue.startsWith('P')) {
                            // Present (Green)
                            cellStyle.fill = { fgColor: { rgb: "C6EFCE" } };
                            cellStyle.font.color = { rgb: "006100" };
                        } else if (cellValue.startsWith('L')) {
                            // Late (Orange/Yellow)
                            cellStyle.fill = { fgColor: { rgb: "FFEB9C" } };
                            cellStyle.font.color = { rgb: "9C5700" };
                        } else if (cellValue === 'A') {
                            // Absent (Red)
                            cellStyle.fill = { fgColor: { rgb: "FFC7CE" } };
                            cellStyle.font.color = { rgb: "9C0006" };
                        } else if (cellValue === 'H') {
                            // Holiday (Blue)
                            cellStyle.fill = { fgColor: { rgb: "BDD7EE" } };
                            cellStyle.font.color = { rgb: "0070C0" };
                        } else if (cellValue === 'W') {
                            // Weekend (Gray)
                            cellStyle.fill = { fgColor: { rgb: "E2E2E2" } };
                            cellStyle.font.color = { rgb: "7C7C7C" };
                        }
                    }
                } else if (C === 0 || C === 1) {
                    // S.No and Name columns
                    cellStyle.fill = { fgColor: { rgb: "F2F2F2" } };
                    cellStyle.font.bold = true;
                    if (C === 1) cellStyle.alignment.horizontal = "left";
                } else if (C > daysInMonth + 1) {
                    // Summary columns
                    cellStyle.fill = { fgColor: { rgb: "E7E6E6" } };
                    cellStyle.font.bold = true;
                }
                
                ws[cellAddress].s = cellStyle;
            }
        }
        
        // Style legend
        const legendStartRow = sheetData.length - 2;
        for (let R = legendStartRow; R < sheetData.length; ++R) {
            for (let C = 0; C <= 4; ++C) {
                const cellAddress = XLSX.utils.encode_cell({r: R, c: C});
                if (ws[cellAddress]) {
                    ws[cellAddress].s = {
                        font: { italic: true, sz: 9 },
                        alignment: { horizontal: "left", vertical: "center" }
                    };
                }
            }
        }
        
        // Set row heights for better readability
        if (!ws['!rows']) ws['!rows'] = [];
        for (let i = 2; i < sheetData.length - 3; i++) {
            ws['!rows'][i] = { hpx: 30 }; // Increase row height for wrapped text
        }
        
        // Add worksheet to workbook
        const sheetName = new Date(analysisYear, analysisMonth - 1).toLocaleString('default', { month: 'short', year: 'numeric' }).replace(' ', '_');
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    
    // Generate filename
    const filename = selectedMonth !== 'all' 
        ? `Attendance_Report_${selectedMonth}_${new Date().toISOString().split('T')[0]}.xlsx`
        : `Attendance_Report_All_Months_${new Date().toISOString().split('T')[0]}.xlsx`;
    
    // Save file
    XLSX.writeFile(wb, filename);
}

// Download summary CSV
function downloadSummaryCSV(selectedMonth) {
    let filteredData = extractedData;
    if (selectedMonth !== 'all') {
        const [year, month] = selectedMonth.split('-');
        filteredData = extractedData.filter(record => 
            record.year == year && record.month == month
        );
    }
    
    if (filteredData.length === 0) {
        alert('No data to download');
        return;
    }
    
    // Create summary CSV content focusing on morning entrance
    const csvContent = [
        ['#', 'Prefect Name', 'Morning Entries Count', 'First Morning Entry', 'Last Morning Entry', 'Sheet', 'Row'],
        ...filteredData.map((employee, index) => {
            const morningEntries = employee.attendanceData.filter(r => r.morning.in && r.dayOfWeek !== 'SUN' && r.dayOfWeek !== 'SAT');
            return [
                index + 1,
                employee.name,
                morningEntries.length,
                morningEntries.length > 0 ? morningEntries[0].morning.in : '',
                morningEntries.length > 0 ? morningEntries[morningEntries.length - 1].morning.in : '',
                employee.sheet,
                employee.row
            ];
        })
    ].map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
    
    downloadCSVFile(csvContent, `fingerprint_morning_summary_${selectedMonth}.csv`);
}

// Download individual employee attendance
function downloadEmployeeAttendance(employeeName, selectedMonth) {
    let filteredData = extractedData;
    if (selectedMonth !== 'all') {
        const [year, month] = selectedMonth.split('-');
        filteredData = extractedData.filter(record => 
            record.year == year && record.month == month
        );
    }
    
    const employee = filteredData.find(record => record.name === employeeName);
    if (!employee) return;
    
    // Filter attendance data by month if needed
    let attendanceToShow = employee.attendanceData;
    if (selectedMonth !== 'all') {
        const [year, month] = selectedMonth.split('-');
        attendanceToShow = employee.attendanceData.filter(record => 
            record.fullDate && record.fullDate.getFullYear() == year && 
            record.fullDate.getMonth() + 1 == month
        );
    }
    
    // Create CSV content for individual employee focusing on morning entrance
    const csvRows = [
        ['Prefect:', employee.name],
        ['Prefect ID:', employee.employeeId],
        [''],
        ['Date', 'Day', 'Morning Entrance', 'Late Status']
    ];
    
    attendanceToShow.forEach(record => {
        const isWeekend = record.dayOfWeek === 'SUN' || record.dayOfWeek === 'SAT';
        const isLate = record.morning.in && record.morning.in > '06:45';
        let status = '';
        
        if (record.morning.in) {
            if (isWeekend) {
                status = 'WEEKEND';
            } else {
                status = isLate ? 'LATE' : 'ON TIME';
            }
        }
        
        csvRows.push([
            record.date,
            record.dayOfWeek,
            record.morning.in || '',
            status
        ]);
    });
    
    const csvContent = csvRows.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
    const filename = `${employee.name.replace(/\s+/g, '_')}_morning_attendance_${selectedMonth === 'all' ? 'all' : selectedMonth}.csv`;
    
    downloadCSVFile(csvContent, filename);
}

// Helper function to check if a date is a holiday
function isHoliday(dateString) {
    return holidays.has(dateString);
}

// Helper function to get working days count (excluding weekends and holidays)
function getWorkingDaysCount(attendanceData) {
    return attendanceData.filter(record => {
        const isWeekend = record.dayOfWeek === 'SUN' || record.dayOfWeek === 'SAT';
        const dateString = record.fullDate ? record.fullDate.toISOString().split('T')[0] : null;
        const isHolidayDate = dateString && isHoliday(dateString);
        return !isWeekend && !isHolidayDate;
    }).length;
}

// Helper function to get present days count (excluding weekends and holidays)
function getPresentDaysCount(attendanceData) {
    return attendanceData.filter(record => {
        const isWeekend = record.dayOfWeek === 'SUN' || record.dayOfWeek === 'SAT';
        const dateString = record.fullDate ? record.fullDate.toISOString().split('T')[0] : null;
        const isHolidayDate = dateString && isHoliday(dateString);
        return record.morning.in && !isWeekend && !isHolidayDate;
    }).length;
}

// Show holiday calendar modal
function showHolidayCalendar(selectedMonth) {
    let currentMonth, currentYear;
    
    if (selectedMonth === 'all' && allMonths.length > 0) {
        // Default to first available month
        [currentYear, currentMonth] = allMonths[0].split('-');
        currentMonth = parseInt(currentMonth);
        currentYear = parseInt(currentYear);
    } else if (selectedMonth !== 'all') {
        [currentYear, currentMonth] = selectedMonth.split('-');
        currentMonth = parseInt(currentMonth);
        currentYear = parseInt(currentYear);
    } else {
        // Default to current month
        const now = new Date();
        currentMonth = now.getMonth() + 1;
        currentYear = now.getFullYear();
    }
    
    const modalHtml = `
        <div class="modal fade" id="holidayModal" tabindex="-1">
            <div class="modal-dialog modal-lg">
                <div class="modal-content border-0 shadow-lg holiday-modal">
                    <div class="modal-header border-0 bg-gradient-primary text-white">
                        <h5 class="modal-title fw-light">
                            <i class="bi bi-calendar-event me-2"></i>Holiday Management
                        </h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body p-4" style="background: #1a1a1a; color: #e2e8f0;">
                        <div class="holiday-controls" data-animate="fadeInUp" data-delay="0">
                            <div class="control-group">
                                <label for="holidayMonthSelect" class="control-label">Select Month:</label>
                                <select id="holidayMonthSelect" class="control-select" onchange="updateHolidayCalendar()">
                                    ${allMonths.map(month => {
                                        const [year, monthNum] = month.split('-');
                                        const monthName = new Date(year, monthNum - 1).toLocaleString('default', { month: 'long', year: 'numeric' });
                                        const isSelected = year == currentYear && monthNum == currentMonth;
                                        return `<option value="${month}" ${isSelected ? 'selected' : ''}>${monthName}</option>`;
                                    }).join('')}
                                </select>
                            </div>
                            <div class="control-buttons">
                                <button class="btn-control btn-warning" onclick="clearAllHolidays()">
                                    <i class="bi bi-trash"></i> Clear All
                                </button>
                                <button class="btn-control btn-success" onclick="saveHolidays()">
                                    <i class="bi bi-check"></i> Save
                                </button>
                            </div>
                        </div>
                        
                        <div class="holiday-info" data-animate="fadeInUp" data-delay="100">
                            <i class="bi bi-info-circle"></i>
                            <span>Click on dates to toggle holiday status. Holidays will be excluded from attendance calculations.</span>
                        </div>
                        
                        <div class="calendar-container" data-animate="fadeInUp" data-delay="200">
                            <div id="holidayCalendarContainer">
                                ${generateCalendarHTML(currentYear, currentMonth)}
                            </div>
                        </div>
                        
                        <div class="holidays-list" data-animate="fadeInUp" data-delay="300">
                            <h6 class="list-title">Current Holidays:</h6>
                            <div id="holidayList" class="holiday-badges">
                                ${generateHolidayList(currentYear, currentMonth)}
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer border-0" style="background: #1a1a1a; border-top: 1px solid #2d2d2d;">
                        <button type="button" class="btn btn-minimal btn-secondary" data-bs-dismiss="modal">Close</button>
                        <button type="button" class="btn btn-minimal btn-download" onclick="applyHolidaysAndRefresh()" data-bs-dismiss="modal">
                            <i class="bi bi-check-circle me-2"></i>Apply Changes
                        </button>
                    </div>
                </div>
            </div>
        </div>
        
        <style>
            /* Mobile-Responsive Holiday Modal Styles */
            .holiday-modal .modal-content {
                border-radius: 20px;
                overflow: hidden;
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
            }
            
            .holiday-controls {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 1.5rem;
                flex-wrap: wrap;
                gap: 1rem;
            }
            
            .control-group {
                display: flex;
                align-items: center;
                gap: 1rem;
                flex-wrap: wrap;
            }
            
            .control-label {
                font-weight: 500;
                color: #10b981;
                margin: 0;
            }
            
            .control-select {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 8px;
                padding: 0.5rem 1rem;
                color: #e2e8f0;
                font-size: 0.875rem;
                min-width: 200px;
            }
            
            .control-buttons {
                display: flex;
                gap: 0.5rem;
                flex-wrap: wrap;
            }
            
            .btn-control {
                border: none;
                border-radius: 8px;
                padding: 0.5rem 1rem;
                font-size: 0.875rem;
                font-weight: 500;
                display: flex;
                align-items: center;
                gap: 0.5rem;
                transition: all 0.2s ease;
                cursor: pointer;
            }
            
            .btn-warning {
                background: linear-gradient(135deg, #ed8936, #dd6b20);
                color: white;
            }
            
            .btn-success {
                background: linear-gradient(135deg, #48bb78, #38a169);
                color: white;
            }
            
            .btn-control:hover {
                transform: translateY(-1px);
                box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            }
            
            .holiday-info {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 8px;
                padding: 1rem;
                margin-bottom: 1.5rem;
                display: flex;
                align-items: center;
                gap: 0.75rem;
                color: #10b981;
            }
            
            .holiday-info i {
                font-size: 1.125rem;
                flex-shrink: 0;
            }
            
            .calendar-container {
                background: #1a1a1a;
                border-radius: 12px;
                padding: 1.5rem;
                box-shadow: 0 2px 20px rgba(0,0,0,0.3);
                margin-bottom: 1.5rem;
                overflow-x: auto;
                border: 1px solid #2d2d2d;
            }
            
            .holidays-list {
                background: #1a1a1a;
                border-radius: 12px;
                padding: 1.5rem;
                border: 1px solid #2d2d2d;
            }
            
            .list-title {
                font-size: 1rem;
                font-weight: 500;
                color: #10b981;
                margin-bottom: 1rem;
            }
            
            .holiday-badges {
                display: flex;
                flex-wrap: wrap;
                gap: 0.5rem;
            }
            
            .holiday-badges .badge {
                background: #1a1a1a;
                color: #ef4444;
                border: 1px solid #ef4444;
                border-radius: 8px;
                padding: 0.5rem 1rem;
                font-size: 0.875rem;
                font-weight: 500;
                display: flex;
                align-items: center;
                gap: 0.5rem;
            }
            
            .holiday-badges .btn-close {
                font-size: 0.75rem;
                margin-left: 0.25rem;
            }
            
            /* Mobile Holiday Modal Responsive */
            @media (max-width: 480px) {
                .holiday-modal .modal-dialog {
                    margin: 0.5rem;
                    max-width: none;
                }
                
                .holiday-modal .modal-content {
                    border-radius: 16px;
                }
                
                .modal-header {
                    padding: 1rem 1.5rem;
                }
                
                .modal-title {
                    font-size: 1rem;
                }
                
                .modal-body {
                    padding: 1rem 1.5rem !important;
                }
                
                .modal-footer {
                    padding: 1rem 1.5rem;
                    flex-direction: column;
                    gap: 0.5rem;
                }
                
                .modal-footer .btn {
                    width: 100%;
                    justify-content: center;
                }
                
                .holiday-controls {
                    flex-direction: column;
                    align-items: stretch;
                    gap: 1rem;
                }
                
                .control-group {
                    flex-direction: column;
                    align-items: stretch;
                    gap: 0.5rem;
                }
                
                .control-label {
                    text-align: center;
                }
                
                .control-select {
                    min-width: 100%;
                    text-align: center;
                }
                
                .control-buttons {
                    justify-content: center;
                }
                
                .btn-control {
                    flex: 1;
                    justify-content: center;
                    font-size: 0.75rem;
                    padding: 0.75rem 1rem;
                }
                
                .holiday-info {
                    padding: 0.75rem;
                    flex-direction: column;
                    text-align: center;
                    gap: 0.5rem;
                }
                
                .calendar-container {
                    padding: 1rem;
                    margin: 0 -1rem 1.5rem -1rem;
                    border-radius: 0;
                }
                
                .holidays-list {
                    padding: 1rem;
                }
                
                .holiday-badges {
                    justify-content: center;
                }
                
                .holiday-badges .badge {
                    font-size: 0.75rem;
                    padding: 0.375rem 0.75rem;
                }
                
                /* Calendar Mobile Adjustments */
                .calendar-table {
                    font-size: 0.75rem;
                }
                
                .calendar-day {
                    height: 50px;
                    padding: 0.25rem;
                }
                
                .day-number {
                    font-size: 0.875rem;
                }
                
                .holiday-icon {
                    top: 2px;
                    right: 2px;
                    font-size: 0.625rem;
                }
                
                .calendar-table th {
                    padding: 0.5rem 0.25rem;
                    font-size: 0.75rem;
                }
            }
            
            @media (min-width: 481px) and (max-width: 768px) {
                .holiday-modal .modal-dialog {
                    margin: 1rem;
                }
                
                .holiday-controls {
                    flex-direction: column;
                    align-items: center;
                }
                
                .control-buttons {
                    justify-content: center;
                }
                
                .modal-footer {
                    justify-content: center;
                    gap: 1rem;
                }
                
                .modal-footer .btn {
                    flex: 1;
                    max-width: 200px;
                }
            }
        </style>
    `;
    
    // Remove existing modal if any
    const existingModal = document.getElementById('holidayModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Add modal to body
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    
    // Show modal
    const modal = new bootstrap.Modal(document.getElementById('holidayModal'));
    
    // Handle modal events for accessibility
    const modalElement = document.getElementById('holidayModal');
    
    // Remove focus from any focused elements before showing modal
    if (document.activeElement) {
        document.activeElement.blur();
    }
    
    // Focus management
    modalElement.addEventListener('shown.bs.modal', function () {
        // Focus the first focusable element in the modal
        const firstFocusable = modalElement.querySelector('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
        if (firstFocusable) {
            firstFocusable.focus();
        }
    });
    
    modalElement.addEventListener('hidden.bs.modal', function () {
        // Clean up the modal element after it's hidden
        setTimeout(() => {
            if (modalElement && modalElement.parentNode) {
                modalElement.remove();
            }
        }, 150);
    });
    
    modal.show();
}

// Generate calendar HTML
function generateCalendarHTML(year, month) {
    const firstDay = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0);
    const daysInMonth = lastDay.getDate();
    const startDayOfWeek = firstDay.getDay(); // 0 = Sunday
    
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    
    let calendarHTML = `
        <div class="calendar-header text-center mb-3">
            <h5>${monthNames[month - 1]} ${year}</h5>
        </div>
        <table class="table table-bordered calendar-table">
            <thead class="table-dark">
                <tr>
                    <th>Sun</th><th>Mon</th><th>Tue</th><th>Wed</th><th>Thu</th><th>Fri</th><th>Sat</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    let currentRow = '<tr>';
    
    // Add empty cells for days before the first day of the month
    for (let i = 0; i < startDayOfWeek; i++) {
        currentRow += '<td class="empty-day"></td>';
    }
    
    // Add days of the month
    for (let day = 1; day <= daysInMonth; day++) {
        const dateString = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const isHolidayDate = isHoliday(dateString);
        const dayOfWeek = new Date(year, month - 1, day).getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6; // Sunday or Saturday
        
        let cellClass = 'calendar-day';
        if (isWeekend) cellClass += ' weekend-day';
        if (isHolidayDate) cellClass += ' holiday-day';
        
        currentRow += `<td class="${cellClass}" onclick="toggleHoliday('${dateString}')" data-date="${dateString}">
                        <span class="day-number">${day}</span>
                        ${isHolidayDate ? '<i class="bi bi-star-fill holiday-icon"></i>' : ''}
                       </td>`;
        
        // If we've reached Saturday, start a new row
        if ((startDayOfWeek + day) % 7 === 0) {
            currentRow += '</tr>';
            calendarHTML += currentRow;
            currentRow = '<tr>';
        }
    }
    
    // Fill remaining cells in the last row
    const remainingCells = 7 - ((startDayOfWeek + daysInMonth) % 7);
    if (remainingCells < 7) {
        for (let i = 0; i < remainingCells; i++) {
            currentRow += '<td class="empty-day"></td>';
        }
        currentRow += '</tr>';
        calendarHTML += currentRow;
    }
    
    calendarHTML += `
            </tbody>
        </table>
        <style>
            .calendar-table { 
                cursor: pointer; 
                width: 100%;
                border-collapse: collapse;
                border-radius: 8px;
                overflow: hidden;
            }
            .calendar-day { 
                height: 60px; 
                vertical-align: top; 
                position: relative;
                transition: all 0.2s ease;
                border: 1px solid #2d2d2d;
                text-align: center;
                padding: 0.5rem;
                background: #0f0f0f;
                color: #e2e8f0;
            }
            .calendar-day:hover { 
                background-color: #1a1a1a; 
                transform: scale(1.02);
            }
            .weekend-day { 
                background-color: #2d1810; 
                color: #f59e0b;
                border: 1px solid #b45309;
            }
            .holiday-day { 
                background-color: #2d1a1a; 
                border: 2px solid #ef4444;
                color: #ef4444;
            }
            .holiday-day:hover { 
                background-color: #3d1a1a; 
            }
            .day-number { 
                font-weight: 600; 
                font-size: 1rem;
                display: block;
                margin-bottom: 0.25rem;
            }
            .holiday-icon { 
                position: absolute; 
                top: 4px; 
                right: 4px; 
                color: #f56565; 
                font-size: 0.75rem;
                animation: pulse 2s infinite;
            }
            .empty-day { 
                background-color: #0a0a0a; 
                opacity: 0.5;
            }
            
            .calendar-header {
                margin-bottom: 1rem;
            }
            
            .calendar-header h5 {
                font-weight: 500;
                color: #10b981;
            }
            
            .calendar-table th {
                background: #0f0f0f;
                color: #10b981;
                padding: 1rem 0.5rem;
                font-weight: 500;
                font-size: 0.875rem;
                text-align: center;
                border: 1px solid #2d2d2d;
            }
            
            @keyframes pulse {
                0%, 100% { opacity: 1; }
                50% { opacity: 0.5; }
            }
        </style>
    `;
    
    return calendarHTML;
}

// Generate holiday list
function generateHolidayList(year, month) {
    const monthHolidays = Array.from(holidays).filter(holiday => {
        return holiday.startsWith(`${year}-${String(month).padStart(2, '0')}`);
    }).sort();
    
    if (monthHolidays.length === 0) {
        return '<span class="text-muted">No holidays set for this month</span>';
    }
    
    return monthHolidays.map(holiday => {
        const date = new Date(holiday + 'T00:00:00');
        const day = date.getDate();
        const dayName = date.toLocaleDateString('en-US', { weekday: 'short' });
        return `<span class="badge bg-danger">${dayName}, ${day} <button type="button" class="btn-close btn-close-white btn-sm ms-1" onclick="removeHoliday('${holiday}')"></button></span>`;
    }).join('');
}

// Toggle holiday status
function toggleHoliday(dateString) {
    if (holidays.has(dateString)) {
        holidays.delete(dateString);
    } else {
        holidays.add(dateString);
    }
    
    // Update the calendar display
    const cell = document.querySelector(`[data-date="${dateString}"]`);
    if (cell) {
        if (holidays.has(dateString)) {
            cell.classList.add('holiday-day');
            const dayNumber = cell.querySelector('.day-number');
            if (dayNumber && !cell.querySelector('.holiday-icon')) {
                dayNumber.insertAdjacentHTML('afterend', '<i class="bi bi-star-fill holiday-icon"></i>');
            }
        } else {
            cell.classList.remove('holiday-day');
            const icon = cell.querySelector('.holiday-icon');
            if (icon) icon.remove();
        }
    }
    
    // Update holiday list
    const [year, month] = dateString.split('-');
    document.getElementById('holidayList').innerHTML = generateHolidayList(parseInt(year), parseInt(month));
}

// Remove specific holiday
function removeHoliday(dateString) {
    holidays.delete(dateString);
    
    // Update calendar display
    const cell = document.querySelector(`[data-date="${dateString}"]`);
    if (cell) {
        cell.classList.remove('holiday-day');
        const icon = cell.querySelector('.holiday-icon');
        if (icon) icon.remove();
    }
    
    // Update holiday list
    const [year, month] = dateString.split('-');
    document.getElementById('holidayList').innerHTML = generateHolidayList(parseInt(year), parseInt(month));
}

// Update calendar when month changes
function updateHolidayCalendar() {
    const select = document.getElementById('holidayMonthSelect');
    const [year, month] = select.value.split('-');
    
    document.getElementById('holidayCalendarContainer').innerHTML = generateCalendarHTML(parseInt(year), parseInt(month));
    document.getElementById('holidayList').innerHTML = generateHolidayList(parseInt(year), parseInt(month));
}

// Clear all holidays
function clearAllHolidays() {
    if (confirm('Are you sure you want to clear all holidays? This action cannot be undone.')) {
        holidays.clear();
        updateHolidayCalendar();
    }
}

// Apply holidays and refresh the main view
function applyHolidaysAndRefresh() {
    // Save holidays to storage
    saveDataToStorage();
    
    // Get current month selection from main interface
    const monthFilter = document.getElementById('monthFilter');
    if (monthFilter) {
        displayDataWithFilters(monthFilter.value);
    }
}

// Export all data as JSON
function exportAllData() {
    const exportData = {
        extractedData: extractedData,
        allMonths: allMonths,
        holidays: [...holidays],
        exportDate: new Date().toISOString(),
        version: '1.0'
    };
    
    const dataStr = JSON.stringify(exportData, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });
    
    const link = document.createElement('a');
    link.href = URL.createObjectURL(dataBlob);
    link.download = `fingerprint_data_backup_${new Date().toISOString().split('T')[0]}.json`;
    link.click();
    
    URL.revokeObjectURL(link.href);
}

// Import data from JSON file
function importData(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const importedData = JSON.parse(e.target.result);
            
            if (importedData.extractedData && Array.isArray(importedData.extractedData)) {
                // Confirm import
                if (confirm(`Import ${importedData.extractedData.length} records? This will merge with existing data.`)) {
                    // Convert date strings back to Date objects
                    importedData.extractedData.forEach(employee => {
                        employee.attendanceData.forEach(record => {
                            if (record.fullDate && typeof record.fullDate === 'string') {
                                record.fullDate = new Date(record.fullDate);
                            }
                        });
                    });
                    
                    // Merge imported data
                    if (extractedData.length === 0) {
                        extractedData = importedData.extractedData;
                    } else {
                        mergeEmployeeData(importedData.extractedData);
                    }
                    
                    // Import holidays if available
                    if (importedData.holidays && Array.isArray(importedData.holidays)) {
                        importedData.holidays.forEach(holiday => holidays.add(holiday));
                    }
                    
                    // Update months and save
                    updateMonthsList();
                    saveDataToStorage();
                    
                    // Refresh display
                    displayDataWithFilters();
                    showDataStatus();
                    
                    showError('Data imported successfully!');
                    setTimeout(() => hideError(), 3000);
                }
            } else {
                showError('Invalid data format. Please select a valid backup file.');
            }
        } catch (error) {
            console.error('Import error:', error);
            showError('Error importing data. Please check the file format.');
        }
    };
    
    reader.readAsText(file);
}

// Show analysis modal with daily attendance chart
function showAnalysis(selectedMonth) {
    let filteredData = extractedData;
    let monthLabel = 'All Data';
    
    if (selectedMonth !== 'all') {
        const [year, month] = selectedMonth.split('-');
        filteredData = extractedData.filter(record => {
            // Check if employee has attendance data for the selected month
            const hasAttendanceInMonth = record.attendanceData.some(attendance => 
                attendance.year == year && attendance.month == month
            );
            return hasAttendanceInMonth;
        });
        
        // Also filter the attendance data within each employee record
        filteredData = filteredData.map(record => ({
            ...record,
            attendanceData: record.attendanceData.filter(attendance => 
                attendance.year == year && attendance.month == month
            )
        }));
        
        monthLabel = new Date(selectedMonth + '-01').toLocaleString('default', { month: 'long', year: 'numeric' });
    }
    
    // Calculate daily attendance counts
    const dailyAttendance = {};
    
    filteredData.forEach(employee => {
        employee.attendanceData.forEach(record => {
            if (record.morning.in && record.dayOfWeek !== 'SUN' && record.dayOfWeek !== 'SAT') {
                const dateString = record.fullDate ? record.fullDate.toISOString().split('T')[0] : null;
                const isHolidayDate = dateString && isHoliday(dateString);
                
                if (!isHolidayDate) {
                    const dateKey = record.date;
                    if (!dailyAttendance[dateKey]) {
                        dailyAttendance[dateKey] = {
                            date: dateKey,
                            fullDate: record.fullDate,
                            dayOfWeek: record.dayOfWeek,
                            count: 0
                        };
                    }
                    dailyAttendance[dateKey].count++;
                }
            }
        });
    });
    
    // Sort dates and prepare chart data
    const sortedDates = Object.values(dailyAttendance).sort((a, b) => a.fullDate - b.fullDate);
    const labels = sortedDates.map(d => `${d.date} (${d.dayOfWeek})`);
    const data = sortedDates.map(d => d.count);
    
    // Generate random chart ID to avoid conflicts
    const chartId = 'attendanceChart_' + Math.random().toString(36).substr(2, 9);
    
    const modalHtml = `
        <div class="modal fade" id="analysisModal" tabindex="-1">
            <div class="modal-dialog modal-xl">
                <div class="modal-content border-0 shadow-lg analysis-modal">
                    <div class="modal-header border-0 bg-gradient-primary text-white">
                        <h5 class="modal-title fw-light">
                            <i class="bi bi-bar-chart me-2"></i>Daily Attendance Analysis
                            <small class="opacity-75 ms-2">${monthLabel}</small>
                        </h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body p-4" style="background: #1a1a1a; color: #e2e8f0;">
                        <div class="row g-3 mb-4">
                            <div class="col-md-3">
                                <div class="stats-card card-hover" data-animate="fadeInUp" data-delay="0">
                                    <div class="stats-icon bg-primary">
                                        <i class="bi bi-calendar-week"></i>
                                    </div>
                                    <div class="stats-content">
                                        <h3 class="stats-number text-primary">${sortedDates.length}</h3>
                                        <p class="stats-label">Working Days</p>
                                        <small class="stats-sublabel">Excl. weekends & holidays</small>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="stats-card card-hover" data-animate="fadeInUp" data-delay="100">
                                    <div class="stats-icon bg-success">
                                        <i class="bi bi-arrow-up"></i>
                                    </div>
                                    <div class="stats-content">
                                        <h3 class="stats-number text-success">${data.length > 0 ? Math.max(...data) : 0}</h3>
                                        <p class="stats-label">Peak Attendance</p>
                                        <small class="stats-sublabel">Maximum prefects</small>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="stats-card card-hover" data-animate="fadeInUp" data-delay="200">
                                    <div class="stats-icon bg-warning">
                                        <i class="bi bi-arrow-down"></i>
                                    </div>
                                    <div class="stats-content">
                                        <h3 class="stats-number text-warning">${data.length > 0 ? Math.min(...data) : 0}</h3>
                                        <p class="stats-label">Lowest Day</p>
                                        <small class="stats-sublabel">Minimum prefects</small>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="stats-card card-hover" data-animate="fadeInUp" data-delay="300">
                                    <div class="stats-icon bg-info">
                                        <i class="bi bi-graph-up"></i>
                                    </div>
                                    <div class="stats-content">
                                        <h3 class="stats-number text-info">${data.length > 0 ? Math.round(data.reduce((a, b) => a + b, 0) / data.length) : 0}</h3>
                                        <p class="stats-label">Daily Average</p>
                                        <small class="stats-sublabel">Typical attendance</small>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="row g-4">
                            <div class="col-12">
                                <div class="chart-container" data-animate="fadeInUp" data-delay="400">
                                    <div class="chart-header">
                                        <h6 class="chart-title">
                                            <i class="bi bi-bar-chart-fill me-2"></i>Attendance Trends
                                        </h6>
                                    </div>
                                    <div class="chart-wrapper">
                                        <canvas id="${chartId}"></canvas>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="row mt-4">
                            <div class="col-12">
                                <div class="data-table-container" data-animate="fadeInUp" data-delay="600">
                                    <div class="table-header">
                                        <h6 class="table-title">
                                            <i class="bi bi-table me-2"></i>Daily Breakdown
                                        </h6>
                                    </div>
                                    <div class="table-wrapper">
                                        <table class="minimal-table">
                                            <thead>
                                                <tr>
                                                    <th>Date</th>
                                                    <th>Day</th>
                                                    <th>Present</th>
                                                    <th>Rate</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                ${sortedDates.map((dayData, index) => {
                                                    const totalPrefects = [...new Set(filteredData.map(r => r.name))].length;
                                                    const attendanceRate = totalPrefects > 0 ? Math.round((dayData.count / totalPrefects) * 100) : 0;
                                                    return `
                                                    <tr class="table-row-animate" style="animation-delay: ${700 + (index * 50)}ms">
                                                        <td><span class="date-badge">${dayData.date}</span></td>
                                                        <td><span class="day-badge">${dayData.dayOfWeek}</span></td>
                                                        <td>
                                                            <span class="count-badge ${dayData.count >= Math.round(totalPrefects * 0.8) ? 'high' : 
                                                                dayData.count >= Math.round(totalPrefects * 0.6) ? 'medium' : 'low'}">${dayData.count}</span>
                                                        </td>
                                                        <td>
                                                            <div class="attendance-rate-container">
                                                                <div class="rate-info">
                                                                    <span class="rate-fraction">${dayData.count}/${totalPrefects}</span>
                                                                    <span class="rate-percentage">${attendanceRate}%</span>
                                                                </div>
                                                                <div class="progress-enhanced">
                                                                    <div class="progress-bar-enhanced ${attendanceRate >= 80 ? 'high' : 
                                                                        attendanceRate >= 60 ? 'medium' : 'low'}" 
                                                                        style="width: ${attendanceRate}%; animation-delay: ${800 + (index * 50)}ms">
                                                                        <div class="progress-shine"></div>
                                                                    </div>
                                                                </div>
                                                                <div class="rate-label">Attendance Rate</div>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                    `;
                                                }).join('')}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer border-0" style="background: #1a1a1a; border-top: 1px solid #2d2d2d;">
                        <button type="button" class="btn btn-minimal btn-download" onclick="downloadAnalysisCSV('${selectedMonth}')">
                            <i class="bi bi-download me-2"></i>Export Data
                        </button>
                        <button type="button" class="btn btn-minimal btn-secondary" data-bs-dismiss="modal">Close</button>
                    </div>
                </div>
            </div>
        </div>
        
        <style>
            /* Minimalistic Modal Styles */
            .analysis-modal {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
            }
            
            .bg-gradient-primary {
                background: linear-gradient(135deg, #065f46 0%, #022c22 100%);
            }
            
            /* Stats Cards */
            .stats-card {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 16px;
                padding: 24px;
                box-shadow: 0 2px 20px rgba(0,0,0,0.3);
                transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                position: relative;
                overflow: hidden;
            }
            
            .stats-card::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                height: 3px;
                background: linear-gradient(90deg, #10b981, #065f46);
                transform: scaleX(0);
                transition: transform 0.3s ease;
            }
            
            .card-hover:hover {
                transform: translateY(-4px);
                box-shadow: 0 8px 40px rgba(16, 185, 129, 0.2);
                border-color: #065f46;
            }
            
            .card-hover:hover::before {
                transform: scaleX(1);
            }
            
            .stats-icon {
                width: 48px;
                height: 48px;
                border-radius: 12px;
                display: flex;
                align-items: center;
                justify-content: center;
                margin-bottom: 16px;
                opacity: 0.9;
                background: linear-gradient(135deg, #10b981, #065f46);
            }
            
            .stats-icon i {
                font-size: 20px;
                color: white;
            }
            
            .stats-content {
                text-align: left;
            }
            
            .stats-number {
                font-size: 2.5rem;
                font-weight: 300;
                line-height: 1;
                margin-bottom: 8px;
                color: #10b981;
            }
            
            .stats-label {
                font-size: 1rem;
                font-weight: 500;
                margin-bottom: 4px;
                color: #a7f3d0;
            }
            
            .stats-sublabel {
                font-size: 0.875rem;
                color: #9ca3af;
                font-weight: 400;
            }
            
            /* Chart Container */
            .chart-container {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 16px;
                padding: 24px;
                box-shadow: 0 2px 20px rgba(0,0,0,0.3);
                transition: all 0.3s ease;
            }
            
            .chart-header {
                margin-bottom: 20px;
                padding-bottom: 16px;
                border-bottom: 1px solid #2d2d2d;
            }
            
            .chart-title {
                font-size: 1.125rem;
                font-weight: 500;
                color: #10b981;
                margin: 0;
            }
            
            .chart-wrapper {
                height: 350px;
                position: relative;
            }
            
            /* Data Table */
            .data-table-container {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 16px;
                padding: 24px;
                box-shadow: 0 2px 20px rgba(0,0,0,0.3);
                overflow: hidden;
            }
            
            .table-header {
                margin-bottom: 20px;
                padding-bottom: 16px;
                border-bottom: 1px solid #2d2d2d;
            }
            
            .table-title {
                font-size: 1.125rem;
                font-weight: 500;
                color: #10b981;
                margin: 0;
            }
            
            .table-wrapper {
                max-height: 400px;
                overflow-y: auto;
                overflow-x: hidden;
            }
            
            .minimal-table {
                width: 100%;
                border-collapse: collapse;
            }
            
            .minimal-table th {
                background: #0f0f0f;
                padding: 16px 12px;
                font-weight: 500;
                font-size: 0.875rem;
                color: #10b981;
                text-align: left;
                border: none;
                border-bottom: 1px solid #2d2d2d;
                position: sticky;
                top: 0;
                z-index: 10;
            }
            
            .minimal-table td {
                padding: 16px 12px;
                border: none;
                border-bottom: 1px solid #2d2d2d;
                font-size: 0.875rem;
                color: #e2e8f0;
            }
            
            .minimal-table tr:hover {
                background-color: #0f0f0f;
            }
            
            /* Badges */
            .date-badge {
                background: #374151;
                color: #10b981;
                padding: 6px 12px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 0.875rem;
                border: 1px solid #4b5563;
            }
            
            .day-badge {
                background: #065f46;
                color: #a7f3d0;
                padding: 4px 8px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 0.75rem;
                border: 1px solid #047857;
            }
            
            .count-badge {
                padding: 6px 12px;
                border-radius: 8px;
                font-weight: 600;
                font-size: 0.875rem;
            }
            
            .count-badge.high {
                background: #065f46;
                color: #a7f3d0;
                border: 1px solid #047857;
            }
            
            .count-badge.medium {
                background: #dc2626;
                color: #fef2f2;
                border: 1px solid #b91c1c;
            }
            
            .count-badge.low {
                background: #dc2626;
                color: #fef2f2;
                border: 1px solid #b91c1c;
            }
            
            /* Enhanced Progress Bar */
            .attendance-rate-container {
                display: flex;
                flex-direction: column;
                gap: 0.5rem;
                min-width: 120px;
            }
            
            .rate-info {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 0.25rem;
            }
            
            .rate-fraction {
                font-size: 0.75rem;
                font-weight: 500;
                color: #9ca3af;
                background: #374151;
                padding: 0.125rem 0.375rem;
                border-radius: 4px;
            }
            
            .rate-percentage {
                font-size: 0.875rem;
                font-weight: 600;
                color: #10b981;
            }
            
            .progress-enhanced {
                width: 100%;
                height: 16px;
                background: #374151;
                border-radius: 8px;
                overflow: hidden;
                position: relative;
                border: 1px solid #4b5563;
                box-shadow: inset 0 1px 3px rgba(0,0,0,0.3);
            }
            
            .progress-bar-enhanced {
                height: 100%;
                border-radius: 7px;
                transition: width 1.2s cubic-bezier(0.4, 0, 0.2, 1);
                position: relative;
                overflow: hidden;
                display: flex;
                align-items: center;
                justify-content: center;
            }
            
            .progress-bar-enhanced.high {
                background: linear-gradient(90deg, #10b981, #059669, #047857);
                box-shadow: 0 0 8px rgba(16, 185, 129, 0.3);
            }
            
            .progress-bar-enhanced.medium {
                background: linear-gradient(90deg, #f59e0b, #d97706, #b45309);
                box-shadow: 0 0 8px rgba(245, 158, 11, 0.3);
            }
            
            .progress-bar-enhanced.low {
                background: linear-gradient(90deg, #ef4444, #dc2626, #b91c1c);
                box-shadow: 0 0 8px rgba(239, 68, 68, 0.3);
            }
            
            .progress-shine {
                position: absolute;
                top: 0;
                left: -100%;
                width: 100%;
                height: 100%;
                background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
                animation: shine 2s infinite;
            }
            
            @keyframes shine {
                0% { left: -100%; }
                100% { left: 100%; }
            }
            
            .rate-label {
                font-size: 0.625rem;
                color: #6b7280;
                text-align: center;
                font-weight: 500;
                text-transform: uppercase;
                letter-spacing: 0.025em;
            }
            
            /* Legacy Progress Bar (keeping for compatibility) */
            .progress-minimal {
                width: 100%;
                height: 8px;
                background: #374151;
                border-radius: 4px;
                overflow: hidden;
                position: relative;
            }
            
            .progress-bar-minimal {
                height: 100%;
                border-radius: 4px;
                transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1);
                position: relative;
                overflow: hidden;
            }
            
            .progress-bar-minimal.high {
                background: linear-gradient(90deg, #10b981, #059669);
            }
            
            .progress-bar-minimal.medium {
                background: linear-gradient(90deg, #f59e0b, #d97706);
            }
            
            .progress-bar-minimal.low {
                background: linear-gradient(90deg, #ef4444, #dc2626);
            }
            
            .progress-text {
                position: absolute;
                right: 8px;
                top: 50%;
                transform: translateY(-50%);
                font-size: 0.75rem;
                font-weight: 600;
                color: white;
                text-shadow: 0 1px 2px rgba(0,0,0,0.3);
            }
            
            /* Mobile Responsive Styles for Enhanced Progress Bar */
            @media (max-width: 768px) {
                .attendance-rate-container {
                    min-width: 100px;
                    gap: 0.375rem;
                }
                
                .rate-info {
                    margin-bottom: 0.125rem;
                }
                
                .rate-fraction {
                    font-size: 0.625rem;
                    padding: 0.125rem 0.25rem;
                }
                
                .rate-percentage {
                    font-size: 0.75rem;
                }
                
                .progress-enhanced {
                    height: 14px;
                }
                
                .rate-label {
                    font-size: 0.5rem;
                }
            }
            
            @media (max-width: 480px) {
                .attendance-rate-container {
                    min-width: 80px;
                    gap: 0.25rem;
                }
                
                .rate-info {
                    flex-direction: column;
                    align-items: flex-start;
                    gap: 0.125rem;
                }
                
                .rate-fraction {
                    font-size: 0.5rem;
                    align-self: center;
                }
                
                .rate-percentage {
                    font-size: 0.625rem;
                    align-self: center;
                }
                
                .progress-enhanced {
                    height: 12px;
                }
                
                .rate-label {
                    font-size: 0.45rem;
                }
            }
            
            /* Buttons */
            .btn-minimal {
                border: none;
                border-radius: 10px;
                padding: 12px 24px;
                font-weight: 500;
                transition: all 0.2s ease;
                position: relative;
                overflow: hidden;
            }
            
            .btn-download {
                background: linear-gradient(135deg, #10b981 0%, #065f46 100%);
                color: white;
                border: 1px solid #065f46;
            }
            
            .btn-download:hover {
                transform: translateY(-1px);
                box-shadow: 0 4px 12px rgba(16, 185, 129, 0.4);
                color: white;
            }
            
            .btn-secondary {
                background: #374151;
                color: #10b981;
                border: 1px solid #374151;
            }
            
            .btn-secondary:hover {
                background: #4b5563;
                color: #a7f3d0;
            }
            
            /* Animations */
            @keyframes fadeInUp {
                from {
                    opacity: 0;
                    transform: translateY(20px);
                }
                to {
                    opacity: 1;
                    transform: translateY(0);
                }
            }
            
            @keyframes slideInLeft {
                from {
                    opacity: 0;
                    transform: translateX(-20px);
                }
                to {
                    opacity: 1;
                    transform: translateX(0);
                }
            }
            
            [data-animate="fadeInUp"] {
                animation: fadeInUp 0.6s cubic-bezier(0.4, 0, 0.2, 1) forwards;
                opacity: 0;
            }
            
            .table-row-animate {
                animation: slideInLeft 0.4s cubic-bezier(0.4, 0, 0.2, 1) forwards;
                opacity: 0;
            }
            
            /* Animation delays */
            [data-delay="0"] { animation-delay: 0ms; }
            [data-delay="100"] { animation-delay: 100ms; }
            [data-delay="200"] { animation-delay: 200ms; }
            [data-delay="300"] { animation-delay: 300ms; }
            [data-delay="400"] { animation-delay: 400ms; }
            [data-delay="500"] { animation-delay: 500ms; }
            [data-delay="600"] { animation-delay: 600ms; }
            
            /* Scrollbar styling */
            .table-wrapper::-webkit-scrollbar {
                width: 6px;
            }
            
            .table-wrapper::-webkit-scrollbar-track {
                background: #2d2d2d;
            }
            
            .table-wrapper::-webkit-scrollbar-thumb {
                background: #065f46;
                border-radius: 3px;
            }
            
            .table-wrapper::-webkit-scrollbar-thumb:hover {
                background: #10b981;
            }
        </style>
    `;
    
    // Remove existing modal if any
    const existingModal = document.getElementById('analysisModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Add modal to body
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    
    // Show modal and create chart
    const modal = new bootstrap.Modal(document.getElementById('analysisModal'));
    
    // Handle modal events for accessibility
    const modalElement = document.getElementById('analysisModal');
    
    // Remove focus from any focused elements before showing modal
    if (document.activeElement) {
        document.activeElement.blur();
    }
    
    // Wait for modal to be shown, then create chart and handle focus
    modalElement.addEventListener('shown.bs.modal', function () {
        createAttendanceChart(chartId, labels, data);
        
        // Focus the first focusable element in the modal
        const firstFocusable = modalElement.querySelector('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
        if (firstFocusable) {
            firstFocusable.focus();
        }
    });
    
    modalElement.addEventListener('hidden.bs.modal', function () {
        // Clean up the modal element after it's hidden
        setTimeout(() => {
            if (modalElement && modalElement.parentNode) {
                modalElement.remove();
            }
        }, 150);
    });
    
    modal.show();
}

// Create attendance chart using Chart.js
function createAttendanceChart(chartId, labels, data) {
    // Load Chart.js if not already loaded
    if (typeof Chart === 'undefined') {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/chart.js';
        script.onload = function() {
            renderChart(chartId, labels, data);
        };
        document.head.appendChild(script);
    } else {
        renderChart(chartId, labels, data);
    }
}

// Render the chart
function renderChart(chartId, labels, data) {
    const ctx = document.getElementById(chartId).getContext('2d');
    
    // Destroy existing chart if it exists
    if (window.attendanceChart) {
        window.attendanceChart.destroy();
    }
    
    // Create gradient
    const gradient = ctx.createLinearGradient(0, 0, 0, 350);
    gradient.addColorStop(0, 'rgba(102, 126, 234, 0.8)');
    gradient.addColorStop(1, 'rgba(102, 126, 234, 0.1)');
    
    window.attendanceChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Prefects Present',
                data: data,
                backgroundColor: gradient,
                borderColor: 'rgba(102, 126, 234, 1)',
                borderWidth: 2,
                borderRadius: 8,
                borderSkipped: false,
                barThickness: 'flex',
                maxBarThickness: 40,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: {
                padding: {
                    top: 20,
                    bottom: 10
                }
            },
            plugins: {
                title: {
                    display: false
                },
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(45, 55, 72, 0.95)',
                    titleColor: 'white',
                    bodyColor: 'white',
                    borderColor: 'rgba(102, 126, 234, 0.8)',
                    borderWidth: 1,
                    cornerRadius: 8,
                    displayColors: false,
                    titleFont: {
                        size: 13,
                        weight: '500'
                    },
                    bodyFont: {
                        size: 12
                    },
                    padding: 12,
                    callbacks: {
                        title: function(context) {
                            return context[0].label;
                        },
                        label: function(context) {
                            return `${context.parsed.y} prefects attended`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        display: false
                    },
                    border: {
                        display: false
                    },
                    ticks: {
                        color: '#718096',
                        font: {
                            size: 11,
                            weight: '500'
                        },
                        maxRotation: 45,
                        minRotation: 45,
                        padding: 10
                    }
                },
                y: {
                    grid: {
                        color: 'rgba(226, 232, 240, 0.8)',
                        lineWidth: 1
                    },
                    border: {
                        display: false
                    },
                    ticks: {
                        color: '#718096',
                        font: {
                            size: 11,
                            weight: '500'
                        },
                        beginAtZero: true,
                        stepSize: 1,
                        padding: 10,
                        callback: function(value) {
                            return Math.floor(value);
                        }
                    }
                }
            },
            interaction: {
                intersect: false,
                mode: 'index',
            },
            animation: {
                duration: 1200,
                easing: 'easeInOutCubic',
                onProgress: function(animation) {
                    // Add a subtle loading effect during animation
                    const progress = animation.currentStep / animation.numSteps;
                    ctx.globalAlpha = 0.1 + (0.9 * progress);
                },
                onComplete: function() {
                    ctx.globalAlpha = 1;
                }
            },
            hover: {
                animationDuration: 200
            }
        }
    });
}

// Download analysis data as CSV
function downloadAnalysisCSV(selectedMonth) {
    let filteredData = extractedData;
    let filename = 'daily_attendance_analysis';
    
    if (selectedMonth !== 'all') {
        const [year, month] = selectedMonth.split('-');
        filteredData = extractedData.filter(record => {
            const hasAttendanceInMonth = record.attendanceData.some(attendance => 
                attendance.year == year && attendance.month == month
            );
            return hasAttendanceInMonth;
        });
        
        filteredData = filteredData.map(record => ({
            ...record,
            attendanceData: record.attendanceData.filter(attendance => 
                attendance.year == year && attendance.month == month
            )
        }));
        
        filename += `_${selectedMonth}`;
    }
    
    // Calculate daily attendance counts
    const dailyAttendance = {};
    
    filteredData.forEach(employee => {
        employee.attendanceData.forEach(record => {
            if (record.morning.in && record.dayOfWeek !== 'SUN' && record.dayOfWeek !== 'SAT') {
                const dateString = record.fullDate ? record.fullDate.toISOString().split('T')[0] : null;
                const isHolidayDate = dateString && isHoliday(dateString);
                
                if (!isHolidayDate) {
                    const dateKey = record.date;
                    if (!dailyAttendance[dateKey]) {
                        dailyAttendance[dateKey] = {
                            date: dateKey,
                            fullDate: record.fullDate,
                            dayOfWeek: record.dayOfWeek,
                            count: 0,
                            prefects: []
                        };
                    }
                    dailyAttendance[dateKey].count++;
                    dailyAttendance[dateKey].prefects.push(employee.name);
                }
            }
        });
    });
    
    const sortedDates = Object.values(dailyAttendance).sort((a, b) => a.fullDate - b.fullDate);
    const totalPrefects = [...new Set(filteredData.map(r => r.name))].length;
    
    const csvRows = [
        ['Date', 'Day', 'Prefects Present', 'Total Prefects', 'Attendance Rate (%)', 'Present Prefects']
    ];
    
    sortedDates.forEach(dayData => {
        const attendanceRate = totalPrefects > 0 ? Math.round((dayData.count / totalPrefects) * 100) : 0;
        csvRows.push([
            dayData.date,
            dayData.dayOfWeek,
            dayData.count,
            totalPrefects,
            attendanceRate,
            dayData.prefects.join('; ')
        ]);
    });
    
    const csvContent = csvRows.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
    downloadCSVFile(csvContent, `${filename}.csv`);
}

// Helper function to download CSV files
function downloadCSVFile(csvContent, filename) {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    
    if (link.download !== undefined) {
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

// Search functionality
let searchTimeout;

// Debounced search function
function debouncedSearch() {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(searchPrefects, 150);
}

function searchPrefects() {
    const searchInput = document.getElementById('prefectSearch');
    if (!searchInput) {
        console.log('Search input not found');
        return;
    }
    
    const searchTerm = searchInput.value.toLowerCase().trim();
    const tableRows = document.querySelectorAll('.data-table tbody tr');
    
    console.log('Search term:', searchTerm);
    console.log('Found rows:', tableRows.length);
    
    if (tableRows.length === 0) {
        console.log('No table rows found - table might not be loaded yet');
        return;
    }
    
    let visibleCount = 0;
    
    tableRows.forEach((row, index) => {
        const prefectName = row.querySelector('.prefect-name');
        if (prefectName) {
            const name = prefectName.textContent.toLowerCase().trim();
            const shouldShow = searchTerm === '' || name.includes(searchTerm);
            
            if (shouldShow) {
                row.style.display = '';
                row.style.opacity = '1';
                row.style.animation = 'fadeIn 0.3s ease';
                visibleCount++;
            } else {
                row.style.display = 'none';
                row.style.opacity = '0';
            }
        }
    });
    
    // Update row numbers for visible rows
    let visibleIndex = 1;
    tableRows.forEach(row => {
        if (row.style.display !== 'none') {
            const rowNumber = row.querySelector('.row-number');
            if (rowNumber) {
                rowNumber.textContent = visibleIndex++;
            }
        }
    });
    
    console.log(`Search completed. ${visibleCount} rows visible out of ${tableRows.length}`);
    
    // Update search result indicator
    updateSearchResultIndicator(visibleCount, tableRows.length, searchTerm);
}

// Add search result indicator
function updateSearchResultIndicator(visibleCount, totalCount, searchTerm) {
    let indicator = document.getElementById('searchResultIndicator');
    
    if (!indicator) {
        // Create indicator if it doesn't exist
        const searchContainer = document.querySelector('.search-container');
        if (searchContainer) {
            indicator = document.createElement('div');
            indicator.id = 'searchResultIndicator';
            indicator.className = 'search-result-indicator';
            searchContainer.appendChild(indicator);
        }
    }
    
    if (indicator) {
        if (searchTerm && visibleCount !== totalCount) {
            indicator.textContent = `${visibleCount} of ${totalCount} prefects`;
            indicator.style.display = 'block';
        } else {
            indicator.style.display = 'none';
        }
    }
}
