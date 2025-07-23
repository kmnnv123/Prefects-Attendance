// Global variables
let workbook = null;
let currentSheet = null;
let extractedData = [];
let allMonths = [];
let holidays = new Set(); // Store holidays as 'YYYY-MM-DD' strings

// Firebase variables
let database = null;
let isFirebaseReady = false;

// Session management variables
let currentSessionId = null;
let passwordChangeListener = null;
let sessionCheckInterval = null;

// Authentication System
async function attemptLogin() {
    const passwordInput = document.getElementById('loginPassword');
    const enteredPassword = passwordInput.value;
    const loginError = document.getElementById('loginError');
    
    console.log('🔐 Login attempt initiated');
    console.log('🌐 Current URL:', window.location.href);
    console.log('🔒 Crypto.subtle available:', !!window.crypto?.subtle);
    console.log('📍 Protocol:', window.location.protocol);
    
    if (!enteredPassword) {
        showLoginError('Please enter a password');
        return;
    }
    
    try {
        // Get stored password hash from Firebase or use default
        const storedPasswordHash = await getStoredPasswordHash();
        console.log('💾 Using stored hash from Firebase:', storedPasswordHash.substring(0, 10) + '...');
        
        // Hash the entered password
        console.log('🔄 Hashing entered password...');
        const enteredPasswordHash = await hashPassword(enteredPassword);
        console.log('✅ Generated hash:', enteredPasswordHash.substring(0, 10) + '...');
        console.log('🔍 Hash match:', enteredPasswordHash === storedPasswordHash);
        
        if (enteredPasswordHash === storedPasswordHash) {
            // Successful login
            console.log('✅ Login successful');
            
            // Start session monitoring
            await startSessionMonitoring();
            
            hideLoginScreen();
            await initializeMainApp();
        } else {
            console.log('❌ Login failed - hash mismatch');
            showLoginError('Invalid password. Contact board administration for access.');
            passwordInput.value = '';
            passwordInput.style.animation = 'shake 0.5s ease-in-out';
            setTimeout(() => {
                passwordInput.style.animation = '';
            }, 500);
        }
    } catch (error) {
        console.error('💥 Login error:', error);
        showLoginError('Login system error. Please try again.');
    }
}

function showLoginError(message) {
    const loginError = document.getElementById('loginError');
    
    if (loginError) {
        loginError.textContent = message;
        loginError.style.display = 'block';
        
        setTimeout(() => {
            loginError.style.display = 'none';
        }, 4000);
    }
}

function hideLoginScreen() {
    const loginScreen = document.getElementById('loginScreen');
    const mainApp = document.getElementById('mainApp');
    
    if (!loginScreen || !mainApp) {
        console.error('Cannot find required elements for login screen transition');
        return;
    }
    
    // Fade out login screen
    loginScreen.style.transition = 'opacity 0.5s ease-out';
    loginScreen.style.opacity = '0';
    
    setTimeout(() => {
        loginScreen.style.display = 'none';
        mainApp.style.display = 'block';
        
        // Fade in main app
        mainApp.style.opacity = '0';
        mainApp.style.transition = 'opacity 0.5s ease-in';
        setTimeout(() => {
            mainApp.style.opacity = '1';
        }, 50);
    }, 500);
}

async function initializeMainApp() {
    console.log('🚀 App starting - DOM loaded');
    
    // Show loading indicator
    showLoadingIndicator('Connecting to BOP25 Server...');
    
    try {
        console.log('📡 Initializing Firebase connection...');
        await initializeFirebase();
        
        updateLoadingIndicator('Loading your data from BOP25 Server...');
        console.log('🔄 Loading data from Firebase...');
        await loadDataFromFirebase();
        
        updateLoadingIndicator('Starting application...');
        console.log('✨ Firebase data loaded, starting app...');
        await initializeApp();
        console.log('🎉 App initialization complete');
        
        hideLoadingIndicator();
    } catch (error) {
        console.error('❌ App initialization error:', error);
        showErrorIndicator('Failed to connect to BOP25 Server. Using offline mode.');
        setTimeout(hideLoadingIndicator, 3000);
    }
}

function logout() {
    // Clear any sensitive data from memory
    extractedData = [];
    allMonths = [];
    holidays = new Set();
    
    // Hide main app and show login screen
    const loginScreen = document.getElementById('loginScreen');
    const mainApp = document.getElementById('mainApp');
    
    // Fade out main app
    mainApp.style.transition = 'opacity 0.3s ease-out';
    mainApp.style.opacity = '0';
    
    setTimeout(() => {
        mainApp.style.display = 'none';
        loginScreen.style.display = 'flex';
        loginScreen.style.opacity = '0';
        
        // Clear results section
        const resultsDiv = document.getElementById('results');
        if (resultsDiv) {
            resultsDiv.innerHTML = '';
        }
        
        // Clear login password and focus on input
        const passwordInput = document.getElementById('loginPassword');
        if (passwordInput) {
            passwordInput.value = '';
            passwordInput.focus();
        }
        
        // Fade in login screen
        loginScreen.style.transition = 'opacity 0.3s ease-in';
        setTimeout(() => {
            loginScreen.style.opacity = '1';
        }, 50);
    }, 300);
}

// Hash password function (used by both login and password change)
async function hashPassword(password) {
    const encoder = new TextEncoder();
    const data = encoder.encode(password);
    const hash = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hash));
    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    return hashHex;
}

// Firebase Password Management
async function getStoredPasswordHash() {
    try {
        if (!isFirebaseReady) {
            await initializeFirebase();
        }
        
        const snapshot = await firebase.database().ref('system/adminPassword').once('value');
        const storedHash = snapshot.val();
        
        if (storedHash) {
            console.log('📱 Password hash retrieved from Firebase');
            return storedHash;
        } else {
            // No password stored yet, use default and store it
            const defaultHash = "f16d847d2182c5d0be13cf3263a4898c2849e96c9beb04b9694258b0d7eb080e"; // Default: PrefectsAdmin2025!
            console.log('🔑 No password in Firebase, setting default');
            await setStoredPasswordHash(defaultHash);
            return defaultHash;
        }
    } catch (error) {
        console.error('❌ Error getting password from Firebase:', error);
        // Fallback to default if Firebase fails
        return "f16d847d2182c5d0be13cf3263a4898c2849e96c9beb04b9694258b0d7eb080e";
    }
}

async function setStoredPasswordHash(passwordHash) {
    try {
        if (!isFirebaseReady) {
            await initializeFirebase();
        }
        
        // Set the new password hash
        await firebase.database().ref('system/adminPassword').set(passwordHash);
        console.log('✅ Password hash stored in Firebase');
        
        // Force logout all active sessions by updating the session invalidation timestamp
        await invalidateAllSessions();
        
        return true;
    } catch (error) {
        console.error('❌ Error storing password in Firebase:', error);
        return false;
    }
}

// Session Management Functions
function generateSessionId() {
    return 'session_' + Date.now() + '_' + Math.random().toString(36).substring(2, 15);
}

async function invalidateAllSessions() {
    try {
        if (!isFirebaseReady) {
            await initializeFirebase();
        }
        
        const invalidationTimestamp = Date.now();
        await firebase.database().ref('system/sessionInvalidation').set(invalidationTimestamp);
        console.log('🚪 All sessions invalidated at:', new Date(invalidationTimestamp).toLocaleString());
        return true;
    } catch (error) {
        console.error('❌ Error invalidating sessions:', error);
        return false;
    }
}

async function checkSessionValidity() {
    try {
        if (!isFirebaseReady || !currentSessionId) {
            return false;
        }
        
        // Get the session invalidation timestamp
        const snapshot = await firebase.database().ref('system/sessionInvalidation').once('value');
        const invalidationTimestamp = snapshot.val();
        
        if (!invalidationTimestamp) {
            return true; // No invalidation timestamp set yet
        }
        
        // Extract session timestamp from session ID
        const sessionTimestamp = parseInt(currentSessionId.split('_')[1]);
        
        // If session was created before the invalidation timestamp, it's invalid
        if (sessionTimestamp < invalidationTimestamp) {
            console.log('⚠️ Session invalid - password was changed after login');
            return false;
        }
        
        return true;
    } catch (error) {
        console.error('❌ Error checking session validity:', error);
        return false;
    }
}

async function startSessionMonitoring() {
    // Generate a new session ID for this login
    currentSessionId = generateSessionId();
    console.log('🆔 Session started:', currentSessionId);
    
    // Listen for password changes in real-time
    if (passwordChangeListener) {
        passwordChangeListener.off(); // Remove previous listener
    }
    
    passwordChangeListener = firebase.database().ref('system/sessionInvalidation').on('value', async (snapshot) => {
        const invalidationTimestamp = snapshot.val();
        
        if (invalidationTimestamp && currentSessionId) {
            const sessionTimestamp = parseInt(currentSessionId.split('_')[1]);
            
            if (sessionTimestamp < invalidationTimestamp) {
                console.log('🚪 Password changed - forcing logout');
                await forceLogout('Password has been changed by administrator. Please log in again.');
            }
        }
    });
    
    // Also check periodically as backup
    if (sessionCheckInterval) {
        clearInterval(sessionCheckInterval);
    }
    
    sessionCheckInterval = setInterval(async () => {
        const isValid = await checkSessionValidity();
        if (!isValid) {
            await forceLogout('Session expired. Please log in again.');
        }
    }, 30000); // Check every 30 seconds
}

async function forceLogout(message = 'You have been logged out.') {
    console.log('🚪 Forcing logout:', message);
    
    // Clear session data
    currentSessionId = null;
    
    // Remove listeners
    if (passwordChangeListener) {
        passwordChangeListener.off();
        passwordChangeListener = null;
    }
    
    if (sessionCheckInterval) {
        clearInterval(sessionCheckInterval);
        sessionCheckInterval = null;
    }
    
    // Show logout message
    if (message) {
        alert(message);
    }
    
    // Redirect to login screen
    showLoginScreen();
}

function showLoginScreen() {
    const loginScreen = document.getElementById('loginScreen');
    const mainApp = document.getElementById('mainApp');
    
    if (loginScreen && mainApp) {
        mainApp.style.display = 'none';
        loginScreen.style.display = 'flex';
        
        // Clear the password input
        const passwordInput = document.getElementById('loginPassword');
        if (passwordInput) {
            passwordInput.value = '';
            passwordInput.focus();
        }
    }
}

// Clean up any legacy password data from localStorage
function cleanupLegacyPasswordStorage() {
    try {
        // Remove any old password-related localStorage items
        const legacyPasswordKeys = [
            'boardAdminPasswordHash',
            'adminPassword',
            'password',
            'loginPassword',
            'boardPassword'
        ];
        
        legacyPasswordKeys.forEach(key => {
            if (localStorage.getItem(key)) {
                console.log(`🧹 Removing legacy password data: ${key}`);
                localStorage.removeItem(key);
            }
        });
        
        console.log('✅ Legacy password cleanup completed');
    } catch (error) {
        console.error('❌ Error during legacy password cleanup:', error);
    }
}

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

// Wait for DOM to load before showing login
document.addEventListener('DOMContentLoaded', async function() {
    console.log('🚀 App loaded - checking for bypass or showing login screen');
    
    // Clean up any legacy password storage
    cleanupLegacyPasswordStorage();
    
    // Check for bypass parameter
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('bypass') === 'true') {
        console.log('🔓 Bypass mode activated - skipping authentication');
        
        // Start session monitoring even in bypass mode
        await startSessionMonitoring();
        
        hideLoginScreen();
        initializeMainApp();
        return;
    }
    
    // Focus on password input
    setTimeout(() => {
        const passwordInput = document.getElementById('loginPassword');
        if (passwordInput) {
            passwordInput.focus();
        }
    }, 100);
});

// Clean up session monitoring when page unloads
window.addEventListener('beforeunload', () => {
    if (passwordChangeListener) {
        passwordChangeListener.off();
    }
    if (sessionCheckInterval) {
        clearInterval(sessionCheckInterval);
    }
});

// Initialize Firebase
async function initializeFirebase() {
    try {
        if (typeof firebase !== 'undefined') {
            console.log('🔥 Initializing Firebase...');
            firebase.initializeApp(firebaseConfig);
            database = firebase.database();
            isFirebaseReady = true;
            console.log('✅ Firebase initialized successfully');
            
            // Test connection
            await testFirebaseConnection();
            console.log('🔗 Firebase connection test completed');
        } else {
            console.warn('Firebase SDK not loaded');
        }
    } catch (error) {
        console.error('Firebase initialization failed:', error);
        isFirebaseReady = false;
    }
}

// Loading indicator functions
function showLoadingIndicator(message) {
    const loadingDiv = document.getElementById('loading');
    const loadingText = loadingDiv.querySelector('span');
    if (loadingText) {
        loadingText.textContent = message;
    }
    loadingDiv.style.display = 'flex';
}

function updateLoadingIndicator(message) {
    const loadingDiv = document.getElementById('loading');
    const loadingText = loadingDiv.querySelector('span');
    if (loadingText) {
        loadingText.textContent = message;
    }
}

function hideLoadingIndicator() {
    const loadingDiv = document.getElementById('loading');
    loadingDiv.style.display = 'none';
}

function showErrorIndicator(message) {
    const errorDiv = document.getElementById('error');
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
}

function hideErrorIndicator() {
    const errorDiv = document.getElementById('error');
    errorDiv.style.display = 'none';
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

// Secure data deletion functions
function confirmDeleteAllData() {
    // Create a custom modal for password input
    const modal = document.createElement('div');
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.8);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 10000;
        backdrop-filter: blur(5px);
    `;
    
    modal.innerHTML = `
        <div style="
            background: #1a1a1a;
            border: 2px solid #dc2626;
            border-radius: 16px;
            padding: 2rem;
            max-width: 400px;
            width: 90%;
            box-shadow: 0 20px 60px rgba(220, 38, 38, 0.3);
            text-align: center;
        ">
            <div style="color: #dc2626; font-size: 3rem; margin-bottom: 1rem;">
                <i class="bi bi-exclamation-triangle"></i>
            </div>
            <h3 style="color: #dc2626; margin-bottom: 1rem; font-weight: 600;">
                🔄 NEW BOARD TRANSITION
            </h3>
            <p style="color: #e2e8f0; margin-bottom: 1.5rem; line-height: 1.5;">
                This will clear all data from the <strong>previous board</strong> to prepare the database for the new administration.
                <br><br>
                <strong>This action is irreversible!</strong>
            </p>
            <div style="margin-bottom: 1.5rem;">
                <label style="display: block; color: #e2e8f0; margin-bottom: 0.5rem; font-weight: 500;">
                    Enter Board Transition Password:
                </label>
                <input 
                    type="password" 
                    id="deletePassword" 
                    placeholder="Board admin password required"
                    style="
                        width: 100%;
                        padding: 0.75rem;
                        border: 2px solid #374151;
                        border-radius: 8px;
                        background: #0a0a0a;
                        color: #e2e8f0;
                        font-size: 1rem;
                        text-align: center;
                    "
                    onkeypress="if(event.key==='Enter') attemptDeleteAllData()"
                >
            </div>
            <div style="display: flex; gap: 1rem; justify-content: center;">
                <button 
                    onclick="closeDeleteModal()"
                    style="
                        padding: 0.75rem 1.5rem;
                        border: 1px solid #374151;
                        border-radius: 8px;
                        background: #374151;
                        color: #e2e8f0;
                        cursor: pointer;
                        font-weight: 500;
                    "
                >
                    Cancel
                </button>
                <button 
                    onclick="attemptDeleteAllData()"
                    style="
                        padding: 0.75rem 1.5rem;
                        border: 1px solid #dc2626;
                        border-radius: 8px;
                        background: #dc2626;
                        color: white;
                        cursor: pointer;
                        font-weight: 500;
                    "
                >
                    <i class="bi bi-trash"></i> RESET FOR NEW BOARD
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    modal.id = 'deleteModal';
    
    // Focus on password input
    setTimeout(() => {
        document.getElementById('deletePassword').focus();
    }, 100);
}

function closeDeleteModal() {
    const modal = document.getElementById('deleteModal');
    if (modal) {
        modal.remove();
    }
}

async function attemptDeleteAllData() {
    const passwordInput = document.getElementById('deletePassword');
    const enteredPassword = passwordInput.value;
    
    // Get stored password hash from Firebase
    const storedPasswordHash = await getStoredPasswordHash();
    
    // Hash the entered password
    const enteredPasswordHash = await hashPassword(enteredPassword);
    
    if (enteredPasswordHash === storedPasswordHash) {
        closeDeleteModal();
        showDeleteConfirmation();
    } else {
        // Shake the input field and show error
        passwordInput.style.borderColor = '#dc2626';
        passwordInput.style.animation = 'shake 0.5s ease-in-out';
        passwordInput.value = '';
        passwordInput.placeholder = '❌ Wrong password! Contact current board admin...';
        
        setTimeout(() => {
            passwordInput.style.borderColor = '#374151';
            passwordInput.style.animation = '';
            passwordInput.placeholder = 'Board admin password required';
        }, 3000);
    }
}

function showDeleteConfirmation() {
    const modal = document.createElement('div');
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.9);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 10000;
        backdrop-filter: blur(5px);
    `;
    
    modal.innerHTML = `
        <div style="
            background: #1a1a1a;
            border: 2px solid #dc2626;
            border-radius: 16px;
            padding: 2rem;
            max-width: 450px;
            width: 90%;
            box-shadow: 0 20px 60px rgba(220, 38, 38, 0.5);
            text-align: center;
        ">
            <div style="color: #dc2626; font-size: 4rem; margin-bottom: 1rem;">
                <i class="bi bi-fire"></i>
            </div>
            <h3 style="color: #dc2626; margin-bottom: 1rem; font-weight: 700;">
                BOARD TRANSITION CONFIRMATION
            </h3>
            <p style="color: #e2e8f0; margin-bottom: 1.5rem; line-height: 1.6; font-size: 1.1rem;">
                You are about to <strong style="color: #dc2626;">PERMANENTLY CLEAR</strong> all previous board data:
                <br><br>
                • Previous prefects' attendance records<br>
                • Previous board member data<br>
                • All historical months data<br>
                • Previous holidays configuration<br>
                <br>
                <strong style="color: #fbbf24;">The new board can start fresh!</strong>
            </p>
            <div style="display: flex; gap: 1rem; justify-content: center;">
                <button 
                    onclick="closeFinalModal()"
                    style="
                        padding: 1rem 2rem;
                        border: 2px solid #10b981;
                        border-radius: 8px;
                        background: #10b981;
                        color: white;
                        cursor: pointer;
                        font-weight: 600;
                        font-size: 1rem;
                    "
                >
                    <i class="bi bi-shield-check"></i> Keep Previous Board Data
                </button>
                <button 
                    onclick="executeDeleteAllData()"
                    style="
                        padding: 1rem 2rem;
                        border: 2px solid #dc2626;
                        border-radius: 8px;
                        background: #dc2626;
                        color: white;
                        cursor: pointer;
                        font-weight: 600;
                        font-size: 1rem;
                        animation: pulse 2s infinite;
                    "
                >
                    <i class="bi bi-arrow-repeat"></i> CLEAR FOR NEW BOARD
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    modal.id = 'finalModal';
}

function closeFinalModal() {
    const modal = document.getElementById('finalModal');
    if (modal) {
        modal.remove();
    }
}

async function executeDeleteAllData() {
    closeFinalModal();
    
    if (!isFirebaseReady || !database) {
        showError('Firebase is not connected. Cannot delete data.');
        return;
    }
    
    try {
        showLoading(true, 'Clearing previous board data from BOP25 Server...');
        
        // Delete all data from Firebase Realtime Database
        await database.ref('fingerprintData').remove();
        
        showLoading(true, 'Resetting local storage...');
        
        // Clear local data
        extractedData = [];
        allMonths = [];
        holidays = new Set();
        
        // Clear localStorage
        localStorage.removeItem('fingerprintData');
        localStorage.removeItem('extractedData');
        localStorage.removeItem('allMonths');
        localStorage.removeItem('holidays');
        
        showLoading(true, 'Board transition completed!');
        
        // Clear the display
        const resultsDiv = document.getElementById('results');
        if (resultsDiv) {
            resultsDiv.innerHTML = '';
        }
        
        console.log('✅ Previous board data successfully cleared - ready for new board');
        
        setTimeout(() => {
            showLoading(false);
            showSuccessMessage('� Database cleared! Ready for new board data.');
        }, 1500);
        
    } catch (error) {
        console.error('❌ Error deleting data:', error);
        showLoading(false);
        showError('Failed to clear previous board data from BOP25 Server. Please try again.');
    }
}

function showSuccessMessage(message) {
    const successDiv = document.createElement('div');
    successDiv.style.cssText = `
        position: fixed;
        top: 20px;
        left: 50%;
        transform: translateX(-50%);
        background: #10b981;
        color: white;
        padding: 1rem 2rem;
        border-radius: 8px;
        box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
        z-index: 10000;
        font-weight: 600;
        text-align: center;
    `;
    successDiv.textContent = message;
    document.body.appendChild(successDiv);
    
    setTimeout(() => {
        successDiv.remove();
    }, 3000);
}

// Password Change Functionality
function changePassword() {
    const modal = document.createElement('div');
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.8);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 10000;
        backdrop-filter: blur(5px);
    `;
    
    modal.innerHTML = `
        <div style="
            background: #1a1a1a;
            border: 2px solid #10b981;
            border-radius: 16px;
            padding: 2rem;
            max-width: 450px;
            width: 90%;
            box-shadow: 0 20px 60px rgba(16, 185, 129, 0.3);
            text-align: center;
        ">
            <div style="color: #10b981; font-size: 3rem; margin-bottom: 1rem;">
                <i class="bi bi-key"></i>
            </div>
            <h3 style="color: #10b981; margin-bottom: 1rem; font-weight: 600;">
                Change Board Admin Password
            </h3>
            <p style="color: #e2e8f0; margin-bottom: 1.5rem; line-height: 1.5;">
                Enter your current password and set a new one for board transitions.
            </p>
            
            <div style="margin-bottom: 1rem; text-align: left;">
                <label style="display: block; color: #e2e8f0; margin-bottom: 0.5rem; font-weight: 500;">
                    Current Password:
                </label>
                <input 
                    type="password" 
                    id="currentPassword" 
                    placeholder="Enter current password"
                    style="
                        width: 100%;
                        padding: 0.75rem;
                        border: 2px solid #374151;
                        border-radius: 8px;
                        background: #0a0a0a;
                        color: #e2e8f0;
                        font-size: 1rem;
                        margin-bottom: 1rem;
                    "
                >
                
                <label style="display: block; color: #e2e8f0; margin-bottom: 0.5rem; font-weight: 500;">
                    New Password:
                </label>
                <input 
                    type="password" 
                    id="newPassword" 
                    placeholder="Enter new password"
                    style="
                        width: 100%;
                        padding: 0.75rem;
                        border: 2px solid #374151;
                        border-radius: 8px;
                        background: #0a0a0a;
                        color: #e2e8f0;
                        font-size: 1rem;
                        margin-bottom: 1rem;
                    "
                >
                
                <label style="display: block; color: #e2e8f0; margin-bottom: 0.5rem; font-weight: 500;">
                    Confirm New Password:
                </label>
                <input 
                    type="password" 
                    id="confirmPassword" 
                    placeholder="Confirm new password"
                    style="
                        width: 100%;
                        padding: 0.75rem;
                        border: 2px solid #374151;
                        border-radius: 8px;
                        background: #0a0a0a;
                        color: #e2e8f0;
                        font-size: 1rem;
                    "
                    onkeypress="if(event.key==='Enter') attemptPasswordChange()"
                >
            </div>
            
            <div style="display: flex; gap: 1rem; justify-content: center;">
                <button 
                    onclick="closePasswordModal()"
                    style="
                        padding: 0.75rem 1.5rem;
                        border: 1px solid #374151;
                        border-radius: 8px;
                        background: #374151;
                        color: #e2e8f0;
                        cursor: pointer;
                        font-weight: 500;
                    "
                >
                    Cancel
                </button>
                <button 
                    onclick="attemptPasswordChange()"
                    style="
                        padding: 0.75rem 1.5rem;
                        border: 1px solid #10b981;
                        border-radius: 8px;
                        background: #10b981;
                        color: white;
                        cursor: pointer;
                        font-weight: 500;
                    "
                >
                    <i class="bi bi-check-circle"></i> Update Password
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    modal.id = 'passwordModal';
    
    // Focus on current password input
    setTimeout(() => {
        document.getElementById('currentPassword').focus();
    }, 100);
}

function closePasswordModal() {
    const modal = document.getElementById('passwordModal');
    if (modal) {
        modal.remove();
    }
}

async function attemptPasswordChange() {
    const currentPasswordInput = document.getElementById('currentPassword');
    const newPasswordInput = document.getElementById('newPassword');
    const confirmPasswordInput = document.getElementById('confirmPassword');
    
    const currentPassword = currentPasswordInput.value;
    const newPassword = newPasswordInput.value;
    const confirmPassword = confirmPasswordInput.value;
    
    try {
        // Validate current password using Firebase
        const currentPasswordHash = await hashPassword(currentPassword);
        const storedPasswordHash = await getStoredPasswordHash();
        
        if (currentPasswordHash !== storedPasswordHash) {
            showPasswordError(currentPasswordInput, 'Current password is incorrect');
            return;
        }
        
        // Validate new password
        if (newPassword.length < 8) {
            showPasswordError(newPasswordInput, 'New password must be at least 8 characters');
            return;
        }
        
        if (newPassword !== confirmPassword) {
            showPasswordError(confirmPasswordInput, 'Passwords do not match');
            return;
        }
        
        // Save new password to Firebase
        const newPasswordHash = await hashPassword(newPassword);
        const success = await setStoredPasswordHash(newPasswordHash);
        
        if (success) {
            closePasswordModal();
            showSuccessMessage('🔑 Password updated successfully! All other active sessions have been logged out for security.');
        } else {
            showPasswordError(newPasswordInput, 'Failed to update password in database');
        }
    } catch (error) {
        console.error('Password change error:', error);
        showPasswordError(newPasswordInput, 'System error occurred');
    }
}

function showPasswordError(inputElement, message) {
    inputElement.style.borderColor = '#dc2626';
    inputElement.style.animation = 'shake 0.5s ease-in-out';
    inputElement.value = '';
    inputElement.placeholder = `❌ ${message}`;
    
    setTimeout(() => {
        inputElement.style.borderColor = '#374151';
        inputElement.style.animation = '';
        inputElement.placeholder = inputElement.id === 'currentPassword' ? 'Enter current password' : 
                                   inputElement.id === 'newPassword' ? 'Enter new password' : 'Confirm new password';
    }, 3000);
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
        showLoading(true, 'Connecting to BOP25 Server...');
        console.log('🔥 Loading data from Firebase...');
        console.log('🔗 Firebase ready status:', isFirebaseReady);
        console.log('💾 Database reference:', database ? 'Available' : 'Not available');
        
        // Load main attendance data
        showLoading(true, 'Loading attendance data from BOP25 Server...');
        console.log('📊 Checking for attendance data...');
        const attendanceSnapshot = await database.ref('fingerprintData/attendance').once('value');
        console.log('📋 Attendance snapshot exists:', attendanceSnapshot.exists());
        
        if (attendanceSnapshot.exists()) {
            const data = attendanceSnapshot.val();
            console.log('📁 Raw data from Firebase:', data);
            extractedData = data.extractedData || [];
            console.log('👥 Extracted employees count:', extractedData.length);
            
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
        showLoading(true, 'Loading months data...');
        console.log('📅 Checking for months data...');
        const monthsSnapshot = await database.ref('fingerprintData/months').once('value');
        console.log('📅 Months snapshot exists:', monthsSnapshot.exists());
        if (monthsSnapshot.exists()) {
            const data = monthsSnapshot.val();
            allMonths = data.allMonths || [];
            console.log('📅 Loaded months:', allMonths.length);
        }
        
        // Load holidays data
        showLoading(true, 'Loading holidays data...');
        console.log('🏖️ Checking for holidays data...');
        const holidaysSnapshot = await database.ref('fingerprintData/holidays').once('value');
        console.log('🏖️ Holidays snapshot exists:', holidaysSnapshot.exists());
        if (holidaysSnapshot.exists()) {
            const data = holidaysSnapshot.val();
            holidays = new Set(data.holidays || []);
            console.log('🏖️ Loaded holidays:', holidays.size);
        }
        
        // Display data if available
        showLoading(true, 'Preparing display...');
        console.log('🎯 Final extractedData length:', extractedData.length);
        if (extractedData.length > 0) {
            console.log('✅ Displaying data...');
            displayDataWithFilters();
            showDataStatus();
            showLoading(true, 'Data loaded successfully!');
            setTimeout(() => showLoading(false), 1000);
        } else {
            console.log('❌ No data to display');
            showLoading(true, 'No existing data found');
            setTimeout(() => showLoading(false), 2000);
        }
        
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
        showLoading(true, 'Saving data to BOP25 Server...');
        console.log('Saving data to Firebase...');
        
        // Prepare data for Firebase (convert Date objects to ISO strings)
        showLoading(true, 'Preparing data for BOP25 Server...');
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
        
        showLoading(true, 'Uploading to BOP25 Server...');
        await database.ref().update(updates);
        console.log('Data saved to Firebase successfully');
        
        // Also save to localStorage as backup
        showLoading(true, 'Creating local backup...');
        saveDataToLocalStorage();
        
        showLoading(true, 'Save completed!');
        setTimeout(() => showLoading(false), 1000);
        
    } catch (error) {
        console.error('Error saving data to Firebase:', error);
        showLoading(false);
        showError('Warning: Could not save data to BOP25 Server. Saved locally instead.');
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
    showLoading(true, 'Reading Excel file...');
    showFileName(file.name);
    hideError();
    
    // Read and process the file
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            showLoading(true, 'Processing Excel data...');
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
    showLoading(true, 'Extracting employee data...');
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
    
    showLoading(true, 'Merging with existing data...');
    // If no existing data, use new data as is
    if (extractedData.length === 0) {
        extractedData = newData;
    } else {
        // Merge with existing data
        mergeEmployeeData(newData);
    }
    
    showLoading(true, 'Updating months list...');
    // Update months list
    updateMonthsList();
    
    showLoading(true, 'Saving data...');
    // Save to storage
    await saveDataFromStorage();
    
    showLoading(true, 'Displaying results...');
    // Display updated data
    displayDataWithFilters();
    showDataStatus();
    
    showLoading(true, 'File processed successfully!');
    setTimeout(() => showLoading(false), 1500);
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
function showLoading(show, message = 'Processing...') {
    const loading = document.getElementById('loading');
    if (loading) {
        const loadingText = loading.querySelector('span');
        if (loadingText && message) {
            loadingText.textContent = message;
        }
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
    
    // Sort by attendance (present days) - best to lowest
    filteredData = filteredData.sort((a, b) => {
        const aPresentDays = getPresentDaysCount(a.attendanceData);
        const bPresentDays = getPresentDaysCount(b.attendanceData);
        return bPresentDays - aPresentDays; // Descending order (best first)
    });
    
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
                                    <th>Ranking</th>
                                    <th>Prefect Name</th>
                                    <th class="mobile-hide">Morning Entrance Times</th>
                                    <th>Working Days</th>
                                    <th>Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${filteredData.map((record, index) => `
                                    <tr class="table-row-fade" style="animation-delay: ${400 + (index * 50)}ms">
                                        <td><span class="ranking-badge rank-${index + 1}">#${index + 1}</span></td>
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
                                            <div class="days-summary">
                                                <span class="days-badge">${getPresentDaysCount(record.attendanceData)} Present</span>
                                                <span class="absent-days">${getWorkingDaysCount(record.attendanceData) - getPresentDaysCount(record.attendanceData)} Absent</span>
                                            </div>
                                            <div class="days-detail">
                                                Working: ${getWorkingDaysCount(record.attendanceData)} | 
                                                Holidays: ${record.attendanceData.filter(r => {
                                                    const dateString = r.fullDate ? r.fullDate.toISOString().split('T')[0] : null;
                                                    return dateString && isHoliday(dateString);
                                                }).length}
                                            </div>
                                        </td>
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
            
            .ranking-badge {
                padding: 0.25rem 0.5rem;
                border-radius: 6px;
                font-weight: 600;
                font-size: 0.75rem;
                display: inline-flex;
                align-items: center;
                justify-content: center;
                min-width: 2rem;
            }
            
            .rank-1 {
                background: linear-gradient(135deg, #ffd700, #ffb700);
                color: #1a1a1a;
                box-shadow: 0 2px 8px rgba(255, 215, 0, 0.3);
            }
            
            .rank-2 {
                background: linear-gradient(135deg, #c0c0c0, #a0a0a0);
                color: #1a1a1a;
                box-shadow: 0 2px 8px rgba(192, 192, 192, 0.3);
            }
            
            .rank-3 {
                background: linear-gradient(135deg, #cd7f32, #b8860b);
                color: #fff;
                box-shadow: 0 2px 8px rgba(205, 127, 50, 0.3);
            }
            
            .ranking-badge:not(.rank-1):not(.rank-2):not(.rank-3) {
                background: #374151;
                color: #10b981;
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
                padding: 4px 8px;
                border-radius: 8px;
                font-weight: 600;
                font-size: 0.75rem;
                border: 1px solid #047857;
                display: inline-block;
            }
            
            .days-detail {
                font-size: 0.75rem;
                color: #9ca3af;
                margin-top: 0.25rem;
            }
            
            .days-summary {
                display: flex;
                gap: 0.5rem;
                align-items: center;
                margin-bottom: 0.25rem;
                flex-wrap: wrap;
            }
            
            .absent-days {
                background-color: #dc2626;
                color: white;
                padding: 4px 8px;
                border-radius: 8px;
                font-weight: 600;
                font-size: 0.75rem;
                display: inline-block;
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
        // Remove any existing event listeners first by cloning and replacing elements
        const oldButtons = document.querySelectorAll('.btn-view');
        oldButtons.forEach(button => {
            const newButton = button.cloneNode(true);
            button.parentNode.replaceChild(newButton, button);
        });
        
        // Now add fresh event listeners to the new buttons
        document.querySelectorAll('.btn-view').forEach(button => {
            button.addEventListener('click', function(e) {
                // Prevent any propagation or default behavior
                e.preventDefault();
                e.stopPropagation();
                
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
            
            <div class="summary-card absent-card">
                <div class="summary-icon">
                    <i class="bi bi-person-x"></i>
                </div>
                <div class="summary-content">
                    <h4 class="summary-number">${attendanceToShow.filter(r => {
                        const isWeekend = r.dayOfWeek === 'SUN' || r.dayOfWeek === 'SAT';
                        const dateString = r.fullDate ? r.fullDate.toISOString().split('T')[0] : null;
                        const isHolidayDate = dateString && isHoliday(dateString);
                        return !r.morning.in && !isWeekend && !isHolidayDate;
                    }).length}</h4>
                    <p class="summary-label">Absent Days</p>
                    <small class="summary-note">No morning entrance</small>
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
            
            .absent-card {
                border-left: 4px solid #dc2626;
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
            
            .absent-card .summary-icon {
                background: rgba(220, 38, 38, 0.2);
                color: #dc2626;
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
    
    // Create a summary dashboard worksheet first
    const summaryData = [];
    summaryData.push(['📊 BOP25 PREFECTS ATTENDANCE REPORT', '', '', '', '', '', '', '']);
    summaryData.push(['Generated:', new Date().toLocaleDateString() + ' ' + new Date().toLocaleTimeString(), '', '', '', '', '', '']);
    summaryData.push(['Period:', selectedMonth === 'all' ? 'All Available Months' : new Date(selectedMonth + '-01').toLocaleDateString('en-US', { month: 'long', year: 'numeric' }), '', '', '', '', '', '']);
    summaryData.push(['Total Prefects:', extractedData.length, '', '', '', '', '', '']);
    summaryData.push(['', '', '', '', '', '', '', '']);
    
    // Add month-wise summary statistics
    summaryData.push(['📅 MONTHLY OVERVIEW', '', '', '', '', '', '', '']);
    summaryData.push(['Month', 'Total Prefects', 'Avg Attendance %', 'Top Performer', 'Most Improved', 'Needs Attention', '', '']);
    
    monthsToProcess.forEach(monthKey => {
        const [year, month] = monthKey.split('-');
        const monthName = new Date(year, month - 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
        
        // Calculate month statistics
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
        
        // Calculate attendance percentages for ranking
        const employeeStats = monthData.map(employee => {
            const workingDays = getWorkingDaysCount(employee.attendanceData);
            const presentDays = getPresentDaysCount(employee.attendanceData);
            const attendanceRate = workingDays > 0 ? Math.round((presentDays / workingDays) * 100) : 0;
            
            return {
                name: employee.name,
                attendanceRate,
                presentDays,
                workingDays,
                lateCount: employee.attendanceData.filter(r => {
                    const isWeekend = r.dayOfWeek === 'SUN' || r.dayOfWeek === 'SAT';
                    const dateString = r.fullDate ? r.fullDate.toISOString().split('T')[0] : null;
                    const isHolidayDate = dateString && isHoliday(dateString);
                    return r.morning.in && r.morning.in > '06:45' && !isWeekend && !isHolidayDate;
                }).length
            };
        });
        
        const avgAttendance = Math.round(employeeStats.reduce((sum, emp) => sum + emp.attendanceRate, 0) / employeeStats.length);
        const topPerformer = employeeStats.reduce((top, emp) => emp.attendanceRate > top.attendanceRate ? emp : top);
        const needsAttention = employeeStats.filter(emp => emp.attendanceRate < 75).length > 0 
            ? employeeStats.reduce((worst, emp) => emp.attendanceRate < worst.attendanceRate ? emp : worst).name
            : 'None';
        
        summaryData.push([
            monthName,
            monthData.length,
            avgAttendance + '%',
            `${topPerformer.name} (${topPerformer.attendanceRate}%)`,
            'TBD', // Could add logic for improvement tracking
            needsAttention,
            '', ''
        ]);
    });
    
    summaryData.push(['', '', '', '', '', '', '', '']);
    summaryData.push(['🏆 PERFORMANCE CATEGORIES', '', '', '', '', '', '', '']);
    summaryData.push(['Excellent (95-100%)', 'Good (85-94%)', 'Average (75-84%)', 'Needs Improvement (<75%)', '', '', '', '']);
    
    // Create summary worksheet with enhanced styling
    const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
    
    // Enhanced styling for summary sheet
    const summaryRange = XLSX.utils.decode_range(summaryWs['!ref']);
    
    // Title styling
    const titleCell = XLSX.utils.encode_cell({r: 0, c: 0});
    summaryWs[titleCell].s = {
        fill: { fgColor: { rgb: "1F4E79" } },
        font: { color: { rgb: "FFFFFF" }, bold: true, sz: 16 },
        alignment: { horizontal: "center", vertical: "center" }
    };
    
    // Merge title cell
    if (!summaryWs['!merges']) summaryWs['!merges'] = [];
    summaryWs['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } });
    
    // Info cells styling
    for (let r = 1; r <= 4; r++) {
        const labelCell = XLSX.utils.encode_cell({r: r, c: 0});
        const valueCell = XLSX.utils.encode_cell({r: r, c: 1});
        if (summaryWs[labelCell]) {
            summaryWs[labelCell].s = {
                font: { bold: true, sz: 11 },
                fill: { fgColor: { rgb: "E7E6E6" } }
            };
        }
        if (summaryWs[valueCell]) {
            summaryWs[valueCell].s = {
                font: { sz: 11 },
                fill: { fgColor: { rgb: "F8F9FA" } }
            };
        }
    }
    
    // Column headers styling
    const headerRow = 7;
    for (let c = 0; c <= 7; c++) {
        const headerCell = XLSX.utils.encode_cell({r: headerRow, c: c});
        if (summaryWs[headerCell]) {
            summaryWs[headerCell].s = {
                fill: { fgColor: { rgb: "366092" } },
                font: { color: { rgb: "FFFFFF" }, bold: true, sz: 11 },
                alignment: { horizontal: "center", vertical: "center" },
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                }
            };
        }
    }
    
    // Set column widths for summary
    summaryWs['!cols'] = [
        { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 20 }, { wch: 10 }, { wch: 10 }
    ];
    
    // Add summary worksheet to workbook
    XLSX.utils.book_append_sheet(wb, summaryWs, '📊 Dashboard');
    
    // Create a worksheet for each month with enhanced formatting
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
        
        // Create worksheet data with enhanced headers
        const sheetData = [];
        
        // Header with month and year and metadata
        sheetData.push([`📅 ${monthName} - BOP25 Prefects Attendance`, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '📊 Performance Summary']);
        sheetData.push([`Generated: ${new Date().toLocaleDateString()} | Total Prefects: ${monthData.length} | Working Days: ${getUniqueWorkingDaysInMonth(monthData, analysisYear, analysisMonth)}`, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        sheetData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        // Enhanced column headers with better descriptions
        const dayHeaders = Array.from({length: daysInMonth}, (_, i) => {
            const date = new Date(analysisYear, analysisMonth - 1, i + 1);
            const dayName = date.toLocaleDateString('en-US', { weekday: 'short' });
            return `${i + 1}\n${dayName}`;
        });
        
        sheetData.push(['S.No', 'Prefect Name', 'ID', ...dayHeaders, '✅ Present', '⏰ Late', '🎯 On Time', '📅 Work Days', '📊 Attend %', '🏆 Grade', '📝 Notes']);
        
        // Add data for each employee with performance grading
        monthData.forEach((employee, index) => {
            const row = [index + 1, employee.name, employee.employeeId || 'N/A'];
            
            // Create attendance map for quick lookup
            const attendanceMap = {};
            employee.attendanceData.forEach(record => {
                if (record.fullDate) {
                    const day = record.fullDate.getDate();
                    attendanceMap[day] = record;
                }
            });
            
            // Fill daily attendance with enhanced formatting
            let presentCount = 0;
            let lateCount = 0;
            let onTimeCount = 0;
            let workingDaysCount = 0;
            let consecutiveAbsent = 0;
            let maxConsecutiveAbsent = 0;
            
            for (let day = 1; day <= daysInMonth; day++) {
                const record = attendanceMap[day];
                const checkDate = new Date(analysisYear, analysisMonth - 1, day);
                const dayOfWeek = checkDate.toLocaleDateString('en-US', { weekday: 'short' }).toUpperCase();
                const dateString = checkDate.toISOString().split('T')[0];
                const isHolidayDate = isHoliday(dateString);
                const isWeekend = dayOfWeek === 'SAT' || dayOfWeek === 'SUN';
                
                if (isHolidayDate) {
                    row.push('🏖️ H'); // Holiday with emoji
                    consecutiveAbsent = 0;
                } else if (isWeekend) {
                    row.push('🏠 W'); // Weekend with emoji
                    consecutiveAbsent = 0;
                } else {
                    workingDaysCount++;
                    if (record && record.morning.in) {
                        const [hours, minutes] = record.morning.in.split(':').map(Number);
                        const isLate = hours > 6 || (hours === 6 && minutes > 45);
                        
                        if (isLate) {
                            row.push(`⏰ L\n${record.morning.in}`); // Late with emoji and time
                            lateCount++;
                        } else {
                            row.push(`✅ P\n${record.morning.in}`); // Present on time with emoji and time
                            onTimeCount++;
                        }
                        presentCount++;
                        consecutiveAbsent = 0;
                    } else {
                        row.push('❌ A'); // Absent with emoji
                        consecutiveAbsent++;
                        maxConsecutiveAbsent = Math.max(maxConsecutiveAbsent, consecutiveAbsent);
                    }
                }
            }
            
            // Enhanced summary columns with performance grading
            const attendanceRate = workingDaysCount > 0 ? Math.round((presentCount / workingDaysCount) * 100) : 0;
            let grade = '';
            let notes = '';
            
            if (attendanceRate >= 95) {
                grade = '🏆 Excellent';
            } else if (attendanceRate >= 85) {
                grade = '🥈 Good';
            } else if (attendanceRate >= 75) {
                grade = '🥉 Average';
            } else {
                grade = '⚠️ Needs Improvement';
            }
            
            // Generate notes
            if (maxConsecutiveAbsent >= 3) {
                notes += `Max ${maxConsecutiveAbsent} consecutive absences. `;
            }
            if (lateCount > workingDaysCount * 0.3) {
                notes += 'Frequent lateness. ';
            }
            if (attendanceRate >= 95 && lateCount <= 2) {
                notes += 'Excellent performance! ';
            }
            if (notes === '') {
                notes = 'Good attendance pattern.';
            }
            
            row.push(presentCount, lateCount, onTimeCount, workingDaysCount, `${attendanceRate}%`, grade, notes.trim());
            
            sheetData.push(row);
        });
        
        // Add enhanced statistics and insights
        sheetData.push([]);
        sheetData.push(['📈 MONTHLY INSIGHTS & STATISTICS', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        const totalWorkingDays = getUniqueWorkingDaysInMonth(monthData, analysisYear, analysisMonth);
        const avgAttendanceRate = monthData.length > 0 ? 
            Math.round(monthData.reduce((sum, emp) => {
                const workingDays = getWorkingDaysCount(emp.attendanceData);
                const presentDays = getPresentDaysCount(emp.attendanceData);
                return sum + (workingDays > 0 ? (presentDays / workingDays) * 100 : 0);
            }, 0) / monthData.length) : 0;
        
        const excellentPerformers = monthData.filter(emp => {
            const workingDays = getWorkingDaysCount(emp.attendanceData);
            const presentDays = getPresentDaysCount(emp.attendanceData);
            return workingDays > 0 && ((presentDays / workingDays) * 100) >= 95;
        }).length;
        
        sheetData.push(['📊 Overall Statistics:', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        sheetData.push([`• Working Days in Month: ${totalWorkingDays}`, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        sheetData.push([`• Average Attendance Rate: ${avgAttendanceRate}%`, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        sheetData.push([`• Excellent Performers (95%+): ${excellentPerformers}/${monthData.length} prefects`, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        sheetData.push([]);
        sheetData.push(['📋 LEGEND & STATUS CODES:', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        sheetData.push(['✅ P = Present (On Time) | ⏰ L = Late (After 6:45 AM) | ❌ A = Absent | 🏖️ H = Holiday | 🏠 W = Weekend', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        sheetData.push(['🏆 Excellent: 95-100% | 🥈 Good: 85-94% | 🥉 Average: 75-84% | ⚠️ Needs Improvement: <75%', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        
        // Enhanced column widths
        const colWidths = [
            { wch: 6 },  // S.No
            { wch: 25 }, // Name (wider)
            { wch: 12 }, // ID
            ...Array.from({length: daysInMonth}, () => ({ wch: 10 })), // Day columns (wider for emoji + time)
            { wch: 10 }, // Present
            { wch: 8 },  // Late
            { wch: 10 }, // On Time
            { wch: 10 }, // Work Days
            { wch: 10 }, // Attend %
            { wch: 18 }, // Grade (wider for emojis)
            { wch: 30 }  // Notes (much wider)
        ];
        ws['!cols'] = colWidths;
        
        // Enhanced cell styling and colors
        const range = XLSX.utils.decode_range(ws['!ref']);
        
        // Style title row (enhanced)
        for (let C = range.s.c; C <= range.e.c; ++C) {
            // First row (month header with enhanced styling)
            const headerCell = XLSX.utils.encode_cell({r: 0, c: C});
            if (!ws[headerCell]) ws[headerCell] = {t: 's', v: ''};
            ws[headerCell].s = {
                fill: { fgColor: { rgb: "1F4E79" } },
                font: { color: { rgb: "FFFFFF" }, bold: true, sz: 14 },
                alignment: { horizontal: "center", vertical: "center" }
            };
            
            // Second row (metadata)
            const metaCell = XLSX.utils.encode_cell({r: 1, c: C});
            if (!ws[metaCell]) ws[metaCell] = {t: 's', v: ''};
            ws[metaCell].s = {
                fill: { fgColor: { rgb: "E7E6E6" } },
                font: { italic: true, sz: 10 },
                alignment: { horizontal: "center", vertical: "center" }
            };
            
            // Column headers (4th row)
            const colHeaderCell = XLSX.utils.encode_cell({r: 3, c: C});
            if (!ws[colHeaderCell]) ws[colHeaderCell] = {t: 's', v: ''};
            ws[colHeaderCell].s = {
                fill: { fgColor: { rgb: "366092" } },
                font: { color: { rgb: "FFFFFF" }, bold: true, sz: 10 },
                alignment: { horizontal: "center", vertical: "center", wrapText: true },
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                }
            };
        }
        
        // Enhanced data row styling with performance-based colors
        const dataStartRow = 4;
        const dataEndRow = 4 + monthData.length - 1;
        
        for (let R = dataStartRow; R <= dataEndRow; ++R) {
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
                
                // Enhanced color coding based on cell content and position
                if (C >= 3 && C <= daysInMonth + 2) { // Day columns (adjusted for ID column)
                    if (typeof cellValue === 'string') {
                        if (cellValue.includes('✅ P')) {
                            // Present (Enhanced Green)
                            cellStyle.fill = { fgColor: { rgb: "C6EFCE" } };
                            cellStyle.font = { color: { rgb: "006100" }, bold: true, sz: 9 };
                        } else if (cellValue.includes('⏰ L')) {
                            // Late (Enhanced Orange)
                            cellStyle.fill = { fgColor: { rgb: "FFEB9C" } };
                            cellStyle.font = { color: { rgb: "9C5700" }, bold: true, sz: 9 };
                        } else if (cellValue.includes('❌ A')) {
                            // Absent (Enhanced Red)
                            cellStyle.fill = { fgColor: { rgb: "FFC7CE" } };
                            cellStyle.font = { color: { rgb: "9C0006" }, bold: true, sz: 9 };
                        } else if (cellValue.includes('🏖️ H')) {
                            // Holiday (Enhanced Blue)
                            cellStyle.fill = { fgColor: { rgb: "BDD7EE" } };
                            cellStyle.font = { color: { rgb: "0070C0" }, bold: true, sz: 9 };
                        } else if (cellValue.includes('🏠 W')) {
                            // Weekend (Enhanced Gray)
                            cellStyle.fill = { fgColor: { rgb: "E2E2E2" } };
                            cellStyle.font = { color: { rgb: "7C7C7C" }, bold: true, sz: 9 };
                        }
                    }
                } else if (C === 0 || C === 1 || C === 2) {
                    // S.No, Name, and ID columns
                    cellStyle.fill = { fgColor: { rgb: "F2F2F2" } };
                    cellStyle.font = { bold: true, sz: 10 };
                    if (C === 1) { // Name column
                        cellStyle.alignment.horizontal = "left";
                        cellStyle.font.sz = 11;
                    }
                } else if (C === range.e.c - 1) { // Grade column
                    // Color based on grade
                    if (typeof cellValue === 'string') {
                        if (cellValue.includes('🏆 Excellent')) {
                            cellStyle.fill = { fgColor: { rgb: "C6EFCE" } };
                            cellStyle.font = { color: { rgb: "006100" }, bold: true };
                        } else if (cellValue.includes('🥈 Good')) {
                            cellStyle.fill = { fgColor: { rgb: "FFEB9C" } };
                            cellStyle.font = { color: { rgb: "9C5700" }, bold: true };
                        } else if (cellValue.includes('🥉 Average')) {
                            cellStyle.fill = { fgColor: { rgb: "FCE4D6" } };
                            cellStyle.font = { color: { rgb: "C65911" }, bold: true };
                        } else if (cellValue.includes('⚠️ Needs')) {
                            cellStyle.fill = { fgColor: { rgb: "FFC7CE" } };
                            cellStyle.font = { color: { rgb: "9C0006" }, bold: true };
                        }
                    }
                } else if (C >= daysInMonth + 3) {
                    // Summary columns
                    cellStyle.fill = { fgColor: { rgb: "E7E6E6" } };
                    cellStyle.font.bold = true;
                    
                    // Special formatting for attendance percentage
                    if (C === range.e.c - 3) { // Attendance % column
                        if (typeof cellValue === 'string' && cellValue.includes('%')) {
                            const percentage = parseInt(cellValue);
                            if (percentage >= 95) {
                                cellStyle.fill = { fgColor: { rgb: "C6EFCE" } };
                                cellStyle.font.color = { rgb: "006100" };
                            } else if (percentage >= 85) {
                                cellStyle.fill = { fgColor: { rgb: "FFEB9C" } };
                                cellStyle.font.color = { rgb: "9C5700" };
                            } else if (percentage >= 75) {
                                cellStyle.fill = { fgColor: { rgb: "FCE4D6" } };
                                cellStyle.font.color = { rgb: "C65911" };
                            } else {
                                cellStyle.fill = { fgColor: { rgb: "FFC7CE" } };
                                cellStyle.font.color = { rgb: "9C0006" };
                            }
                        }
                    }
                }
                
                ws[cellAddress].s = cellStyle;
            }
        }
        
        // Style insights and legend sections
        const insightsStartRow = dataEndRow + 2;
        for (let R = insightsStartRow; R < sheetData.length; ++R) {
            for (let C = 0; C <= 10; ++C) {
                const cellAddress = XLSX.utils.encode_cell({r: R, c: C});
                if (ws[cellAddress]) {
                    let style = {
                        alignment: { horizontal: "left", vertical: "center" }
                    };
                    
                    if (ws[cellAddress].v && ws[cellAddress].v.includes('📈')) {
                        // Insights header
                        style.font = { bold: true, sz: 12, color: { rgb: "1F4E79" } };
                        style.fill = { fgColor: { rgb: "E7E6E6" } };
                    } else if (ws[cellAddress].v && ws[cellAddress].v.includes('📊')) {
                        // Statistics header
                        style.font = { bold: true, sz: 11, color: { rgb: "366092" } };
                    } else if (ws[cellAddress].v && ws[cellAddress].v.includes('📋')) {
                        // Legend header
                        style.font = { bold: true, sz: 11, color: { rgb: "7030A0" } };
                        style.fill = { fgColor: { rgb: "E7E6E6" } };
                    } else {
                        // Regular text
                        style.font = { sz: 10 };
                    }
                    
                    ws[cellAddress].s = style;
                }
            }
        }
        
        // Set enhanced row heights for better readability
        if (!ws['!rows']) ws['!rows'] = [];
        // Title and header rows
        ws['!rows'][0] = { hpx: 40 }; // Title row
        ws['!rows'][1] = { hpx: 25 }; // Metadata row
        ws['!rows'][3] = { hpx: 35 }; // Column headers
        
        // Data rows
        for (let i = dataStartRow; i <= dataEndRow; i++) {
            ws['!rows'][i] = { hpx: 35 }; // Increase row height for emoji and wrapped text
        }
        
        // Merge cells for better presentation
        if (!ws['!merges']) ws['!merges'] = [];
        // Merge title row
        ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: daysInMonth + 6 } });
        // Merge metadata row
        ws['!merges'].push({ s: { r: 1, c: 0 }, e: { r: 1, c: daysInMonth + 6 } });
        
        // Add worksheet to workbook with enhanced sheet name
        const sheetName = new Date(analysisYear, analysisMonth - 1).toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '_');
        XLSX.utils.book_append_sheet(wb, ws, `📅 ${sheetName}`);
    });
    
    // Generate enhanced filename with timestamp
    const timestamp = new Date().toISOString().split('T')[0];
    const timeStr = new Date().toLocaleTimeString('en-US', { hour12: false }).replace(/:/g, '-');
    const filename = selectedMonth !== 'all' 
        ? `BOP25_Attendance_Report_${selectedMonth}_${timestamp}_${timeStr}.xlsx`
        : `BOP25_Attendance_Report_All_Months_${timestamp}_${timeStr}.xlsx`;
    
    // Save file
    XLSX.writeFile(wb, filename);
    
    // Show success message
    showSuccessMessage(`📊 Enhanced attendance report downloaded successfully! File: ${filename}`);
}

// Helper function to get unique working days in a month
function getUniqueWorkingDaysInMonth(monthData, year, month) {
    const daysInMonth = new Date(year, month, 0).getDate();
    let workingDays = 0;
    
    for (let day = 1; day <= daysInMonth; day++) {
        const checkDate = new Date(year, month - 1, day);
        const dayOfWeek = checkDate.toLocaleDateString('en-US', { weekday: 'short' }).toUpperCase();
        const dateString = checkDate.toISOString().split('T')[0];
        const isHolidayDate = isHoliday(dateString);
        const isWeekend = dayOfWeek === 'SAT' || dayOfWeek === 'SUN';
        
        if (!isWeekend && !isHolidayDate) {
            workingDays++;
        }
    }
    
    return workingDays;
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
    
    // Calculate comprehensive statistics
    const workingDays = getWorkingDaysCount(attendanceToShow);
    const presentDays = getPresentDaysCount(attendanceToShow);
    const lateDays = attendanceToShow.filter(r => {
        const isWeekend = r.dayOfWeek === 'SUN' || r.dayOfWeek === 'SAT';
        const dateString = r.fullDate ? r.fullDate.toISOString().split('T')[0] : null;
        const isHolidayDate = dateString && isHoliday(dateString);
        return r.morning.in && r.morning.in > '06:45' && !isWeekend && !isHolidayDate;
    }).length;
    const absentDays = workingDays - presentDays;
    const attendanceRate = workingDays > 0 ? Math.round((presentDays / workingDays) * 100) : 0;
    const punctualityRate = presentDays > 0 ? Math.round(((presentDays - lateDays) / presentDays) * 100) : 0;
    
    // Enhanced CSV content with comprehensive analysis
    const csvRows = [
        ['🎓 BOP25 PREFECT INDIVIDUAL ATTENDANCE REPORT'],
        [''],
        ['👤 PREFECT INFORMATION'],
        ['Name:', employee.name],
        ['Prefect ID:', employee.employeeId || 'N/A'],
        ['Report Period:', selectedMonth === 'all' ? 'All Available Months' : new Date(selectedMonth + '-01').toLocaleDateString('en-US', { month: 'long', year: 'numeric' })],
        ['Generated On:', new Date().toLocaleDateString() + ' at ' + new Date().toLocaleTimeString()],
        [''],
        ['📊 PERFORMANCE SUMMARY'],
        ['Total Working Days:', workingDays],
        ['Days Present:', presentDays],
        ['Days Late:', lateDays],
        ['Days Absent:', absentDays],
        ['Overall Attendance Rate:', attendanceRate + '%'],
        ['Punctuality Rate:', punctualityRate + '%'],
        [''],
        ['🏆 PERFORMANCE GRADE'],
        ['Grade:', attendanceRate >= 95 ? 'Excellent (🏆)' : attendanceRate >= 85 ? 'Good (🥈)' : attendanceRate >= 75 ? 'Average (🥉)' : 'Needs Improvement (⚠️)'],
        [''],
        ['📈 DETAILED ATTENDANCE ANALYSIS'],
        [''],
        ['📅 Date', '🗓️ Day', '⏰ Morning Entry', '📊 Status', '✅ On Time?', '📝 Notes'],
        ['']
    ];
    
    // Sort attendance data by date for better readability
    const sortedAttendance = attendanceToShow.sort((a, b) => {
        if (a.fullDate && b.fullDate) {
            return a.fullDate.getTime() - b.fullDate.getTime();
        }
        return 0;
    });
    
    // Add detailed attendance records with enhanced analysis
    sortedAttendance.forEach(record => {
        const isWeekend = record.dayOfWeek === 'SUN' || record.dayOfWeek === 'SAT';
        const dateString = record.fullDate ? record.fullDate.toISOString().split('T')[0] : null;
        const isHolidayDate = dateString && isHoliday(dateString);
        const isLate = record.morning.in && record.morning.in > '06:45';
        
        let status = '';
        let onTime = '';
        let notes = '';
        
        if (isHolidayDate) {
            status = '🏖️ HOLIDAY';
            onTime = 'N/A';
            notes = 'Official holiday - no attendance required';
        } else if (isWeekend) {
            status = '🏠 WEEKEND';
            onTime = 'N/A';
            notes = 'Weekend - no attendance required';
        } else if (record.morning.in) {
            if (isLate) {
                status = '⏰ LATE';
                onTime = '❌ No';
                const [hours, minutes] = record.morning.in.split(':').map(Number);
                const lateMinutes = (hours - 6) * 60 + (minutes - 45);
                notes = `Late by ${lateMinutes} minutes`;
            } else {
                status = '✅ ON TIME';
                onTime = '✅ Yes';
                notes = 'Arrived on time - excellent!';
            }
        } else {
            status = '❌ ABSENT';
            onTime = '❌ No';
            notes = 'No morning entrance recorded';
        }
        
        csvRows.push([
            record.date,
            record.dayOfWeek,
            record.morning.in || 'No Entry',
            status,
            onTime,
            notes
        ]);
    });
    
    // Add insights and recommendations
    csvRows.push(['']);
    csvRows.push(['💡 INSIGHTS & RECOMMENDATIONS']);
    csvRows.push(['']);
    
    if (attendanceRate >= 95) {
        csvRows.push(['🏆 Excellent Performance!', 'Outstanding attendance record. Keep up the excellent work!']);
    } else if (attendanceRate >= 85) {
        csvRows.push(['🥈 Good Performance', 'Solid attendance record with room for minor improvements.']);
    } else if (attendanceRate >= 75) {
        csvRows.push(['🥉 Average Performance', 'Attendance is acceptable but could be improved for better results.']);
    } else {
        csvRows.push(['⚠️ Needs Improvement', 'Attendance requires immediate attention and improvement.']);
    }
    
    if (lateDays > 0) {
        csvRows.push(['Punctuality Note:', `Consider improving arrival time. Late on ${lateDays} out of ${presentDays} present days.`]);
    }
    
    if (absentDays > workingDays * 0.2) {
        csvRows.push(['Attendance Alert:', 'High absence rate detected. Please discuss with board administration.']);
    }
    
    // Add monthly breakdown if viewing all months
    if (selectedMonth === 'all' && allMonths.length > 1) {
        csvRows.push(['']);
        csvRows.push(['📅 MONTHLY BREAKDOWN']);
        csvRows.push(['Month', 'Working Days', 'Present', 'Absent', 'Late', 'Attendance %']);
        
        allMonths.forEach(month => {
            const [year, monthNum] = month.split('-');
            const monthData = employee.attendanceData.filter(record => 
                record.fullDate && record.fullDate.getFullYear() == year && 
                record.fullDate.getMonth() + 1 == monthNum
            );
            
            if (monthData.length > 0) {
                const monthWorkingDays = getWorkingDaysCount(monthData);
                const monthPresentDays = getPresentDaysCount(monthData);
                const monthLateDays = monthData.filter(r => {
                    const isWeekend = r.dayOfWeek === 'SUN' || r.dayOfWeek === 'SAT';
                    const dateString = r.fullDate ? r.fullDate.toISOString().split('T')[0] : null;
                    const isHolidayDate = dateString && isHoliday(dateString);
                    return r.morning.in && r.morning.in > '06:45' && !isWeekend && !isHolidayDate;
                }).length;
                const monthAbsentDays = monthWorkingDays - monthPresentDays;
                const monthAttendanceRate = monthWorkingDays > 0 ? Math.round((monthPresentDays / monthWorkingDays) * 100) : 0;
                
                const monthName = new Date(year, monthNum - 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
                csvRows.push([monthName, monthWorkingDays, monthPresentDays, monthAbsentDays, monthLateDays, monthAttendanceRate + '%']);
            }
        });
    }
    
    // Add footer
    csvRows.push(['']);
    csvRows.push(['📋 REPORT FOOTER']);
    csvRows.push(['Generated by:', 'BOP25 Prefects Attendance Management System']);
    csvRows.push(['Contact:', 'Board of Prefects 2025']);
    csvRows.push(['Note:', 'This report is generated automatically and contains official attendance data.']);
    
    const csvContent = csvRows.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
    const timestamp = new Date().toISOString().split('T')[0];
    const filename = `BOP25_${employee.name.replace(/\s+/g, '_')}_Detailed_Report_${selectedMonth === 'all' ? 'All_Months' : selectedMonth}_${timestamp}.csv`;
    
    downloadCSVFile(csvContent, filename);
    
    // Show success message
    showSuccessMessage(`📊 Detailed attendance report for ${employee.name} downloaded successfully!`);
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
                    <div class="modal-header border-0" style="background: #0a0a0a; border-bottom: 1px solid #2d2d2d;">
                        <h5 class="modal-title fw-light" style="color: #10b981;">
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
                    <div class="modal-header border-0" style="background: #0a0a0a; border-bottom: 1px solid #2d2d2d;">
                        <h5 class="modal-title fw-light" style="color: #e2e8f0;">
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
    
    // Update ranking for visible rows
    let visibleIndex = 1;
    tableRows.forEach(row => {
        if (row.style.display !== 'none') {
            const rankingBadge = row.querySelector('.ranking-badge');
            if (rankingBadge) {
                // Update ranking text
                rankingBadge.textContent = `#${visibleIndex}`;
                
                // Update ranking class for top 3
                rankingBadge.className = 'ranking-badge';
                if (visibleIndex === 1) {
                    rankingBadge.classList.add('rank-1');
                } else if (visibleIndex === 2) {
                    rankingBadge.classList.add('rank-2');
                } else if (visibleIndex === 3) {
                    rankingBadge.classList.add('rank-3');
                }
                
                visibleIndex++;
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
