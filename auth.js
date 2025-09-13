class AuthenticationSystem {
    constructor() {
        this.isAuthenticated = false;
        this.currentUser = null;
        this.sessionTimeout = 30 * 60 * 1000; // 30 minutes
        this.sessionTimer = null;
        
        this.init();
    }

    init() {
        // Check if user is already logged in
        this.checkExistingSession();
        
        // Initialize login form
        this.initLoginForm();
        
        // Add logout functionality to main app
        this.addLogoutButton();
    }

    checkExistingSession() {
        const token = localStorage.getItem('authToken');
        const user = localStorage.getItem('currentUser');
        const loginTime = localStorage.getItem('loginTime');
        
        if (token && user && loginTime) {
            const timeSinceLogin = Date.now() - parseInt(loginTime);
            
            // Check if session is still valid
            if (timeSinceLogin < this.sessionTimeout) {
                this.isAuthenticated = true;
                this.currentUser = JSON.parse(user);
                this.startSessionTimer();
                
                // Redirect to main app if on login page
                if (window.location.pathname.includes('login.html')) {
                    window.location.href = 'index.html';
                }
            } else {
                // Session expired
                this.logout();
            }
        } else if (window.location.pathname.includes('index.html')) {
            // Not logged in and trying to access main app
            window.location.href = 'login.html';
        }
    }

    initLoginForm() {
        const loginForm = document.getElementById('loginForm');
        const loginBtn = document.getElementById('loginBtn');
        const loading = document.getElementById('loading');

        if (loginForm) {
            loginForm.addEventListener('submit', (e) => {
                e.preventDefault();
                this.handleLogin();
            });
        }
    }

    async handleLogin() {
        const username = document.getElementById('username').value.trim();
        const password = document.getElementById('password').value;
        const loginBtn = document.getElementById('loginBtn');
        const loading = document.getElementById('loading');

        // Show loading state
        loginBtn.style.display = 'none';
        loading.style.display = 'block';

        try {
            // Simulate API call delay
            await new Promise(resolve => setTimeout(resolve, 1000));

            // Validate credentials
            if (this.validateCredentials(username, password)) {
                // Generate session token
                const token = this.generateToken();
                
                // Store session data (currentUser is already set in validateCredentials)
                localStorage.setItem('authToken', token);
                localStorage.setItem('currentUser', JSON.stringify(this.currentUser));
                localStorage.setItem('loginTime', Date.now().toString());
                
                this.isAuthenticated = true;
                this.startSessionTimer();
                
                // Show success message with access level
                const accessLevel = this.currentUser.accessLevel === 'admin' ? 'Admin (Read/Write)' : 'User (Read-Only)';
                this.showMessage(`Login successful! Access Level: ${accessLevel}. Redirecting...`, 'success');
                
                // Redirect to main app
                setTimeout(() => {
                    window.location.href = 'index.html';
                }, 1500);
                
            } else {
                this.showMessage('Invalid username or password', 'error');
            }
        } catch (error) {
            this.showMessage('Login failed. Please try again.', 'error');
        } finally {
            // Hide loading state
            loginBtn.style.display = 'block';
            loading.style.display = 'none';
        }
    }

    validateCredentials(username, password) {
        // Production credentials with access levels
        const validUsers = {
            'admin': {
                password: 'admin123',
                accessLevel: 'admin', // Read/Write access
                permissions: ['read', 'write', 'download', 'upload', 'delete']
            },
            'user': {
                password: 'user12345',
                accessLevel: 'user', // Read-only access
                permissions: ['read']
            }
        };

        const user = validUsers[username];
        if (user && user.password === password) {
            // Store user permissions in session
            this.currentUser = {
                username: username,
                accessLevel: user.accessLevel,
                permissions: user.permissions,
                loginTime: Date.now()
            };
            return true;
        }
        return false;
    }

    generateToken() {
        // Generate a simple token - in production, use JWT or similar
        return 'token_' + Math.random().toString(36).substr(2, 9) + '_' + Date.now();
    }

    startSessionTimer() {
        // Clear existing timer
        if (this.sessionTimer) {
            clearTimeout(this.sessionTimer);
        }

        // Set new timer
        this.sessionTimer = setTimeout(() => {
            this.logout();
        }, this.sessionTimeout);
    }

    logout() {
        // Clear session data
        localStorage.removeItem('authToken');
        localStorage.removeItem('currentUser');
        localStorage.removeItem('loginTime');
        
        // Clear timer
        if (this.sessionTimer) {
            clearTimeout(this.sessionTimer);
            this.sessionTimer = null;
        }
        
        this.isAuthenticated = false;
        this.currentUser = null;
        
        // Redirect to login page
        window.location.href = 'login.html';
    }

    addLogoutButton() {
        // Add logout button to main app if authenticated
        if (this.isAuthenticated && !window.location.pathname.includes('login.html')) {
            this.createLogoutButton();
        }
    }

    createLogoutButton() {
        // Find the header to add logout button
        const header = document.querySelector('header');
        if (header) {
            const logoutContainer = document.createElement('div');
            logoutContainer.style.cssText = `
                position: absolute;
                top: 20px;
                right: 20px;
                display: flex;
                align-items: center;
                gap: 15px;
            `;

            const userInfo = document.createElement('span');
            userInfo.style.cssText = `
                color: white;
                font-size: 14px;
                opacity: 0.9;
            `;
            userInfo.textContent = `Welcome, ${this.currentUser.username}`;

            const logoutBtn = document.createElement('button');
            logoutBtn.style.cssText = `
                background: rgba(255, 255, 255, 0.2);
                color: white;
                border: 1px solid rgba(255, 255, 255, 0.3);
                padding: 8px 16px;
                border-radius: 6px;
                cursor: pointer;
                font-size: 14px;
                transition: all 0.3s ease;
            `;
            logoutBtn.textContent = 'Logout';
            logoutBtn.addEventListener('click', () => this.logout());
            logoutBtn.addEventListener('mouseenter', () => {
                logoutBtn.style.background = 'rgba(255, 255, 255, 0.3)';
            });
            logoutBtn.addEventListener('mouseleave', () => {
                logoutBtn.style.background = 'rgba(255, 255, 255, 0.2)';
            });

            logoutContainer.appendChild(userInfo);
            logoutContainer.appendChild(logoutBtn);
            header.style.position = 'relative';
            header.appendChild(logoutContainer);
        }
    }

    showMessage(message, type) {
        const errorDiv = document.getElementById('errorMessage');
        const errorText = document.getElementById('errorText');
        
        if (errorDiv && errorText) {
            errorText.textContent = message;
            errorDiv.style.display = 'block';
            
            if (type === 'success') {
                errorDiv.style.background = '#e8f5e8';
                errorDiv.style.color = '#2e7d32';
                errorDiv.style.borderLeftColor = '#4caf50';
            } else {
                errorDiv.style.background = '#ffebee';
                errorDiv.style.color = '#c62828';
                errorDiv.style.borderLeftColor = '#c62828';
            }
            
            // Auto-hide success messages
            if (type === 'success') {
                setTimeout(() => {
                    errorDiv.style.display = 'none';
                }, 3000);
            }
        }
    }

    // Public method to check if user is authenticated
    isUserAuthenticated() {
        return this.isAuthenticated;
    }

    // Public method to get current user
    getCurrentUser() {
        return this.currentUser;
    }

    // Check if user has specific permission
    hasPermission(permission) {
        if (!this.isAuthenticated || !this.currentUser) {
            return false;
        }
        return this.currentUser.permissions.includes(permission);
    }

    // Check if user is admin
    isAdmin() {
        return this.hasPermission('write');
    }

    // Check if user can write/modify data
    canWrite() {
        return this.hasPermission('write');
    }

    // Check if user can download files
    canDownload() {
        return this.hasPermission('download');
    }

    // Check if user can upload files
    canUpload() {
        return this.hasPermission('upload');
    }

    // Check if user can delete data
    canDelete() {
        return this.hasPermission('delete');
    }
}

// Initialize authentication system
window.auth = new AuthenticationSystem();

// Export for use in other files
if (typeof module !== 'undefined' && module.exports) {
    module.exports = AuthenticationSystem;
}
