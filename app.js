// Acre Profit Calculator Application

document.addEventListener('DOMContentLoaded', function() {
    // Initialize calculator
    initCalculator();

    // Initialize mobile menu
    initMobileMenu();

    // Initialize smooth scrolling
    initSmoothScroll();

    // Run initial calculation
    calculateProfit();
});

// ===== Calculator Functions =====

function initCalculator() {
    const inputs = document.querySelectorAll('.calculator-form input');
    const calculateBtn = document.getElementById('calculateBtn');

    // Add event listeners to inputs for real-time calculation
    inputs.forEach(input => {
        input.addEventListener('input', calculateProfit);
        input.addEventListener('change', calculateProfit);
    });

    // Calculate button click
    if (calculateBtn) {
        calculateBtn.addEventListener('click', function() {
            calculateProfit();
            // Visual feedback
            this.classList.add('calculating');
            setTimeout(() => this.classList.remove('calculating'), 300);
        });
    }
}

function calculateProfit() {
    // Get input values
    const acreage = parseFloat(document.getElementById('acreage').value) || 0;
    const purchasePrice = parseFloat(document.getElementById('purchasePrice').value) || 0;
    const currentValue = parseFloat(document.getElementById('currentValue').value) || 0;
    const holdingPeriod = parseFloat(document.getElementById('holdingPeriod').value) || 1;
    const annualIncome = parseFloat(document.getElementById('annualIncome').value) || 0;
    const annualExpenses = parseFloat(document.getElementById('annualExpenses').value) || 0;

    // Calculate metrics
    const pricePerAcre = acreage > 0 ? purchasePrice / acreage : 0;
    const valuePerAcre = acreage > 0 ? currentValue / acreage : 0;
    const capitalGain = currentValue - purchasePrice;
    const netAnnualIncome = annualIncome - annualExpenses;
    const totalNetIncome = netAnnualIncome * holdingPeriod;
    const totalProfit = capitalGain + totalNetIncome;
    const totalROI = purchasePrice > 0 ? (totalProfit / purchasePrice) * 100 : 0;
    const annualROI = holdingPeriod > 0 ? totalROI / holdingPeriod : 0;

    // Update display
    updateDisplay('totalProfit', formatCurrency(totalProfit));
    updateDisplay('pricePerAcre', formatCurrency(pricePerAcre));
    updateDisplay('valuePerAcre', formatCurrency(valuePerAcre));
    updateDisplay('capitalGain', formatCurrency(capitalGain));
    updateDisplay('netIncome', formatCurrency(totalNetIncome));
    updateDisplay('totalROI', formatPercent(totalROI));
    updateDisplay('annualROI', formatPercent(annualROI));

    // Update profit breakdown bar
    updateProfitBreakdown(capitalGain, totalNetIncome, totalProfit);

    // Update colors based on profit/loss
    updateProfitColors(totalProfit);
}

function updateDisplay(elementId, value) {
    const element = document.getElementById(elementId);
    if (element) {
        element.textContent = value;
    }
}

function formatCurrency(value) {
    const absValue = Math.abs(value);
    const formatted = new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(absValue);

    return value < 0 ? '-' + formatted : formatted;
}

function formatPercent(value) {
    return value.toFixed(1) + '%';
}

function updateProfitBreakdown(capitalGain, netIncome, totalProfit) {
    const capitalGainBar = document.getElementById('capitalGainBar');
    const incomeBar = document.getElementById('incomeBar');

    if (!capitalGainBar || !incomeBar) return;

    if (totalProfit <= 0) {
        capitalGainBar.style.width = '50%';
        incomeBar.style.width = '50%';
        return;
    }

    const capitalGainPercent = Math.max(0, (capitalGain / totalProfit) * 100);
    const incomePercent = Math.max(0, (netIncome / totalProfit) * 100);

    // Ensure minimum visibility
    const minWidth = 15;
    let cgWidth = capitalGainPercent;
    let incWidth = incomePercent;

    if (cgWidth > 0 && cgWidth < minWidth) cgWidth = minWidth;
    if (incWidth > 0 && incWidth < minWidth) incWidth = minWidth;

    // Normalize to 100%
    const total = cgWidth + incWidth;
    if (total > 0) {
        cgWidth = (cgWidth / total) * 100;
        incWidth = (incWidth / total) * 100;
    }

    capitalGainBar.style.width = cgWidth + '%';
    incomeBar.style.width = incWidth + '%';
}

function updateProfitColors(totalProfit) {
    const highlightResult = document.querySelector('.highlight-result');
    if (!highlightResult) return;

    if (totalProfit < 0) {
        highlightResult.style.background = 'linear-gradient(135deg, #c0392b 0%, #e74c3c 100%)';
    } else {
        highlightResult.style.background = 'linear-gradient(135deg, #2d5a27 0%, #4a7c43 100%)';
    }
}

// ===== Mobile Menu =====

function initMobileMenu() {
    const menuBtn = document.querySelector('.mobile-menu-btn');
    const navLinks = document.querySelector('.nav-links');

    if (menuBtn && navLinks) {
        menuBtn.addEventListener('click', function() {
            navLinks.classList.toggle('active');
            this.classList.toggle('active');
        });

        // Close menu when clicking a link
        navLinks.querySelectorAll('a').forEach(link => {
            link.addEventListener('click', function() {
                navLinks.classList.remove('active');
                menuBtn.classList.remove('active');
            });
        });
    }
}

// ===== Smooth Scrolling =====

function initSmoothScroll() {
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function(e) {
            e.preventDefault();
            const targetId = this.getAttribute('href');
            const targetElement = document.querySelector(targetId);

            if (targetElement) {
                const navHeight = document.querySelector('.navbar').offsetHeight;
                const targetPosition = targetElement.offsetTop - navHeight - 20;

                window.scrollTo({
                    top: targetPosition,
                    behavior: 'smooth'
                });
            }
        });
    });
}

// ===== Utility Functions =====

// Debounce function for performance optimization
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Add scroll behavior for navbar
window.addEventListener('scroll', function() {
    const navbar = document.querySelector('.navbar');
    if (window.scrollY > 50) {
        navbar.classList.add('scrolled');
    } else {
        navbar.classList.remove('scrolled');
    }
});
