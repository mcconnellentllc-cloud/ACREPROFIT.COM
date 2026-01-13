// Pioneer Analytics - Grower Data Dashboard Application
// Private internal tool for sales forecasting and historical analysis

const ACCESS_CODE = 'pioneer2024'; // Change this to your secure access code
const STORAGE_KEY = 'pioneer_grower_data';
const AUTH_KEY = 'pioneer_authenticated';

// Sample data structure
let growerData = [];
let charts = {};

// Years for analysis
const YEARS = [2022, 2023, 2024, 2025, 2026];
const FORECAST_YEAR = 2027;

// Product categories
const PRODUCTS = ['Corn Seed', 'Soybean Seed', 'Herbicide', 'Fungicide', 'Insecticide', 'Fertilizer', 'Equipment', 'Other'];

// ============================================
// AUTHENTICATION
// ============================================

function checkAuth() {
    const isAuth = localStorage.getItem(AUTH_KEY);
    if (isAuth === 'true') {
        showDashboard();
    }
}

document.getElementById('loginForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const password = document.getElementById('password').value;

    if (password === ACCESS_CODE) {
        localStorage.setItem(AUTH_KEY, 'true');
        showDashboard();
    } else {
        document.getElementById('loginError').textContent = 'Invalid access code';
    }
});

function showDashboard() {
    document.getElementById('loginModal').classList.add('hidden');
    document.getElementById('dashboard').classList.remove('hidden');
    loadData();
    initializeCharts();
}

function logout() {
    localStorage.removeItem(AUTH_KEY);
    location.reload();
}

// ============================================
// DATA MANAGEMENT
// ============================================

function loadData() {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
        growerData = JSON.parse(stored);
    }
    updateAllDisplays();
}

function saveData() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(growerData));
    updateAllDisplays();
}

function addDataEntry(entry) {
    entry.id = Date.now() + Math.random().toString(36).substr(2, 9);
    growerData.push(entry);
    saveData();
}

function deleteEntry(id) {
    growerData = growerData.filter(d => d.id !== id);
    saveData();
}

function clearAllData() {
    if (confirm('Are you sure you want to delete all data? This cannot be undone.')) {
        growerData = [];
        saveData();
    }
}

// ============================================
// DATA ANALYSIS
// ============================================

function getYearlyData(year) {
    return growerData.filter(d => new Date(d.date).getFullYear() === year);
}

function getYearlyRevenue(year) {
    return getYearlyData(year).reduce((sum, d) => sum + parseFloat(d.amount || 0), 0);
}

function getYearlyInvoiceCount(year) {
    const invoices = new Set(getYearlyData(year).map(d => d.invoice_number));
    return invoices.size || getYearlyData(year).length;
}

function getUniqueGrowers() {
    const growers = new Set(growerData.map(d => d.grower_name));
    return [...growers];
}

function getGrowerRevenue(growerName, year = null) {
    let data = growerData.filter(d => d.grower_name === growerName);
    if (year) {
        data = data.filter(d => new Date(d.date).getFullYear() === year);
    }
    return data.reduce((sum, d) => sum + parseFloat(d.amount || 0), 0);
}

function getProductRevenue(product, year = null) {
    let data = growerData.filter(d => d.product === product);
    if (year) {
        data = data.filter(d => new Date(d.date).getFullYear() === year);
    }
    return data.reduce((sum, d) => sum + parseFloat(d.amount || 0), 0);
}

function getMonthlyRevenue(year) {
    const monthly = Array(12).fill(0);
    getYearlyData(year).forEach(d => {
        const month = new Date(d.date).getMonth();
        monthly[month] += parseFloat(d.amount || 0);
    });
    return monthly;
}

// ============================================
// FORECASTING ALGORITHMS
// ============================================

function calculateGrowthRate() {
    const revenues = YEARS.map(y => getYearlyRevenue(y)).filter(r => r > 0);
    if (revenues.length < 2) return 0;

    let totalGrowth = 0;
    let growthCount = 0;

    for (let i = 1; i < revenues.length; i++) {
        if (revenues[i-1] > 0) {
            totalGrowth += (revenues[i] - revenues[i-1]) / revenues[i-1];
            growthCount++;
        }
    }

    return growthCount > 0 ? totalGrowth / growthCount : 0;
}

function calculateLinearTrend(values) {
    const n = values.length;
    if (n < 2) return { slope: 0, intercept: values[0] || 0 };

    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;

    for (let i = 0; i < n; i++) {
        sumX += i;
        sumY += values[i];
        sumXY += i * values[i];
        sumX2 += i * i;
    }

    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    const intercept = (sumY - slope * sumX) / n;

    return { slope, intercept };
}

function forecast2027Revenue() {
    const revenues = YEARS.map(y => getYearlyRevenue(y));
    const validRevenues = revenues.filter(r => r > 0);

    if (validRevenues.length === 0) {
        return { conservative: 0, likely: 0, optimistic: 0 };
    }

    // Linear trend projection
    const trend = calculateLinearTrend(revenues);
    const trendProjection = trend.slope * YEARS.length + trend.intercept;

    // Growth rate projection
    const growthRate = calculateGrowthRate();
    const lastRevenue = revenues[revenues.length - 1] || validRevenues[validRevenues.length - 1];
    const growthProjection = lastRevenue * (1 + growthRate);

    // Average of methods
    const avgProjection = (trendProjection + growthProjection) / 2;

    // Confidence intervals
    const volatility = calculateVolatility(revenues);

    return {
        conservative: Math.max(0, avgProjection * (1 - volatility * 0.5)),
        likely: Math.max(0, avgProjection),
        optimistic: avgProjection * (1 + volatility * 0.5)
    };
}

function calculateVolatility(values) {
    const validValues = values.filter(v => v > 0);
    if (validValues.length < 2) return 0.1;

    const mean = validValues.reduce((a, b) => a + b, 0) / validValues.length;
    const variance = validValues.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / validValues.length;
    const stdDev = Math.sqrt(variance);

    return Math.min(0.5, stdDev / mean);
}

function forecastGrowerRevenue(growerName) {
    const revenues = YEARS.map(y => getGrowerRevenue(growerName, y));
    const validRevenues = revenues.filter(r => r > 0);

    if (validRevenues.length === 0) return 0;
    if (validRevenues.length === 1) return validRevenues[0];

    // Weighted average favoring recent years
    const weights = [0.1, 0.15, 0.2, 0.25, 0.3];
    let weightedSum = 0;
    let weightSum = 0;

    revenues.forEach((r, i) => {
        if (r > 0) {
            weightedSum += r * weights[i];
            weightSum += weights[i];
        }
    });

    const baseProjection = weightedSum / weightSum;
    const growthRate = calculateGrowthRate();

    return baseProjection * (1 + growthRate * 0.5);
}

function forecastMonthlyRevenue() {
    // Average monthly distribution across years
    const monthlyTotals = Array(12).fill(0);
    const monthCounts = Array(12).fill(0);

    YEARS.forEach(year => {
        const monthly = getMonthlyRevenue(year);
        monthly.forEach((amount, month) => {
            if (amount > 0) {
                monthlyTotals[month] += amount;
                monthCounts[month]++;
            }
        });
    });

    const avgMonthly = monthlyTotals.map((total, i) =>
        monthCounts[i] > 0 ? total / monthCounts[i] : 0
    );

    // Apply growth projection
    const forecast = forecast2027Revenue();
    const historicalTotal = avgMonthly.reduce((a, b) => a + b, 0);

    if (historicalTotal === 0) return Array(12).fill(0);

    const scaleFactor = forecast.likely / historicalTotal;
    return avgMonthly.map(m => m * scaleFactor);
}

// ============================================
// CHART INITIALIZATION & UPDATES
// ============================================

const chartColors = {
    primary: '#1a5f2a',
    primaryLight: 'rgba(26, 95, 42, 0.1)',
    accent: '#e8b923',
    accentLight: 'rgba(232, 185, 35, 0.3)',
    danger: '#dc3545',
    success: '#28a745',
    blues: ['#1e88e5', '#42a5f5', '#64b5f6', '#90caf9', '#bbdefb'],
    greens: ['#0d4a1a', '#1a5f2a', '#2e7d32', '#43a047', '#66bb6a']
};

function initializeCharts() {
    createRevenueChart();
    createInvoiceChart();
    createMonthlyTrendChart();
    createAvgInvoiceChart();
    createGrowthChart();
    createForecastChart();
    createProbabilityChart();
    createMonthlyForecastChart();
    createTopGrowersChart();
    createRetentionChart();
    createNewVsReturningChart();
    createSpendDistributionChart();
    createProductCategoryChart();
    createProductTrendChart();
    createProductForecastChart();
}

function createRevenueChart() {
    const ctx = document.getElementById('revenueChart').getContext('2d');
    charts.revenue = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: YEARS,
            datasets: [{
                label: 'Revenue',
                data: YEARS.map(y => getYearlyRevenue(y)),
                backgroundColor: chartColors.primary,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => '$' + ctx.raw.toLocaleString()
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: value => '$' + value.toLocaleString()
                    }
                }
            }
        }
    });
}

function createInvoiceChart() {
    const ctx = document.getElementById('invoiceChart').getContext('2d');
    charts.invoice = new Chart(ctx, {
        type: 'line',
        data: {
            labels: YEARS,
            datasets: [{
                label: 'Invoices',
                data: YEARS.map(y => getYearlyInvoiceCount(y)),
                borderColor: chartColors.accent,
                backgroundColor: chartColors.accentLight,
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            plugins: { legend: { display: false } },
            scales: { y: { beginAtZero: true } }
        }
    });
}

function createMonthlyTrendChart() {
    const ctx = document.getElementById('monthlyTrendChart').getContext('2d');
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    const datasets = YEARS.map((year, i) => ({
        label: year.toString(),
        data: getMonthlyRevenue(year),
        borderColor: chartColors.greens[i % chartColors.greens.length],
        backgroundColor: 'transparent',
        tension: 0.3
    }));

    charts.monthlyTrend = new Chart(ctx, {
        type: 'line',
        data: { labels: months, datasets },
        options: {
            responsive: true,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.dataset.label + ': $' + ctx.raw.toLocaleString()
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { callback: value => '$' + value.toLocaleString() }
                }
            }
        }
    });
}

function createAvgInvoiceChart() {
    const ctx = document.getElementById('avgInvoiceChart').getContext('2d');
    const avgValues = YEARS.map(y => {
        const count = getYearlyInvoiceCount(y);
        return count > 0 ? getYearlyRevenue(y) / count : 0;
    });

    charts.avgInvoice = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: YEARS,
            datasets: [{
                label: 'Avg Invoice Value',
                data: avgValues,
                backgroundColor: chartColors.blues[0],
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => '$' + ctx.raw.toLocaleString(undefined, {maximumFractionDigits: 0})
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { callback: value => '$' + value.toLocaleString() }
                }
            }
        }
    });
}

function createGrowthChart() {
    const ctx = document.getElementById('growthChart').getContext('2d');
    const revenues = YEARS.map(y => getYearlyRevenue(y));
    const growth = revenues.map((r, i) => {
        if (i === 0 || revenues[i-1] === 0) return 0;
        return ((r - revenues[i-1]) / revenues[i-1]) * 100;
    });

    charts.growth = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: YEARS,
            datasets: [{
                label: 'YoY Growth',
                data: growth,
                backgroundColor: growth.map(g => g >= 0 ? chartColors.success : chartColors.danger),
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.raw.toFixed(1) + '%'
                    }
                }
            },
            scales: {
                y: {
                    ticks: { callback: value => value + '%' }
                }
            }
        }
    });
}

function createForecastChart() {
    const ctx = document.getElementById('forecastChart').getContext('2d');
    const revenues = YEARS.map(y => getYearlyRevenue(y));
    const forecast = forecast2027Revenue();

    charts.forecast = new Chart(ctx, {
        type: 'line',
        data: {
            labels: [...YEARS, FORECAST_YEAR],
            datasets: [
                {
                    label: 'Historical',
                    data: [...revenues, null],
                    borderColor: chartColors.primary,
                    backgroundColor: chartColors.primaryLight,
                    fill: true,
                    tension: 0.3
                },
                {
                    label: 'Forecast Range',
                    data: [...Array(5).fill(null), forecast.optimistic],
                    borderColor: chartColors.accent,
                    backgroundColor: 'rgba(232, 185, 35, 0.2)',
                    borderDash: [5, 5],
                    pointRadius: 8,
                    pointBackgroundColor: chartColors.accent
                },
                {
                    label: 'Conservative',
                    data: [...Array(5).fill(null), forecast.conservative],
                    borderColor: chartColors.blues[0],
                    borderDash: [5, 5],
                    pointRadius: 6,
                    pointBackgroundColor: chartColors.blues[0]
                }
            ]
        },
        options: {
            responsive: true,
            plugins: {
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.dataset.label + ': $' + (ctx.raw || 0).toLocaleString()
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { callback: value => '$' + value.toLocaleString() }
                }
            }
        }
    });
}

function createProbabilityChart() {
    const ctx = document.getElementById('probabilityChart').getContext('2d');
    const forecast = forecast2027Revenue();

    // Generate probability distribution
    const mean = forecast.likely;
    const stdDev = (forecast.optimistic - forecast.conservative) / 4;
    const labels = [];
    const values = [];

    for (let i = 0; i <= 10; i++) {
        const x = forecast.conservative + (forecast.optimistic - forecast.conservative) * (i / 10);
        labels.push('$' + (x / 1000).toFixed(0) + 'k');
        // Normal distribution approximation
        const z = (x - mean) / (stdDev || 1);
        values.push(Math.exp(-0.5 * z * z));
    }

    charts.probability = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [{
                label: 'Probability',
                data: values,
                borderColor: chartColors.primary,
                backgroundColor: chartColors.primaryLight,
                fill: true,
                tension: 0.4
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => 'Relative Likelihood: ' + (ctx.raw * 100).toFixed(0) + '%'
                    }
                }
            },
            scales: {
                y: {
                    display: false
                }
            }
        }
    });
}

function createMonthlyForecastChart() {
    const ctx = document.getElementById('monthlyForecastChart').getContext('2d');
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const forecastMonthly = forecastMonthlyRevenue();

    // Last year for comparison
    const lastYear = YEARS[YEARS.length - 1];
    const lastYearMonthly = getMonthlyRevenue(lastYear);

    charts.monthlyForecast = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: months,
            datasets: [
                {
                    label: lastYear.toString(),
                    data: lastYearMonthly,
                    backgroundColor: 'rgba(26, 95, 42, 0.3)',
                    borderRadius: 4
                },
                {
                    label: '2027 Forecast',
                    data: forecastMonthly,
                    backgroundColor: chartColors.accent,
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            plugins: {
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.dataset.label + ': $' + ctx.raw.toLocaleString()
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { callback: value => '$' + value.toLocaleString() }
                }
            }
        }
    });
}

function createTopGrowersChart() {
    const ctx = document.getElementById('topGrowersChart').getContext('2d');
    const growers = getUniqueGrowers();
    const growerRevenues = growers.map(g => ({
        name: g,
        revenue: getGrowerRevenue(g)
    })).sort((a, b) => b.revenue - a.revenue).slice(0, 10);

    charts.topGrowers = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: growerRevenues.map(g => g.name.length > 15 ? g.name.substring(0, 15) + '...' : g.name),
            datasets: [{
                label: 'Total Revenue',
                data: growerRevenues.map(g => g.revenue),
                backgroundColor: chartColors.primary,
                borderRadius: 6
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => '$' + ctx.raw.toLocaleString()
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: { callback: value => '$' + value.toLocaleString() }
                }
            }
        }
    });
}

function createRetentionChart() {
    const ctx = document.getElementById('retentionChart').getContext('2d');

    // Calculate retention rates
    const retentionRates = [];
    for (let i = 1; i < YEARS.length; i++) {
        const prevYearGrowers = new Set(getYearlyData(YEARS[i-1]).map(d => d.grower_name));
        const currYearGrowers = new Set(getYearlyData(YEARS[i]).map(d => d.grower_name));

        let retained = 0;
        prevYearGrowers.forEach(g => {
            if (currYearGrowers.has(g)) retained++;
        });

        const rate = prevYearGrowers.size > 0 ? (retained / prevYearGrowers.size) * 100 : 0;
        retentionRates.push(rate);
    }

    charts.retention = new Chart(ctx, {
        type: 'line',
        data: {
            labels: YEARS.slice(1),
            datasets: [{
                label: 'Retention Rate',
                data: retentionRates,
                borderColor: chartColors.success,
                backgroundColor: 'rgba(40, 167, 69, 0.1)',
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.raw.toFixed(1) + '% retained'
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100,
                    ticks: { callback: value => value + '%' }
                }
            }
        }
    });
}

function createNewVsReturningChart() {
    const ctx = document.getElementById('newVsReturningChart').getContext('2d');

    const newGrowers = [];
    const returningGrowers = [];
    let allPreviousGrowers = new Set();

    YEARS.forEach((year, i) => {
        const yearGrowers = new Set(getYearlyData(year).map(d => d.grower_name));
        let newCount = 0;
        let returningCount = 0;

        yearGrowers.forEach(g => {
            if (allPreviousGrowers.has(g)) {
                returningCount++;
            } else {
                newCount++;
            }
        });

        newGrowers.push(newCount);
        returningGrowers.push(returningCount);

        yearGrowers.forEach(g => allPreviousGrowers.add(g));
    });

    charts.newVsReturning = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: YEARS,
            datasets: [
                {
                    label: 'New Growers',
                    data: newGrowers,
                    backgroundColor: chartColors.accent,
                    borderRadius: 4
                },
                {
                    label: 'Returning',
                    data: returningGrowers,
                    backgroundColor: chartColors.primary,
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom' }
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true, beginAtZero: true }
            }
        }
    });
}

function createSpendDistributionChart() {
    const ctx = document.getElementById('spendDistributionChart').getContext('2d');
    const growers = getUniqueGrowers();

    const brackets = [
        { label: '<$10k', min: 0, max: 10000, count: 0 },
        { label: '$10k-$25k', min: 10000, max: 25000, count: 0 },
        { label: '$25k-$50k', min: 25000, max: 50000, count: 0 },
        { label: '$50k-$100k', min: 50000, max: 100000, count: 0 },
        { label: '>$100k', min: 100000, max: Infinity, count: 0 }
    ];

    growers.forEach(g => {
        const revenue = getGrowerRevenue(g);
        const bracket = brackets.find(b => revenue >= b.min && revenue < b.max);
        if (bracket) bracket.count++;
    });

    charts.spendDistribution = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: brackets.map(b => b.label),
            datasets: [{
                data: brackets.map(b => b.count),
                backgroundColor: chartColors.greens
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom' }
            }
        }
    });
}

function createProductCategoryChart() {
    const ctx = document.getElementById('productCategoryChart').getContext('2d');

    const productRevenues = PRODUCTS.map(p => getProductRevenue(p)).filter(r => r > 0);
    const productLabels = PRODUCTS.filter(p => getProductRevenue(p) > 0);

    charts.productCategory = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: productLabels,
            datasets: [{
                data: productRevenues,
                backgroundColor: [...chartColors.greens, ...chartColors.blues]
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.label + ': $' + ctx.raw.toLocaleString()
                    }
                }
            }
        }
    });
}

function createProductTrendChart() {
    const ctx = document.getElementById('productTrendChart').getContext('2d');

    const datasets = PRODUCTS.slice(0, 5).map((product, i) => ({
        label: product,
        data: YEARS.map(y => getProductRevenue(product, y)),
        borderColor: [...chartColors.greens, ...chartColors.blues][i],
        backgroundColor: 'transparent',
        tension: 0.3
    }));

    charts.productTrend = new Chart(ctx, {
        type: 'line',
        data: { labels: YEARS, datasets },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.dataset.label + ': $' + ctx.raw.toLocaleString()
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { callback: value => '$' + value.toLocaleString() }
                }
            }
        }
    });
}

function createProductForecastChart() {
    const ctx = document.getElementById('productForecastChart').getContext('2d');

    // Project product mix for 2027 based on trends
    const productForecasts = PRODUCTS.map(product => {
        const revenues = YEARS.map(y => getProductRevenue(product, y));
        const trend = calculateLinearTrend(revenues);
        return {
            product,
            forecast: Math.max(0, trend.slope * YEARS.length + trend.intercept)
        };
    }).filter(p => p.forecast > 0);

    charts.productForecast = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: productForecasts.map(p => p.product),
            datasets: [{
                label: '2027 Projected Revenue',
                data: productForecasts.map(p => p.forecast),
                backgroundColor: chartColors.accent,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => '$' + ctx.raw.toLocaleString()
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { callback: value => '$' + value.toLocaleString() }
                }
            }
        }
    });
}

// ============================================
// UI UPDATES
// ============================================

function updateAllDisplays() {
    updateSummaryStats();
    updateGrowerTable();
    updateRawDataTable();
    updateCharts();
}

function updateSummaryStats() {
    const totalRevenue = growerData.reduce((sum, d) => sum + parseFloat(d.amount || 0), 0);
    const uniqueGrowers = getUniqueGrowers().length;
    const growthRate = calculateGrowthRate() * 100;
    const forecast = forecast2027Revenue();

    document.getElementById('totalGrowers').textContent = uniqueGrowers.toLocaleString();
    document.getElementById('totalRevenue').textContent = '$' + totalRevenue.toLocaleString();
    document.getElementById('avgGrowth').textContent = growthRate.toFixed(1) + '%';
    document.getElementById('projected2027').textContent = '$' + forecast.likely.toLocaleString();

    // Update forecast details
    document.getElementById('conservativeEstimate').textContent = '$' + forecast.conservative.toLocaleString();
    document.getElementById('likelyEstimate').textContent = '$' + forecast.likely.toLocaleString();
    document.getElementById('optimisticEstimate').textContent = '$' + forecast.optimistic.toLocaleString();
}

function updateGrowerTable() {
    const tbody = document.getElementById('growerTableBody');
    const growers = getUniqueGrowers();

    tbody.innerHTML = growers.map(grower => {
        const yearRevenues = YEARS.map(y => getGrowerRevenue(grower, y));
        const total = yearRevenues.reduce((a, b) => a + b, 0);
        const projected = forecastGrowerRevenue(grower);

        return `
            <tr>
                <td>${grower}</td>
                ${yearRevenues.map(r => `<td class="amount">$${r.toLocaleString()}</td>`).join('')}
                <td class="amount">$${total.toLocaleString()}</td>
                <td class="projected">$${projected.toLocaleString()}</td>
            </tr>
        `;
    }).join('');
}

function updateRawDataTable() {
    const tbody = document.getElementById('rawDataBody');
    const yearFilter = document.getElementById('yearFilter').value;

    let filteredData = [...growerData];
    if (yearFilter !== 'all') {
        filteredData = filteredData.filter(d =>
            new Date(d.date).getFullYear() === parseInt(yearFilter)
        );
    }

    // Sort by date descending
    filteredData.sort((a, b) => new Date(b.date) - new Date(a.date));

    tbody.innerHTML = filteredData.map(d => `
        <tr>
            <td>${new Date(d.date).toLocaleDateString()}</td>
            <td>${d.invoice_number || '-'}</td>
            <td>${d.grower_name}</td>
            <td>${d.product}</td>
            <td>${d.quantity || '-'}</td>
            <td class="amount">$${parseFloat(d.amount).toLocaleString()}</td>
            <td><button class="btn-delete" onclick="deleteEntry('${d.id}')">Delete</button></td>
        </tr>
    `).join('');
}

function updateCharts() {
    // Update all chart data
    Object.keys(charts).forEach(key => {
        if (charts[key]) {
            charts[key].destroy();
        }
    });

    initializeCharts();
}

function filterDataByYear() {
    updateRawDataTable();
}

// ============================================
// FILE UPLOAD & DATA IMPORT
// ============================================

function showUploadModal() {
    document.getElementById('uploadModal').classList.remove('hidden');
}

function closeUploadModal() {
    document.getElementById('uploadModal').classList.add('hidden');
}

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        parseCSV(text);
    };
    reader.readAsText(file);
}

function parseCSV(text) {
    const lines = text.split('\n').filter(line => line.trim());
    const headers = lines[0].toLowerCase().split(',').map(h => h.trim());

    let imported = 0;
    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',').map(v => v.trim());
        if (values.length >= 3) {
            const entry = {
                date: values[headers.indexOf('date')] || values[0] || new Date().toISOString().split('T')[0],
                invoice_number: values[headers.indexOf('invoice_number')] || values[headers.indexOf('invoice')] || values[1] || '',
                grower_name: values[headers.indexOf('grower_name')] || values[headers.indexOf('grower')] || values[2] || 'Unknown',
                product: values[headers.indexOf('product')] || values[3] || 'Other',
                quantity: parseFloat(values[headers.indexOf('quantity')] || values[4]) || 0,
                amount: parseFloat(values[headers.indexOf('amount')] || values[5]) || 0
            };

            if (entry.amount > 0 || entry.grower_name !== 'Unknown') {
                addDataEntry(entry);
                imported++;
            }
        }
    }

    alert(`Successfully imported ${imported} records`);
    closeUploadModal();
}

function processBulkData() {
    const text = document.getElementById('bulkData').value;
    if (!text.trim()) {
        alert('Please paste some data first');
        return;
    }

    const lines = text.split('\n').filter(line => line.trim());
    let imported = 0;

    lines.forEach(line => {
        // Support both comma and tab separated
        const values = line.includes('\t') ? line.split('\t') : line.split(',');
        values.forEach((v, i) => values[i] = v.trim());

        if (values.length >= 3) {
            const entry = {
                date: values[0] || new Date().toISOString().split('T')[0],
                invoice_number: values[1] || '',
                grower_name: values[2] || 'Unknown',
                product: values[3] || 'Other',
                quantity: parseFloat(values[4]) || 0,
                amount: parseFloat(values[5]) || 0
            };

            if (entry.grower_name !== 'Unknown') {
                addDataEntry(entry);
                imported++;
            }
        }
    });

    alert(`Successfully imported ${imported} records`);
    document.getElementById('bulkData').value = '';
    closeUploadModal();
}

// Manual entry form
document.getElementById('manualEntryForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const formData = new FormData(this);

    const entry = {
        date: formData.get('date'),
        invoice_number: formData.get('invoice_number'),
        grower_name: formData.get('grower_name'),
        product: formData.get('product'),
        quantity: parseFloat(formData.get('quantity')) || 0,
        amount: parseFloat(formData.get('amount')) || 0
    };

    addDataEntry(entry);
    this.reset();
    alert('Entry added successfully');
});

// ============================================
// EXPORT
// ============================================

function exportData() {
    const forecast = forecast2027Revenue();

    let csvContent = 'Pioneer Analytics - Grower Data Report\n';
    csvContent += `Generated: ${new Date().toLocaleDateString()}\n\n`;

    // Summary
    csvContent += 'SUMMARY\n';
    csvContent += `Total Growers,${getUniqueGrowers().length}\n`;
    csvContent += `Total Revenue (5yr),$${growerData.reduce((s, d) => s + parseFloat(d.amount || 0), 0).toLocaleString()}\n`;
    csvContent += `2027 Projected (Conservative),$${forecast.conservative.toLocaleString()}\n`;
    csvContent += `2027 Projected (Likely),$${forecast.likely.toLocaleString()}\n`;
    csvContent += `2027 Projected (Optimistic),$${forecast.optimistic.toLocaleString()}\n\n`;

    // Yearly breakdown
    csvContent += 'YEARLY REVENUE\n';
    csvContent += 'Year,Revenue,Invoice Count,Avg Invoice\n';
    YEARS.forEach(year => {
        const revenue = getYearlyRevenue(year);
        const count = getYearlyInvoiceCount(year);
        const avg = count > 0 ? revenue / count : 0;
        csvContent += `${year},$${revenue.toLocaleString()},${count},$${avg.toLocaleString()}\n`;
    });
    csvContent += '\n';

    // Grower breakdown
    csvContent += 'GROWER SUMMARY\n';
    csvContent += 'Grower,' + YEARS.join(',') + ',Total,2027 Projected\n';
    getUniqueGrowers().forEach(grower => {
        const yearRevs = YEARS.map(y => getGrowerRevenue(grower, y));
        const total = yearRevs.reduce((a, b) => a + b, 0);
        const projected = forecastGrowerRevenue(grower);
        csvContent += `${grower},${yearRevs.map(r => '$' + r.toLocaleString()).join(',')},$${total.toLocaleString()},$${projected.toLocaleString()}\n`;
    });
    csvContent += '\n';

    // Raw data
    csvContent += 'RAW DATA\n';
    csvContent += 'Date,Invoice,Grower,Product,Quantity,Amount\n';
    growerData.forEach(d => {
        csvContent += `${d.date},${d.invoice_number},${d.grower_name},${d.product},${d.quantity},$${d.amount}\n`;
    });

    // Download
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `pioneer_analytics_report_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
}

// ============================================
// TAB NAVIGATION
// ============================================

document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', function() {
        const tab = this.dataset.tab;

        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

        this.classList.add('active');
        document.getElementById(tab).classList.add('active');
    });
});

// Upload tab navigation
document.querySelectorAll('.upload-tab').forEach(btn => {
    btn.addEventListener('click', function() {
        const tab = this.dataset.upload;

        document.querySelectorAll('.upload-tab').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.upload-content').forEach(c => c.classList.remove('active'));

        this.classList.add('active');
        document.getElementById(tab + '-upload').classList.add('active');
    });
});

// Drag and drop
const dropZone = document.getElementById('dropZone');
dropZone.addEventListener('click', () => document.getElementById('csvFile').click());
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.style.borderColor = '#1a5f2a';
    dropZone.style.background = '#f8fff8';
});
dropZone.addEventListener('dragleave', () => {
    dropZone.style.borderColor = '#e0e0e0';
    dropZone.style.background = 'transparent';
});
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.style.borderColor = '#e0e0e0';
    dropZone.style.background = 'transparent';

    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.csv')) {
        const reader = new FileReader();
        reader.onload = (e) => parseCSV(e.target.result);
        reader.readAsText(file);
    }
});

// Initialize
checkAuth();
