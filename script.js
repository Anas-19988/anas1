// 1. تنظيف الرابط تلقائياً وحذف علامات الهاش العالقة من الكاش لمنع خطأ الـ 404
if (window.location.hash) {
    window.history.replaceState("", document.title, window.location.pathname);
}

// تشغيل العمليات فور اكتمال بناء واجهة الويب
document.addEventListener("DOMContentLoaded", () => {
    initNetworkMonitor();
    initCharts();
    initAuthSystem();
    checkLoginStatus(); 
    
    const uploader = document.getElementById('excel-uploader');
    if (uploader) {
        uploader.addEventListener('change', handleFileUpload);
    }
});

let revenueChartInstance = null;
let expenseChartInstance = null;

// 2. نظام إدارة بوابات الدخول والتحكم بالتبديل بين الواجهة التعريفية ولوحة التحكم
function initAuthSystem() {
    const openLoginBtn = document.getElementById('open-login-btn');
    const heroActionBtn = document.getElementById('hero-action-btn');
    const closeLoginBtn = document.getElementById('close-login-btn');
    const loginModal = document.getElementById('login-modal');
    const loginForm = document.getElementById('login-form');
    const logoutBtn = document.getElementById('logout-btn');

    if (openLoginBtn) openLoginBtn.addEventListener('click', () => loginModal.classList.remove('hidden'));
    if (heroActionBtn) heroActionBtn.addEventListener('click', () => loginModal.classList.remove('hidden'));
    if (closeLoginBtn) closeLoginBtn.addEventListener('click', () => loginModal.classList.add('hidden'));

    if (loginForm) {
        loginForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const user = document.getElementById('username').value.trim().toLowerCase();
            const errorEl = document.getElementById('login-error');

            if (user === 'anas' || user === 'admin') {
                localStorage.setItem('isLoggedIn', 'true');
                localStorage.setItem('sessionUser', user);
                errorEl.classList.add('hidden');
                loginModal.classList.add('hidden');
                checkLoginStatus();
            } else {
                errorEl.classList.remove('hidden');
            }
        });
    }

    if (logoutBtn) {
        logoutBtn.addEventListener('click', () => {
            localStorage.removeItem('isLoggedIn');
            localStorage.removeItem('sessionUser');
            checkLoginStatus();
        });
    }
}

function checkLoginStatus() {
    const isLoggedIn = localStorage.getItem('isLoggedIn') === 'true';
    const landingPage = document.getElementById('landing-page');
    const dashboardPage = document.getElementById('dashboard-page');
    const userDisplay = document.getElementById('user-display-name');

    if (isLoggedIn) {
        if (landingPage) landingPage.classList.add('hidden');
        if (dashboardPage) dashboardPage.classList.remove('hidden');
        if (userDisplay) userDisplay.innerText = `مرحباً، ${localStorage.getItem('sessionUser')}`;
        loadSavedData(); 
    } else {
        if (dashboardPage) dashboardPage.classList.add('hidden');
        if (landingPage) landingPage.classList.remove('hidden');
    }
}

// 3. مراقبة حالة اتصال المستخدم بالإنترنت
function initNetworkMonitor() {
    const toast = document.getElementById('offline-toast');
    if (!toast) return;
    window.addEventListener('offline', () => toast.classList.remove('hidden'));
    window.addEventListener('online', () => toast.classList.add('hidden'));
    if (!navigator.onLine) toast.classList.remove('hidden');
}

// 4. تهيئة الرسوم البيانية بهياكل مالية فارغة
function initCharts() {
    const ctxRev = document.getElementById('revenueChart');
    if (ctxRev) {
        revenueChartInstance = new Chart(ctxRev.getContext('2d'), {
            type: 'bar',
            data: { 
                labels: ['إجمالي الإيرادات', 'إجمالي المصروفات', 'صافي الربح'], 
                datasets: [{ label: 'قائمة الدخل (ر.س)', data: [0, 0, 0], backgroundColor: ['#112233', '#e11d48', '#2ec4b6'] }] 
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }

    const ctxExp = document.getElementById('expenseChart');
    if (ctxExp) {
        expenseChartInstance = new Chart(ctxExp.getContext('2d'), {
            type: 'doughnut',
            data: { 
                labels: ['الأصول', 'الالتزامات', 'حقوق الملكية'], 
                datasets: [{ data: [0, 0, 0], backgroundColor: ['#2ec4b6', '#f59e0b', '#112233'] }] 
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }
}

// 5. قراءة ملف ميزان المراجعة المرفوع
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // قراءة الملف كصفوف ومصفوفات لضمان التوافق التام مع أي هيكلية ميزانية
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        processFinancialData(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

// 6. الفرز المالي الشامل المطور بالكامل لتلافي الأخطاء الصامتة والأعمدة الفارغة
function processFinancialData(matrixData) {
    if (!matrixData || matrixData.length < 2) return;

    // تمشيط الصف الأول للعثور على مواقع الأعمدة الصحيحة تلقائياً
    const headers = matrixData[0];
    let nameIdx = -1;
    let debitIdx = -1;
    let creditIdx = -1;

    for (let col = 0; col < headers.length; col++) {
        const headerText = String(headers[col] || '').trim().toLowerCase();
        if (headerText.includes('اسم') || headerText.includes('حساب') || headerText.includes('account') || headerText.includes('name') || headerText.includes('البيان')) {
            nameIdx = col;
        } else if (headerText.includes('مدين') || headerText.includes('debit') || headerText.includes('dr')) {
            debitIdx = col;
        } else if (headerText.includes('دائن') || headerText.includes('credit') || headerText.includes('cr')) {
            creditIdx = col;
        }
    }

    // حل بديل ذكي في حال لم تكن العناوين مطابقة للكلمات المفتاحية
    if (nameIdx === -1 || debitIdx === -1 || creditIdx === -1) {
        if (headers.length >= 4) {
            nameIdx = 1;
            debitIdx = 2;
            creditIdx = 3;
        } else {
            nameIdx = 0;
            debitIdx = 1;
            creditIdx = 2;
        }
    }

    let totalRevenue = 0;     
    let totalExpenses = 0;    
    let totalAssets = 0;      
    let totalLiabilities = 0; 
    let totalEquity = 0;      

    // فرز وتحليل البيانات صفاً بصف
    for (let i = 1; i < matrixData.length; i++) {
        const row = matrixData[i];
        if (!row || row.length === 0) continue; // تجنب الأخطاء إذا كان الصف فارغاً بالكامل

        // جلب اسم الحساب والتنظيف الآمن للأرصدة المدججة بالقيمة الصفرية أو الفواصل
        const accountName = nameIdx !== -1 && row[nameIdx] ? String(row[nameIdx]).trim() : '';
        const debit = debitIdx !== -1 && row[debitIdx] !== undefined ? parseFloat(String(row[debitIdx]).replace(/[\s,]/g, '')) || 0 : 0;
        const credit = creditIdx !== -1 && row[creditIdx] !== undefined ? parseFloat(String(row[creditIdx]).replace(/[\s,]/g, '')) || 0 : 0;

        if (!accountName) continue;

        // قواعد تصنيف الحسابات المحاسبية المرنة بناءً على الكلمات المفتاحية
        let classified = false;
        if (accountName.includes('إيراد') || accountName.includes('مبيعات') || accountName.includes('Revenue') || accountName.includes('Sales')) {
            totalRevenue += (credit > 0 ? credit : debit); 
            classified = true;
        } 
        else if (accountName.includes('مصروف') || accountName.includes('رواتب') || accountName.includes('إيجار') || accountName.includes('تكلفة') || accountName.includes('تكاليف') || accountName.includes('استهلاك') || accountName.includes('Expenses') || accountName.includes('Cost')) {
            totalExpenses += debit; 
            classified = true;
        }
        else if (accountName.includes('نقدية') || accountName.includes('بنك') || accountName.includes('عملاء') || accountName.includes('أصول') || accountName.includes('ثابتة') || accountName.includes('مخزون') || accountName.includes('Assets') || accountName.includes('Cash')) {
            totalAssets += (debit - credit); 
            classified = true;
        }
        else if (accountName.includes('موردين') || accountName.includes('قروض') || accountName.includes('التزامات') || accountName.includes('خصوم') || accountName.includes('دائنون') || accountName.includes('Liabilities') || accountName.includes('Payable')) {
            totalLiabilities += (credit - debit); 
            classified = true;
        }
        else if (accountName.includes('رأس المال') || accountName.includes('ملكية') || accountName.includes('حقوق') || accountName.includes('أرباح مبقاة') || accountName.includes('Equity') || accountName.includes('Capital')) {
            totalEquity += (credit - debit); 
            classified = true;
        }

        // معامل الحماية التلقائي: في حال لم يطابق الحساب أي كلمة، يتم فرزه تلقائياً للحفاظ على توازن الميزانية 100%
        if (!classified) {
            if (debit > 0) {
                totalAssets += (debit - credit);
            } else if (credit > 0) {
                totalLiabilities += (credit - debit);
            }
        }
    }

    const netProfit = totalRevenue - totalExpenses;
    const finalEquity = totalEquity + netProfit;

    const financialSummary = {
        totalRev: totalRevenue,
        totalExp: totalExpenses,
        netProfit: netProfit,
        assets: Math.abs(totalAssets),
        liabilities: Math.abs(totalLiabilities),
        equity: finalEquity
    };

    localStorage.setItem('financialData', JSON.stringify(financialSummary));
    updateUI(financialSummary);
    updateChartsDynamically(financialSummary);
}

// 7. جلب وعرض البيانات المحفوظة تلقائياً عند إعادة زيارة الموقع
function loadSavedData() {
    const saved = localStorage.getItem('financialData');
    if (saved) {
        const data = JSON.parse(saved);
        updateUI(data);
        setTimeout(() => updateChartsDynamically(data), 200);
    }
}

// 8. تحديث الأرقام والنصوص في الواجهة وتفادي استخدام ترميز الأرقام المتناهية الصغر المربك للمحاسبين
function updateUI(data) {
    const revEl = document.getElementById('total-revenues');
    const expEl = document.getElementById('total-expenses');
    const profEl = document.getElementById('net-profit');
    const assetEl = document.getElementById('balance-assets');
    const eqLiabEl = document.getElementById('balance-eq-liab');

    // استخدام التنسيق الرقمي العالمي الإنجليزي لتفادي تحويل الأصفار لنقاط مبهمة
    if (revEl) revEl.innerText = data.totalRev.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ر.س';
    if (expEl) expEl.innerText = data.totalExp.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ر.س';
    if (profEl) profEl.innerText = data.netProfit.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ر.س';
    if (assetEl) assetEl.innerText = data.assets.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ر.س';
    if (eqLiabEl) eqLiabEl.innerText = (data.liabilities + data.equity).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ر.س';
}

// 9. ضخ البيانات الجديدة داخل المخططات الرسومية
function updateChartsDynamically(summary) {
    if (revenueChartInstance && expenseChartInstance) {
        revenueChartInstance.data.datasets[0].data = [summary.totalRev, summary.totalExp, summary.netProfit];
        revenueChartInstance.update();

        expenseChartInstance.data.datasets[0].data = [summary.assets, summary.liabilities, summary.equity];
        expenseChartInstance.update();
    }
}
