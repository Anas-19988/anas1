// 1. تنظيف الرابط تلقائياً وحذف علامات الهاش العالقة من الكاش لمنع خطأ الـ 404
if (window.location.hash) {
    window.history.replaceState("", document.title, window.location.pathname);
}

// تشغيل العمليات فور اكتمال بناء واجهة الويب
document.addEventListener("DOMContentLoaded", () => {
    initNetworkMonitor();
    initCharts();
    initAuthSystem();
    checkLoginStatus(); // التحقق إذا كان المستخدم مسجلاً لدخوله مسبقاً
    
    // ربط مستمع الحدث لالتقاط ميزان المراجعة المرفوع
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

    // فتح شباك تسجيل الدخول
    if (openLoginBtn) openLoginBtn.addEventListener('click', () => loginModal.classList.remove('hidden'));
    if (heroActionBtn) heroActionBtn.addEventListener('click', () => loginModal.classList.remove('hidden'));
    
    // إغلاق شباك تسجيل الدخول
    if (closeLoginBtn) closeLoginBtn.addEventListener('click', () => loginModal.classList.add('hidden'));

    // معالجة تقديم نموذج الدخول
    if (loginForm) {
        loginForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const user = document.getElementById('username').value.trim().toLowerCase();
            const pass = document.getElementById('password').value;
            const errorEl = document.getElementById('login-error');

            // يمكنك تغيير بيانات الدخول هنا (تم تعيين اسم المستخدم: anas)
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

    // تسجيل الخروج والعودة للصفحة التعريفية
    if (logoutBtn) {
        logoutBtn.addEventListener('click', () => {
            localStorage.removeItem('isLoggedIn');
            localStorage.removeItem('sessionUser');
            checkLoginStatus();
        });
    }
}

// فحص حالة الجلسة وتغيير محتوى الواجهة فوراً
function checkLoginStatus() {
    const isLoggedIn = localStorage.getItem('isLoggedIn') === 'true';
    const landingPage = document.getElementById('landing-page');
    const dashboardPage = document.getElementById('dashboard-page');
    const userDisplay = document.getElementById('user-display-name');

    if (isLoggedIn) {
        if (landingPage) landingPage.classList.add('hidden');
        if (dashboardPage) dashboardPage.classList.remove('hidden');
        if (userDisplay) userDisplay.innerText = `مرحباً، ${localStorage.getItem('sessionUser')}`;
        loadSavedData(); // جلب أي بيانات إكسيل تم رفعها سابقاً لتظل ثابتة
    } else {
        if (dashboardPage) dashboardPage.classList.add('hidden');
        if (landingPage) landingPage.classList.remove('hidden');
    }
}

// 3. مراقبة حالة اتصال المستخدم بالإنترنت لمنع توقف المكتبات الخارجية
function initNetworkMonitor() {
    const toast = document.getElementById('offline-toast');
    if (!toast) return;
    window.addEventListener('offline', () => toast.classList.remove('hidden'));
    window.addEventListener('online', () => toast.classList.add('hidden'));
    if (!navigator.onLine) toast.classList.remove('hidden');
}

// 4. تهيئة الرسوم البيانية بهياكل مالية فارغة حتى يتم رفع الملف
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

// 5. قراءة ملف ميزان المراجعة المرفوع وتحويله إلى كائن برمجى (JSON)
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        processFinancialData(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

// 6. الفرز المالي الذكي لميزان المراجعة وتوليد القوائم والميزانية العمومية
function processFinancialData(data) {
    let totalRevenue = 0;     
    let totalExpenses = 0;    
    let totalAssets = 0;      
    let totalLiabilities = 0; 
    let totalEquity = 0;      

    data.forEach(row => {
        const accountName = (row['اسم الحساب'] || row['الحساب'] || row['اسم_الحساب'] || row['Account'] || row['Account Name'] || '').toString();
        const debit = parseFloat(row['الرصيد المدين'] || row['مدين'] || row['المدين'] || row['Debit'] || row['debit'] || 0);
        const credit = parseFloat(row['الرصيد الدائن'] || row['دائن'] || row['الدائن'] || row['Credit'] || row['credit'] || 0);

        if (accountName.includes('إيراد') || accountName.includes('مبيعات') || accountName.includes('Revenue') || accountName.includes('Sales')) {
            totalRevenue += credit; 
        } 
        else if (accountName.includes('مصروف') || accountName.includes('رواتب') || accountName.includes('إيجار') || accountName.includes('تكلفة') || accountName.includes('تكاليف') || accountName.includes('Expenses') || accountName.includes('Cost')) {
            totalExpenses += debit; 
        }
        else if (accountName.includes('نقدية') || accountName.includes('بنك') || accountName.includes('عملاء') || accountName.includes('أصول') || accountName.includes('ثابتة') || accountName.includes('Assets') || accountName.includes('Cash')) {
            totalAssets += (debit - credit); 
        }
        else if (accountName.includes('موردين') || accountName.includes('قروض') || accountName.includes('التزامات') || accountName.includes('خصوم') || accountName.includes('Liabilities') || accountName.includes('Payable')) {
            totalLiabilities += (credit - debit); 
        }
        else if (accountName.includes('رأس المال') || accountName.includes('ملكية') || accountName.includes('حقوق') || accountName.includes('Equity') || accountName.includes('Capital')) {
            totalEquity += (credit - debit); 
        }
    });

    const netProfit = totalRevenue - totalExpenses;
    const finalEquity = totalEquity + netProfit;

    const financialSummary = {
        totalRev: totalRevenue,
        totalExp: totalExpenses,
        netProfit: netProfit,
        assets: totalAssets,
        liabilities: totalLiabilities,
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

// 8. تحديث الأرقام والنصوص في الواجهة والميزانية المتوازنة
function updateUI(data) {
    const revEl = document.getElementById('total-revenues');
    const expEl = document.getElementById('total-expenses');
    const profEl = document.getElementById('net-profit');
    const assetEl = document.getElementById('balance-assets');
    const eqLiabEl = document.getElementById('balance-eq-liab');

    if (revEl) revEl.innerText = data.totalRev.toLocaleString('ar-SA') + ' ر.س';
    if (expEl) expEl.innerText = data.totalExp.toLocaleString('ar-SA') + ' ر.س';
    if (profEl) profEl.innerText = data.netProfit.toLocaleString('ar-SA') + ' ر.س';
    if (assetEl) assetEl.innerText = data.assets.toLocaleString('ar-SA') + ' ر.س';
    if (eqLiabEl) eqLiabEl.innerText = (data.liabilities + data.equity).toLocaleString('ar-SA') + ' ر.س';
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
