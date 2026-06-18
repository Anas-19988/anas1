// تشغيل العمليات والمراقبة فور اكتمال بناء واجهة الويب
document.addEventListener("DOMContentLoaded", () => {
    initNetworkMonitor();
    initCharts();
    loadSavedData();
    
    // ربط مستمع الحدث لالتقاط ميزان المراجعة المرفوع
    const uploader = document.getElementById('excel-uploader');
    if (uploader) {
        uploader.addEventListener('change', handleFileUpload);
    }
});

let revenueChartInstance = null;
let expenseChartInstance = null;

// 1. مراقبة حالة اتصال المستخدم بالإنترنت لمنع توقف المكتبات
function initNetworkMonitor() {
    const toast = document.getElementById('offline-toast');
    if (!toast) return;
    window.addEventListener('offline', () => toast.classList.remove('hidden'));
    window.addEventListener('online', () => toast.classList.add('hidden'));
    if (!navigator.onLine) toast.classList.remove('hidden');
}

// 2. تهيئة الرسوم البيانية بهياكل مالية فارغة حتى يتم رفع الملف
function initCharts() {
    const ctxRev = document.getElementById('revenueChart');
    if (ctxRev) {
        revenueChartInstance = new Chart(ctxRev.getContext('2d'), {
            type: 'bar',
            data: { 
                labels: ['إجمالي الإيرادات', 'إجمالي المصروفات', 'صافي الربح'], 
                datasets: [{ label: 'قائمة الدخل (ر.س)', data: [0, 0, 0], backgroundColor: ['#2563eb', '#e11d48', '#10b981'] }] 
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
                datasets: [{ data: [0, 0, 0], backgroundColor: ['#06b6d4', '#f59e0b', '#7c3aed'] }] 
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }
}

// 3. قراءة ملف ميزان المراجعة المرفوع وتحويله إلى كائن برمجى (JSON)
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
        
        // الانتقال الفوري لفرز وتصنيف ميزان المراجعة لتوليد القوائم
        processFinancialData(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

// 4. الفرز المالي الذكي لميزان المراجعة وتوليد القوائم والميزانية العمومية
function processFinancialData(data) {
    let totalRevenue = 0;     
    let totalExpenses = 0;    
    let totalAssets = 0;      
    let totalLiabilities = 0; 
    let totalEquity = 0;      

    data.forEach(row => {
        // فحص مسميات الأعمدة الشائعة في برامج المحاسبة وموازين المراجعة
        const accountName = (row['اسم الحساب'] || row['الحساب'] || row['Account'] || '').toString();
        const debit = parseFloat(row['الرصيد المدين'] || row['مدين'] || row['المدين'] || row['Debit'] || 0);
        const credit = parseFloat(row['الرصيد الدائن'] || row['دائن'] || row['الدائن'] || row['Credit'] || 0);

        // التصنيف الذكي بناءً على الهيكل المحاسبي للميزان
        if (accountName.includes('إيراد') || accountName.includes('مبيعات') || accountName.includes('Revenue') || accountName.includes('Sales')) {
            totalRevenue += credit; 
        } 
        else if (accountName.includes('مصروف') || accountName.includes('رواتب') || accountName.includes('إيجار') || accountName.includes('تكلفة') || accountName.includes('Expenses') || accountName.includes('Cost')) {
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

    // احتساب صافي الدخل (قائمة الدخل)
    const netProfit = totalRevenue - totalExpenses;

    // دمج وتثبيت أرباح العام الحالي داخل حقوق الملكية لتوازن الميزانية تماماً
    const finalEquity = totalEquity + netProfit;

    const financialSummary = {
        totalRev: totalRevenue,
        totalExp: totalExpenses,
        netProfit: netProfit,
        assets: totalAssets,
        liabilities: totalLiabilities,
        equity: finalEquity
    };

    // حفظ المخرجات في ذاكرة المتصفح للرجوع إليها لاحقاً
    localStorage.setItem('financialData', JSON.stringify(financialSummary));

    // تحديث الأرقام والرسوم البيانية المباشرة
    updateUI(financialSummary);
    updateChartsDynamically(financialSummary);
}

// 5. جلب وعرض البيانات المحفوظة تلقائياً عند إعادة زيارة الموقع
function loadSavedData() {
    const saved = localStorage.getItem('financialData');
    if (saved) {
        const data = JSON.parse(saved);
        updateUI(data);
        setTimeout(() => updateChartsDynamically(data), 200);
    }
}

// 6. تحديث الأرقام والنصوص في الواجهة والميزانية المتوازنة
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

// 7. ضخ البيانات الجديدة داخل المخططات الرسومية لإعادة رسمها ديناميكياً
function updateChartsDynamically(summary) {
    if (revenueChartInstance && expenseChartInstance) {
        revenueChartInstance.data.datasets[0].data = [summary.totalRev, summary.totalExp, summary.netProfit];
        revenueChartInstance.update();

        expenseChartInstance.data.datasets[0].data = [summary.assets, summary.liabilities, summary.equity];
        expenseChartInstance.update();
    }
}
