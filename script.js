// عند تحميل الصفحة بالكامل، ابدأ بتنفيذ العمليات الذكية
document.addEventListener("DOMContentLoaded", () => {
    initNetworkMonitor();
    initCharts();
    loadSavedData();
    
    const uploader = document.getElementById('excel-uploader');
    if (uploader) {
        uploader.addEventListener('change', handleFileUpload);
    }
});

// متغيرات الرسوم البيانية العامة ليتم تحديثها ديناميكياً
let revenueChartInstance = null;
let expenseChartInstance = null;

// 1. مراقبة حالة شبكة المستخدم
function initNetworkMonitor() {
    const toast = document.getElementById('offline-toast');
    if (!toast) return;

    const toggleToast = () => {
        if (!navigator.onLine) toast.classList.remove('hidden');
        else toast.classList.add('hidden');
    };

    window.addEventListener('offline', toggleToast);
    window.addEventListener('online', toggleToast);
    toggleToast();
}

// 2. إعداد الرسوم البيانية الافتراضية
function initCharts() {
    const revCanvas = document.getElementById('revenueChart');
    const expCanvas = document.getElementById('expenseChart');
    if (!revCanvas || !expCanvas) return;

    const ctxRev = revCanvas.getContext('2d');
    revenueChartInstance = new Chart(ctxRev, {
        type: 'bar',
        data: {
            labels: ['يناير', 'فبراير', 'مارس', 'أبريل'],
            datasets: [{
                label: 'الإيرادات ر.س',
                data: [0, 0, 0, 0],
                backgroundColor: '#2563eb'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: true },
                tooltip: { enabled: true }
            },
            scales: {
                y: { beginAtZero: true }
            }
        }
    });

    const ctxExp = expCanvas.getContext('2d');
    expenseChartInstance = new Chart(ctxExp, {
        type: 'doughnut',
        data: {
            labels: ['التشغيل', 'التسويق', 'الأجور'],
            datasets: [{
                data: [0, 0, 0],
                backgroundColor: ['#e11d48', '#f59e0b', '#7c3aed']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' }
            }
        }
    });
}

// 3. معالجة وقراءة ملفات الإكسل المرفوعة
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
        showError("الرجاء رفع ملف Excel أو CSV فقط.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            if (!jsonData || jsonData.length === 0) {
                showError("الملف لا يحتوي على بيانات قابلة للقراءة.");
                return;
            }

            processFinancialData(jsonData, file.name);

        } catch (err) {
            console.error(err);
            showError("حدث خطأ أثناء قراءة الملف. تأكد أن الملف سليم.");
        }
    };

    reader.readAsArrayBuffer(file);
}

// 4. معالجة الأرقام المالية وتحديث الواجهة + LocalStorage
function processFinancialData(data, fileName = 'غير معروف') {
    let totalRev = 0;
    let totalExp = 0;

    const revenueKeys = ['الإيرادات', 'ايرادات', 'Revenue', 'Revenues', 'Sales', 'Income'];
    const expenseKeys = ['المصروفات', 'مصروفات', 'Expenses', 'Expense', 'Cost', 'Costs'];

    data.forEach(row => {
        let rev = 0;
        let exp = 0;

        for (let key of revenueKeys) {
            if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
                const val = parseFloat(row[key]);
                if (!isNaN(val)) {
                    rev = val;
                    break;
                }
            }
        }

        for (let key of expenseKeys) {
            if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
                const val = parseFloat(row[key]);
                if (!isNaN(val)) {
                    exp = val;
                    break;
                }
            }
        }

        totalRev += rev;
        totalExp += exp;
    });

    const netProfit = totalRev - totalExp;

    const financialSummary = {
        totalRev,
        totalExp,
        netProfit,
        fileName,
        date: new Date().toISOString()
    };

    localStorage.setItem('financialData', JSON.stringify(financialSummary));

    updateUI(financialSummary);
    updateCharts(totalRev, totalExp);
    updateFinancialAnalysis(financialSummary);
}

// 5. جلب وتثبيت البيانات المحفوظة سابقاً
function loadSavedData() {
    const saved = localStorage.getItem('financialData');
    if (saved) {
        const data = JSON.parse(saved);
        updateUI(data);
        updateCharts(data.totalRev || 0, data.totalExp || 0);
        updateFinancialAnalysis(data);
    }
}

// تحديث العناصر النصية في الواجهة
function updateUI(data) {
    const totalRevEl = document.getElementById('total-revenues');
    const totalExpEl = document.getElementById('total-expenses');
    const netProfitEl = document.getElementById('net-profit');
    const profitMarginEl = document.getElementById('profit-margin');
    const lastUpdateEl = document.getElementById('last-update');

    const totalRev = data.totalRev || 0;
    const totalExp = data.totalExp || 0;
    const netProfit = data.netProfit || 0;

    if (totalRevEl) totalRevEl.innerText = totalRev.toLocaleString('ar-SA') + ' ر.س';
    if (totalExpEl) totalExpEl.innerText = totalExp.toLocaleString('ar-SA') + ' ر.س';
    if (netProfitEl) netProfitEl.innerText = netProfit.toLocaleString('ar-SA') + ' ر.س';

    const profitMargin = totalRev > 0 ? (netProfit / totalRev) * 100 : 0;
    if (profitMarginEl) profitMarginEl.innerText = profitMargin.toFixed(1) + '٪';

    if (lastUpdateEl && data.date) {
        const d = new Date(data.date);
        lastUpdateEl.innerText = `آخر تحديث: ${d.toLocaleString('ar-SA')} — ملف: ${data.fileName || 'غير معروف'}`;
    }
}

// تحديث الرسوم البيانية
function updateCharts(totalRev, totalExp) {
    if (revenueChartInstance) {
        const revData = [
            totalRev * 0.25,
            totalRev * 0.15,
            totalRev * 0.35,
            totalRev * 0.25
        ];
        revenueChartInstance.data.datasets[0].data = revData;
        revenueChartInstance.update();
    }

    if (expenseChartInstance) {
        const expData = [
            totalExp * 0.45,
            totalExp * 0.30,
            totalExp * 0.25
        ];
        expenseChartInstance.data.datasets[0].data = expData;
        expenseChartInstance.update();
    }
}

// 6. نظام التحليل المالي الكامل
function updateFinancialAnalysis(summary) {
    const totalRev = summary.totalRev || 0;
    const totalExp = summary.totalExp || 0;
    const netProfit = summary.netProfit || 0;

    const expenseRatioEl = document.getElementById('expense-ratio');
    const netProfitRatioEl = document.getElementById('net-profit-ratio');
    const breakEvenEl = document.getElementById('break-even-status');
    const healthEl = document.getElementById('financial-health');
    const analysisList = document.getElementById('analysis-summary');

    const expenseRatio = totalRev > 0 ? (totalExp / totalRev) * 100 : 0;
    const netProfitRatio = totalRev > 0 ? (netProfit / totalRev) * 100 : 0;
    const breakEven = netProfit >= 0;

    if (expenseRatioEl) expenseRatioEl.innerText = expenseRatio.toFixed(1) + '٪';
    if (netProfitRatioEl) netProfitRatioEl.innerText = netProfitRatio.toFixed(1) + '٪';
    if (breakEvenEl) breakEvenEl.innerText = breakEven ? 'متحقق' : 'غير متحقق';

    let healthText = 'غير متوفر';
    if (totalRev === 0 && totalExp === 0) {
        healthText = 'لا توجد بيانات كافية لتقييم الحالة المالية.';
    } else if (netProfitRatio > 25 && expenseRatio < 60) {
        healthText = 'الحالة المالية ممتازة مع ربحية عالية ومصروفات تحت السيطرة.';
    } else if (netProfitRatio > 10 && expenseRatio < 80) {
        healthText = 'الحالة المالية جيدة لكن يمكن تحسين كفاءة المصروفات.';
    } else if (netProfitRatio >= 0) {
        healthText = 'الحالة المالية مستقرة لكن الربحية منخفضة وتحتاج مراجعة.';
    } else {
        healthText = 'الحالة المالية حرجة مع خسائر قائمة، يلزم اتخاذ إجراءات تصحيحية.';
    }

    if (healthEl) healthEl.innerText = healthText;

    if (analysisList) {
        analysisList.innerHTML = '';

        const points = [];

        points.push(`إجمالي الإيرادات: ${totalRev.toLocaleString('ar-SA')} ر.س`);
        points.push(`إجمالي المصروفات: ${totalExp.toLocaleString('ar-SA')} ر.س`);
        points.push(`صافي الربح: ${netProfit.toLocaleString('ar-SA')} ر.س`);
        points.push(`نسبة المصروفات إلى الإيرادات: ${expenseRatio.toFixed(1)}٪`);
        points.push(`نسبة صافي الربح: ${netProfitRatio.toFixed(1)}٪`);
        points.push(`نقطة التعادل: ${breakEven ? 'تم تجاوز نقطة التعادل (المنشأة رابحة).' : 'لم يتم تجاوز نقطة التعادل (المنشأة في حالة خسارة أو تعادل).'}`);

        points.forEach(text => {
            const li = document.createElement('li');
            li.innerText = text;
            analysisList.appendChild(li);
        });
    }
}

// 7. نظام تنبيه الأخطاء
function showError(message) {
    const toast = document.getElementById('error-toast');
    const msgEl = document.getElementById('error-message');
    if (!toast || !msgEl) return;

    msgEl.innerText = message;
    toast.classList.remove('hidden');

    setTimeout(() => {
        toast.classList.add('hidden');
    }, 5000);
}
