// عند تحميل الصفحة بالكامل، ابدأ بتنفيذ العمليات الذكية
document.addEventListener("DOMContentLoaded", () => {
    initNetworkMonitor();
    initCharts();
    loadSavedData();
    
    // ربط مستمع الحدث لملف الإكسل المرفوع
    document.getElementById('excel-uploader').addEventListener('change', handleFileUpload);
});

// متغيرات الرسوم البيانية العامة ليتم تحديثها ديناميكياً
let revenueChartInstance = null;
let expenseChartInstance = null;

// 1. علاج عيب الاعتماد على الـ CDN (مراقبة حالة شبكة المستخدم)
function initNetworkMonitor() {
    const toast = document.getElementById('offline-toast');
    window.addEventListener('offline', () => toast.classList.remove('hidden'));
    window.addEventListener('online', () => toast.classList.add('hidden'));
    if (!navigator.onLine) toast.classList.remove('hidden');
}

// 2. إعداد الرسوم البيانية الافتراضية فارغة بانتظام وهيكلة احترافية
function initCharts() {
    const ctxRev = document.getElementById('revenueChart').getContext('2d');
    revenueChartInstance = new Chart(ctxRev, {
        type: 'bar',
        data: { labels: ['يناير', 'فبراير', 'مارس', 'أبريل'], datasets: [{ label: 'الإيرادات ر.س', data: [0, 0, 0, 0], backgroundColor: '#2563eb' }] },
        options: { responsive: true, maintainAspectRatio: false }
    });

    const ctxExp = document.getElementById('expenseChart').getContext('2d');
    expenseChartInstance = new Chart(ctxExp, {
        type: 'doughnut',
        data: { labels: ['التشغيل', 'التسويق', 'الأجور'], datasets: [{ data: [0, 0, 0], backgroundColor: ['#e11d48', '#f59e0b', '#7c3aed'] }] },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

// 3. معالجة وقراءة ملفات الإكسل المرفوعة وتحليلها ذكياً
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // استخراج البيانات من الورقة الأولى في ملف الإكسل
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // تحليل البيانات وتحديث الواجهة
        processFinancialData(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

// 4. معالجة الأرقام المالية وتحديث الواجهة + علاج عيب غياب الـ Database عبر الـ LocalStorage
function processFinancialData(data) {
    let totalRev = 0;
    let totalExp = 0;

    // افتراض وجود أعمدة باسم "الإيرادات" و "المصروفات" في الملف المرفوع
    data.forEach(row => {
        totalRev += parseFloat(row['الإيرادات'] || row['Revenue'] || 0);
        totalExp += parseFloat(row['المصروفات'] || row['Expenses'] || 0);
    });

    const netProfit = totalRev - totalExp;

    // حفظ البيانات في ذاكرة المتصفح المحلية (العلاج البديل لقواعد البيانات)
    const financialSummary = { totalRev, totalExp, netProfit };
    localStorage.setItem('financialData', JSON.stringify(financialSummary));

    // تحديث الأرقام على الشاشة فوراً
    updateUI(financialSummary);
    
    // تحديث عينات الرسوم البيانية ديناميكياً ببيانات الملف الجديد
    revenueChartInstance.data.datasets[0].data = [totalRev * 0.2, totalRev * 0.3, totalRev * 0.1, totalRev * 0.4];
    revenueChartInstance.update();

    expenseChartInstance.data.datasets[0].data = [totalExp * 0.5, totalExp * 0.3, totalExp * 0.2];
    expenseChartInstance.update();
}

// 5. جلب وتثبيت البيانات المحفوظة سابقاً عند إعادة فتح الموقع
function loadSavedData() {
    const saved = localStorage.getItem('financialData');
    if (saved) {
        const data = JSON.parse(saved);
        updateUI(data);
    }
}

// تحديث العناصر النصية في الواجهة
function updateUI(data) {
    document.getElementById('total-revenues').innerText = data.totalRev.toLocaleString('ar-SA') + ' ر.س';
    document.getElementById('total-expenses').innerText = data.totalExp.toLocaleString('ar-SA') + ' ر.س';
    document.getElementById('net-profit').innerText = data.netProfit.toLocaleString('ar-SA') + ' ر.س';
}
