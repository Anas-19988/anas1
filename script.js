// 1. تنظيف الرابط تلقائياً وحذف علامات الهاش العالقة من الكاش لمنع خطأ الـ 404
if (window.location.hash) {
    window.history.replaceState("", document.title, window.location.pathname);
}

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

// 2. مراقبة حالة اتصال المستخدم بالإنترنت لمنع توقف المكتبات الخارجية
function initNetworkMonitor() {
    const toast = document.getElementById('offline-toast');
    if (!toast) return;
    window.addEventListener('offline', () => toast.classList.remove('hidden'));
    window.addEventListener('online', () => toast.classList.add('hidden'));
    if (!navigator.onLine) toast.classList.remove('hidden');
}

// 3. تهيئة الرسوم البيانية بهياكل مالية فارغة حتى يتم رفع الملف
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

// 4. قراءة ملف ميزان المراجعة المرفوع وتحويله إلى كائن برمجى (JSON)
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

// 5. الفرز المالي الذكي لميزان المراجعة وتوليد القوائم والميزانية العمومية
function processFinancialData(data) {
    let totalRevenue = 0;     
    let totalExpenses = 0;    
    let totalAssets = 0;      
    let totalLiabilities = 0; 
    let totalEquity = 0;      

    data.forEach(row => {
        // فحص مسميات الأعمدة الشائعة في برامج المحاسبة وموازين المراجعة (عربي وإنجليزي)
