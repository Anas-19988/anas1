// الدالة المحدثة لقراءة ميزان المراجعة وتوليد القوائم المالية ديناميكياً
function processFinancialData(data) {
    let totalRevenue = 0;     // الإيرادات / المبيعات
    let totalExpenses = 0;    // المصروفات (إدارية، عمومية، تشغيلية)
    let totalAssets = 0;      // الأصول (متداولة وثابتة)
    let totalLiabilities = 0; // الالتزامات / الخصوم
    let totalEquity = 0;      // حقوق الملكية (رأس المال، الأرباح المبقاة)

    data.forEach(row => {
        // جلب اسم الحساب وقيم المدين والدائن مع تحويلها لأرقام
        const accountName = (row['اسم الحساب'] || row['الحساب'] || row['Account'] || '').toString();
        const debit = parseFloat(row['الرصيد المدين'] || row['مدين'] || row['Debit'] || 0);
        const credit = parseFloat(row['الرصيد الدائن'] || row['دائن'] || row['Credit'] || 0);

        // الفرز الذكي بناءً على الكلمات المفتاحية في ميزان المراجعة
        if (accountName.includes('إيراد') || accountName.includes('مبيعات') || accountName.includes('Revenue') || accountName.includes('Sales')) {
            totalRevenue += credit; // الإيرادات طبيعتها دائنة
        } 
        else if (accountName.includes('مصروف') || accountName.includes('رواتب') || accountName.includes('إيجار') || accountName.includes('Expenses') || accountName.includes('Cost')) {
            totalExpenses += debit; // المصروفات طبيعتها مدينة
        }
        else if (accountName.includes('نقدية') || accountName.includes('بنك') || accountName.includes('عملاء') || accountName.includes('أصول') || accountName.includes('Assets') || accountName.includes('Cash')) {
            totalAssets += (debit - credit); // الأصول صافي المدين
        }
        else if (accountName.includes('موردين') || accountName.includes('قروض') || accountName.includes('التزامات') || accountName.includes('Liabilities') || accountName.includes('Payable')) {
            totalLiabilities += (credit - debit); // الالتزامات صافي الدائن
        }
        else if (accountName.includes('رأس المال') || accountName.includes('ملكية') || accountName.includes('Equity') || accountName.includes('Capital')) {
            totalEquity += (credit - debit); // حقوق الملكية صافي الدائن
        }
    });

    // حساب صافي الربح أو الخسارة لقائمة الدخل
    const netProfit = totalRevenue - totalExpenses;

    // ترحيل صافي الربح إلى حقوق الملكية لتوازن الميزانية العمومية
    const finalEquity = totalEquity + netProfit;

    // تجميع البيانات في كائن واحد لحفظه وتحديث الواجهة
    const financialSummary = {
        totalRev: totalRevenue,
        totalExp: totalExpenses,
        netProfit: netProfit,
        assets: totalAssets,
        liabilities: totalLiabilities,
        equity: finalEquity
    };

    // حفظ القوائم المالية في ذاكرة المتصفح (حل عيب غياب قاعدة البيانات)
    localStorage.setItem('financialData', JSON.stringify(financialSummary));

    // تحديث الأرقام والرسوم البيانية في الواجهة فوراً
    updateUI(financialSummary);
    updateChartsDynamically(financialSummary);
}

// دالة لتحديث الرسوم البيانية بناءً على التصنيفات الجديدة لميزان المراجعة
function updateChartsDynamically(summary) {
    // 1. تحديث مخطط قائمة الدخل (إيرادات ومصروفات)
    revenueChartInstance.data.labels = ['إجمالي الإيرادات', 'إجمالي المصروفات', 'صافي الربح'];
    revenueChartInstance.data.datasets[0].data = [summary.totalRev, summary.totalExp, summary.netProfit];
    revenueChartInstance.data.datasets[0].backgroundColor = ['#2563eb', '#e11d48', '#10b981'];
    revenueChartInstance.update();

    // 2. تحديث مخطط الميزانية العمومية (أصول ضد خصوم وحقوق ملكية)
    expenseChartInstance.data.labels = ['الأصول', 'الالتزامات', 'حقوق الملكية'];
    expenseChartInstance.data.datasets[0].data = [summary.assets, summary.liabilities, summary.equity];
    expenseChartInstance.data.datasets[0].backgroundColor = ['#06b6d4', '#f59e0b', '#7c3aed'];
    expenseChartInstance.update();
}
