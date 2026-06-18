console.log("Enterprise Pro Max — Glass Ultra Loaded");

document.addEventListener("DOMContentLoaded", () => {
    initCharts();
    loadSavedData();

    document.getElementById("excel-uploader")
        .addEventListener("change", handleFileUpload);
});

let revenueChartInstance = null;
let expenseChartInstance = null;

/* ---------------------- الرسوم ---------------------- */

function initCharts() {
    const rev = document.getElementById("revenueChart");
    const exp = document.getElementById("expenseChart");

    revenueChartInstance = new Chart(rev, {
        type: "line",
        data: {
            labels: ["يناير", "فبراير", "مارس", "أبريل"],
            datasets: [{
                label: "الإيرادات",
                data: [0, 0, 0, 0],
                borderColor: "#ffffff",
                backgroundColor: "rgba(255,255,255,0.2)",
                borderWidth: 3,
                tension: 0.4
            }]
        },
        options: { responsive: true }
    });

    expenseChartInstance = new Chart(exp, {
        type: "doughnut",
        data: {
            labels: ["تشغيل", "تسويق", "أجور"],
            datasets: [{
                data: [0, 0, 0],
                backgroundColor: ["#ff4d6d", "#ffd166", "#6a4c93"]
            }]
        },
        options: { responsive: true }
    });
}

/* ---------------------- رفع الملف ---------------------- */

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheet = workbook.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);

        processFinancialData(json, file.name);
    };

    reader.readAsArrayBuffer(file);
}

/* ---------------------- معالجة البيانات ---------------------- */

function processFinancialData(data, fileName) {
    let totalRev = 0;
    let totalExp = 0;

    data.forEach(row => {
        totalRev += Number(row["Revenue"] || row["الإيرادات"] || 0);
        totalExp += Number(row["Expenses"] || row["المصروفات"] || 0);
    });

    const netProfit = totalRev - totalExp;

    const summary = {
        totalRev,
        totalExp,
        netProfit,
        date: new Date().toISOString(),
        fileName
    };

    localStorage.setItem("financialData", JSON.stringify(summary));

    updateUI(summary);
    updateCharts(summary);
    updateAnalysis(summary);
}

/* ---------------------- تحديث الواجهة ---------------------- */

function loadSavedData() {
    const saved = localStorage.getItem("financialData");
    if (!saved) return;

    const data = JSON.parse(saved);
    updateUI(data);
    updateCharts(data);
    updateAnalysis(data);
}

function updateUI(data) {
    document.getElementById("total-revenues").innerText =
        data.totalRev.toLocaleString("ar-SA") + " ر.س";

    document.getElementById("total-expenses").innerText =
        data.totalExp.toLocaleString("ar-SA") + " ر.س";

    document.getElementById("net-profit").innerText =
        data.netProfit.toLocaleString("ar-SA") + " ر.س";

    document.getElementById("profit-margin").innerText =
        ((data.netProfit / data.totalRev) * 100 || 0).toFixed(1) + "%";

    document.getElementById("last-update").innerText =
        "آخر تحديث: " + new Date(data.date).toLocaleString("ar-SA");
}

/* ---------------------- تحديث الرسوم ---------------------- */

function updateCharts(data) {
    revenueChartInstance.data.datasets[0].data = [
        data.totalRev * 0.2,
        data.totalRev * 0.3,
        data.totalRev * 0.25,
        data.totalRev * 0.25
    ];
    revenueChartInstance.update();

    expenseChartInstance.data.datasets[0].data = [
        data.totalExp * 0.5,
        data.totalExp * 0.3,
        data.totalExp * 0.2
    ];
    expenseChartInstance.update();
}

/* ---------------------- التحليل المالي ---------------------- */

function updateAnalysis(data) {
    const expRatio = (data.totalExp / data.totalRev) * 100 || 0;
    const netRatio = (data.netProfit / data.totalRev) * 100 || 0;

    document.getElementById("expense-ratio").innerText = expRatio.toFixed(1) + "%";
    document.getElementById("net-profit-ratio").innerText = netRatio.toFixed(1) + "%";
    document.getElementById("break-even-status").innerText =
        data.netProfit >= 0 ? "متحقق" : "غير متحقق";

    let health = "غير متوفر";

    if (netRatio > 25) health = "ممتاز";
    else if (netRatio > 10) health = "جيد";
    else if (netRatio >= 0) health = "مقبول";
    else health = "حرج";

    document.getElementById("financial-health").innerText = health;

    const list = document.getElementById("analysis-summary");
    list.innerHTML = `
        <li>إجمالي الإيرادات: ${data.totalRev.toLocaleString("ar-SA")} ر.س</li>
        <li>إجمالي المصروفات: ${data.totalExp.toLocaleString("ar-SA")} ر.س</li>
        <li>صافي الربح: ${data.netProfit.toLocaleString("ar-SA")} ر.س</li>
        <li>نسبة المصروفات: ${expRatio.toFixed(1)}%</li>
        <li>نسبة الربحية: ${netRatio.toFixed(1)}%</li>
    `;
}
