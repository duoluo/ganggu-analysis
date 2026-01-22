// 港股打新收益分析系统 - 主脚本
let reportData = null;
let currentDetailAccount = null;

// 页面加载时初始化
document.addEventListener('DOMContentLoaded', async function() {
    try {
        // 加载数据
        const response = await fetch('report_data.json');
        reportData = await response.json();

        // 初始化页面
        initializePage();
    } catch (error) {
        console.error('加载数据失败:', error);
        alert('数据加载失败，请确保 report_data.json 文件存在！');
    }
});

// 初始化页面
function initializePage() {
    // 更新统计卡片
    updateStatsCards();

    // 更新时间
    document.getElementById('updateTime').textContent = `数据更新时间: ${reportData.generated_at}`;

    // 初始化图表
    initializeCharts();

    // 初始化表格
    updateAccountRevenueTable();
    updateCommissionSummaryTable();
    updateSpecialRangeTable();
    updateMissingRecords();

    // 绑定筛选器事件
    document.getElementById('accountGroupFilter').addEventListener('change', updateAccountRevenueTable);
    document.getElementById('accountSortBy').addEventListener('change', updateAccountRevenueTable);
    document.getElementById('commissionGroupFilter').addEventListener('change', updateCommissionSummaryTable);
}

// 更新统计卡片
function updateStatsCards() {
    const summary = reportData.summary;
    const html = `
        <div class="stat-card success">
            <div class="label">总收益</div>
            <div class="value">¥${formatNumber(summary.total_revenue)}</div>
        </div>
        <div class="stat-card info">
            <div class="label">总分成</div>
            <div class="value">¥${formatNumber(summary.total_commission)}</div>
        </div>
        <div class="stat-card danger">
            <div class="label">总亏损</div>
            <div class="value">¥${formatNumber(summary.total_loss)}</div>
        </div>
        <div class="stat-card warning">
            <div class="label">股票数量</div>
            <div class="value">${summary.total_stocks}</div>
        </div>
        <div class="stat-card info">
            <div class="label">账户数量</div>
            <div class="value">${summary.total_accounts}</div>
        </div>
    `;
    document.getElementById('statsCards').innerHTML = html;
}

// 初始化图表
function initializeCharts() {
    // 股票收益饼图
    const pieCtx = document.getElementById('stockPieChart').getContext('2d');

    // 只显示收益前10的股票，其他合并为"其他"
    const sortedStocks = [...reportData.stocks].sort((a, b) => b.revenue - a.revenue);
    const top10 = sortedStocks.slice(0, 10);
    const others = sortedStocks.slice(10);
    const othersSum = others.reduce((sum, s) => sum + s.revenue, 0);

    const pieLabels = top10.map(s => s.name);
    const pieData = top10.map(s => s.revenue);

    if (othersSum > 0) {
        pieLabels.push('其他');
        pieData.push(othersSum);
    }

    new Chart(pieCtx, {
        type: 'pie',
        data: {
            labels: pieLabels,
            datasets: [{
                data: pieData,
                backgroundColor: [
                    '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF',
                    '#FF9F40', '#FF6384', '#C9CBCF', '#4BC0C0', '#FF6384', '#36A2EB'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        boxWidth: 12,
                        font: { size: 10 }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.parsed;
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((value / total) * 100).toFixed(1);
                            return `${context.label}: ¥${formatNumber(value)} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });

    // 股票收益柱状图（前15名）
    const barCtx = document.getElementById('stockBarChart').getContext('2d');
    const top15 = sortedStocks.slice(0, 15);

    new Chart(barCtx, {
        type: 'bar',
        data: {
            labels: top15.map(s => s.name.length > 15 ? s.name.substring(0, 15) + '...' : s.name),
            datasets: [{
                label: '收益（元）',
                data: top15.map(s => s.revenue),
                backgroundColor: top15.map(s => s.revenue >= 0 ? '#27ae60' : '#e74c3c')
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `收益: ¥${formatNumber(context.parsed.y)}`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return '¥' + formatNumber(value);
                        }
                    }
                },
                x: {
                    ticks: {
                        font: { size: 10 }
                    }
                }
            }
        }
    });

    // 亏损股票分析
    const lossStocks = reportData.stocks.filter(s => s.revenue < 0);

    if (lossStocks.length > 0) {
        // 亏损股票饼图
        const lossPieCtx = document.getElementById('lossStockPieChart').getContext('2d');
        const sortedLossStocks = [...lossStocks].sort((a, b) => a.revenue - b.revenue); // 从小到大排序（最亏的在前）
        const top10Loss = sortedLossStocks.slice(0, 10);
        const othersLoss = sortedLossStocks.slice(10);
        const othersLossSum = othersLoss.reduce((sum, s) => sum + Math.abs(s.revenue), 0);

        const lossPieLabels = top10Loss.map(s => s.name);
        const lossPieData = top10Loss.map(s => Math.abs(s.revenue));

        if (othersLossSum > 0) {
            lossPieLabels.push('其他');
            lossPieData.push(othersLossSum);
        }

        new Chart(lossPieCtx, {
            type: 'pie',
            data: {
                labels: lossPieLabels,
                datasets: [{
                    data: lossPieData,
                    backgroundColor: [
                        '#e74c3c', '#c0392b', '#e67e22', '#d35400', '#f39c12',
                        '#f1c40f', '#e8b4b8', '#d98880', '#cd6155', '#c0392b', '#922b21'
                    ]
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right',
                        labels: {
                            boxWidth: 12,
                            font: { size: 10 }
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const value = context.parsed;
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = ((value / total) * 100).toFixed(1);
                                return `${context.label}: ¥${formatNumber(value)} (${percentage}%)`;
                            }
                        }
                    }
                }
            }
        });

        // 亏损股票柱状图
        const lossBarCtx = document.getElementById('lossStockBarChart').getContext('2d');
        const top15Loss = sortedLossStocks.slice(0, 15);

        new Chart(lossBarCtx, {
            type: 'bar',
            data: {
                labels: top15Loss.map(s => s.name.length > 15 ? s.name.substring(0, 15) + '...' : s.name),
                datasets: [{
                    label: '亏损（元）',
                    data: top15Loss.map(s => Math.abs(s.revenue)),
                    backgroundColor: '#e74c3c'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return `亏损: ¥${formatNumber(context.parsed.y)}`;
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return '¥' + formatNumber(value);
                            }
                        }
                    },
                    x: {
                        ticks: {
                            font: { size: 10 }
                        }
                    }
                }
            }
        });
    } else {
        // 如果没有亏损股票，显示提示信息
        document.getElementById('lossStockPieChart').parentElement.innerHTML =
            '<div class="alert alert-success text-center">🎉 没有亏损股票！</div>';
        document.getElementById('lossStockBarChart').parentElement.innerHTML =
            '<div class="alert alert-success text-center">🎉 没有亏损股票！</div>';
    }
}

// 更新账户收益表格
function updateAccountRevenueTable() {
    const groupFilter = document.getElementById('accountGroupFilter').value;
    const sortBy = document.getElementById('accountSortBy').value;

    let accounts = [...reportData.accounts];

    // 筛选
    if (groupFilter !== 'all') {
        accounts = accounts.filter(a => a.management_group === groupFilter);
    }

    // 排序
    accounts.sort((a, b) => {
        if (sortBy === 'revenue') return b.total_revenue - a.total_revenue;
        if (sortBy === 'commission') return b.total_commission - a.total_commission;
        if (sortBy === 'loss') return b.total_loss - a.total_loss;
        return 0;
    });

    // 计算筛选后的统计
    const filteredRevenue = accounts.reduce((sum, a) => sum + a.total_revenue, 0);
    const filteredCommission = accounts.reduce((sum, a) => sum + a.total_commission, 0);
    document.getElementById('filteredStats').innerHTML =
        `总收益: ¥${formatNumber(filteredRevenue)} | 分成: ¥${formatNumber(filteredCommission)}`;

    const tbody = document.getElementById('accountRevenueBody');
    tbody.innerHTML = accounts.map((acc, index) => `
        <tr>
            <td>${index + 1}</td>
            <td><strong>${acc.account}</strong></td>
            <td><span class="badge-rate badge-rate-${acc.rate_group.replace('%', '')}">${acc.rate_group}</span></td>
            <td class="${acc.total_revenue >= 0 ? 'positive' : 'negative'}">¥${formatNumber(acc.total_revenue)}</td>
            <td class="positive">¥${formatNumber(acc.total_commission)}</td>
            <td class="negative">¥${formatNumber(acc.total_loss)}</td>
        </tr>
    `).join('');
}

// 更新分成汇总表格
function updateCommissionSummaryTable() {
    const groupFilter = document.getElementById('commissionGroupFilter').value;

    let accounts = [...reportData.accounts];

    // 筛选
    if (groupFilter !== 'all') {
        accounts = accounts.filter(a => a.management_group === groupFilter);
    }

    // 按总分成排序
    accounts.sort((a, b) => b.total_commission - a.total_commission);

    const tbody = document.getElementById('commissionSummaryBody');
    tbody.innerHTML = accounts.map(acc => `
        <tr>
            <td><strong>${acc.account}</strong></td>
            <td><span class="badge-rate badge-rate-${acc.rate_group.replace('%', '')}">${acc.rate_group}</span></td>
            <td class="${acc.total_revenue >= 0 ? 'positive' : 'negative'}">¥${formatNumber(acc.total_revenue)}</td>
            <td class="positive">¥${formatNumber(acc.total_commission)}</td>
            <td class="negative">¥${formatNumber(acc.total_loss)}</td>
            <td>
                <button class="btn btn-sm btn-primary" onclick="showAccountDetail('${acc.account}')">
                    查看详情
                </button>
            </td>
        </tr>
    `).join('');
}

// 更新特定范围分成表格
function updateSpecialRangeTable() {
    const data = [...reportData.special_range];
    data.sort((a, b) => b.range_commission - a.range_commission);

    // 计算特定范围统计
    const totalRangeRevenue = data.reduce((sum, item) => sum + item.range_revenue, 0);
    const totalRangeCommission = data.reduce((sum, item) => sum + item.range_commission, 0);
    document.getElementById('specialRangeStats').innerHTML =
        `总收益: ¥${formatNumber(totalRangeRevenue)} | 分成总收益: ¥${formatNumber(totalRangeCommission)}`;

    const tbody = document.getElementById('specialRangeBody');
    tbody.innerHTML = data.map(item => `
        <tr>
            <td><strong>${item.account}</strong></td>
            <td><span class="badge-rate badge-rate-${item.rate_group.replace('%', '')}">${item.rate_group}</span></td>
            <td class="${item.range_revenue >= 0 ? 'positive' : 'negative'}">¥${formatNumber(item.range_revenue)}</td>
            <td class="positive">¥${formatNumber(item.range_commission)}</td>
            <td>${item.has_zijin_extra ? '<span class="badge bg-warning">含紫金国际</span>' : '-'}</td>
            <td>
                <button class="btn btn-sm btn-primary" onclick="showSpecialRangeDetail('${item.account}')">
                    查看详情
                </button>
            </td>
        </tr>
    `).join('');
}

// 更新缺失记录
function updateMissingRecords() {
    const container = document.getElementById('missingRecordsContainer');

    if (reportData.missing_records.length === 0) {
        container.innerHTML = '<div class="alert alert-success">✓ 没有缺失记录，所有中签都已填写卖出价格！</div>';
        return;
    }

    container.innerHTML = reportData.missing_records.map(record => `
        <div class="alert-missing">
            <strong>账户：</strong>${record.account} &nbsp;|&nbsp;
            <strong>股票：</strong>${record.stock} &nbsp;|&nbsp;
            <strong>位置：</strong>第 ${record.row} 行，${record.col} 列
        </div>
    `).join('');
}

// 显示账户详情
function showAccountDetail(accountName) {
    const account = reportData.accounts.find(a => a.account === accountName);
    if (!account) return;

    currentDetailAccount = account;

    document.getElementById('detailModalTitle').textContent = `${accountName} - 详细分成`;

    const tbody = document.getElementById('detailModalBody');
    tbody.innerHTML = account.stocks.map(stock => `
        <tr>
            <td>${stock.stock_name}</td>
            <td class="${stock.revenue >= 0 ? 'positive' : 'negative'}">¥${formatNumber(stock.revenue)}</td>
            <td class="positive">¥${formatNumber(stock.commission)}</td>
            <td>${stock.special_note ? '<span class="badge bg-info">' + stock.special_note + '</span>' : '-'}</td>
        </tr>
    `).join('');

    const modal = new bootstrap.Modal(document.getElementById('detailModal'));
    modal.show();
}

// 显示特定范围详情
function showSpecialRangeDetail(accountName) {
    const rangeData = reportData.special_range.find(r => r.account === accountName);
    if (!rangeData) return;

    currentDetailAccount = rangeData;

    document.getElementById('detailModalTitle').textContent = `${accountName} - 特定范围详细分成`;

    const tbody = document.getElementById('detailModalBody');
    tbody.innerHTML = rangeData.stocks.map(stock => `
        <tr>
            <td>${stock.stock_name}</td>
            <td class="${stock.revenue >= 0 ? 'positive' : 'negative'}">¥${formatNumber(stock.revenue)}</td>
            <td class="positive">¥${formatNumber(stock.commission)}</td>
            <td>${stock.special_note ? '<span class="badge bg-info">' + stock.special_note + '</span>' : '-'}</td>
        </tr>
    `).join('');

    const modal = new bootstrap.Modal(document.getElementById('detailModal'));
    modal.show();
}

// 导出账户收益表
function exportAccountRevenue() {
    const groupFilter = document.getElementById('accountGroupFilter').value;
    const sortBy = document.getElementById('accountSortBy').value;

    let accounts = [...reportData.accounts];

    if (groupFilter !== 'all') {
        accounts = accounts.filter(a => a.management_group === groupFilter);
    }

    accounts.sort((a, b) => {
        if (sortBy === 'revenue') return b.total_revenue - a.total_revenue;
        if (sortBy === 'commission') return b.total_commission - a.total_commission;
        if (sortBy === 'loss') return b.total_loss - a.total_loss;
        return 0;
    });

    const data = accounts.map((acc, index) => ({
        '排名': index + 1,
        '账户名称': acc.account,
        '分成比例': acc.rate_group,
        '总收益': acc.total_revenue,
        '总分成': acc.total_commission,
        '亏损金额': acc.total_loss
    }));

    exportToExcel(data, '账户收益表');
}

// 导出分成汇总表
function exportCommissionSummary() {
    const groupFilter = document.getElementById('commissionGroupFilter').value;

    let accounts = [...reportData.accounts];

    if (groupFilter !== 'all') {
        accounts = accounts.filter(a => a.management_group === groupFilter);
    }

    accounts.sort((a, b) => b.total_commission - a.total_commission);

    const data = accounts.map(acc => ({
        '账户名称': acc.account,
        '分成比例': acc.rate_group,
        '总收益': acc.total_revenue,
        '总分成': acc.total_commission,
        '亏损金额': acc.total_loss
    }));

    exportToExcel(data, '分成汇总表');
}

// 导出特定范围分成表
function exportSpecialRange() {
    const data = reportData.special_range.map(item => ({
        '账户名称': item.account,
        '分成比例': item.rate_group,
        '范围内收益': item.range_revenue,
        '范围内分成': item.range_commission,
        '备注': item.has_zijin_extra ? '含紫金国际' : ''
    }));

    exportToExcel(data, '特定范围分成表');
}

// 导出账户详情
function exportAccountDetail() {
    if (!currentDetailAccount) return;

    const data = currentDetailAccount.stocks.map(stock => ({
        '股票名称': stock.stock_name,
        '收益': stock.revenue,
        '分成': stock.commission,
        '说明': stock.special_note || ''
    }));

    exportToExcel(data, `${currentDetailAccount.account}_详细分成`);
}

// 通用Excel导出函数
function exportToExcel(data, filename) {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${filename}_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// 格式化数字
function formatNumber(num) {
    if (num === null || num === undefined) return '0.00';
    return num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
