# -*- coding: utf-8 -*-
import re

path = r'D:\WK工作空间\live-dashboard\live-dashboard\trend-demo.html'
with open(path, 'rb') as f:
    html = f.read().decode('utf-8')

# 统一换行为 \n 处理（清除所有历史残留 \r，规范化行尾）
html = html.replace('\r', '')

# ── 1. 图表 x 轴 label：🥇🥈🥉 → 数字前缀 + 金/银/铜刻度色 ──
old_chart = """        chartY1 = new Chart(ctxY1, {
          data: {
            labels: labels.map((l, i) => {
              const rankings = findTop3Indexes(labels.map(label => datasets[label].metrics[selectedY1Metrics[0]] || 0));
              return showTop3 && rankings[i] ? `${['🥇','🥈','🥉'][rankings[i]-1]}${l}` : l;
            }),
            datasets: y1Datasets
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { position: 'top', labels: { boxWidth: 12, padding: 15 } } },
            scales: { y: { type: 'linear', position: 'left', title: { display: true, text: '绝对值' } } }
          }
        });"""

new_chart = """        // 排名标签：数字前缀 + 金/银/铜色
        const rankColors = ['#c9a962', '#9ea4ad', '#b08a3c'];
        const y1LabelInfo = labels.map((l, i) => {
          const rankings = findTop3Indexes(labels.map(label => datasets[label].metrics[selectedY1Metrics[0]] || 0));
          const rank = rankings[i];
          return showTop3 && rank ? { text: `${rank}. ${l}`, color: rankColors[rank-1] } : { text: l, color: '#7a8290' };
        });

        chartY1 = new Chart(ctxY1, {
          data: {
            labels: y1LabelInfo.map(x => x.text),
            datasets: y1Datasets
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { position: 'top', labels: { boxWidth: 12, padding: 15 } } },
            scales: {
              y: { type: 'linear', position: 'left', title: { display: true, text: '绝对值' } },
              x: { ticks: { color: (ctx) => y1LabelInfo[ctx.index] ? y1LabelInfo[ctx.index].color : '#7a8290' } }
            }
          }
        });"""

assert old_chart in html, '图表代码未找到'
html = html.replace(old_chart, new_chart, 1)
print('OK 图表排名标签已替换')

# ── 2. 表格徽标：🥇🥈🥉 → Lucide medal 图标 ──
old_badge = """        const topBadge = showTop3 && rankings[idx] ? `<span class="top-badge top-${rankings[idx]}">${['🥇','🥈','🥉'][rankings[idx]-1]}</span>` : '';"""

new_badge = """        const topBadge = showTop3 && rankings[idx] ? `<span class="top-badge top-${rankings[idx]}"><i data-lucide="medal"></i>${rankings[idx]}</span>` : '';"""

assert old_badge in html, '表格徽标代码未找到'
html = html.replace(old_badge, new_badge, 1)
print('OK 表格徽标已替换')

# ── 3. 徽标样式 ──
old_css = """    .top-badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 11px; margin-left: 4px; }
    .top-1 { background: #F2D37E; color: #7a5c12; }
    .top-2 { background: #E4E1D8; color: #5c584d; }
    .top-3 { background: #E0A878; color: #fff; }"""

new_css = """    .top-badge { display: inline-flex; align-items: center; gap: 3px; padding: 2px 7px; border-radius: 4px; font-size: 11px; font-weight: 600; margin-left: 4px; }
    .top-badge i[data-lucide] { width: 12px; height: 12px; color: currentColor; }
    .top-1 { background: #F2D37E; color: #7a5c12; }
    .top-2 { background: #E4E1D8; color: #5c584d; }
    .top-3 { background: #E0A878; color: #fff; }"""

assert old_css in html, 'CSS未找到'
html = html.replace(old_css, new_css, 1)
print('OK 徽标CSS已更新')

with open(path, 'w', encoding='utf-8', newline='\n') as f:
    f.write(html)
print('ALL DONE trend-demo.html 更新完成')
