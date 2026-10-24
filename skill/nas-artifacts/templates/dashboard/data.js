// 只改這個檔就能更新儀表板:重新發佈 data.js 即可。
window.DASHBOARD = {
  title: '營運儀表板',
  updated: '2026-10-24 09:00',
  kpis: [
    { label: '本月營收', value: '1,284,000', unit: '元', delta: 12.4, goodWhenUp: true },
    { label: '活躍用戶', value: '8,932', unit: '人', delta: 3.1, goodWhenUp: true },
    { label: '平均回應時間', value: '420', unit: 'ms', delta: 8.0, goodWhenUp: false },
    { label: '錯誤率', value: '0.42', unit: '%', delta: -0.1, goodWhenUp: false }
  ],
  trend: {
    title: '近 7 天訪客數',
    xLabel: '日期',
    yLabel: '訪客數(人)',
    series: [
      { name: '本週', values: [820, 932, 901, 1034, 1190, 1130, 1320] },
      { name: '上週', values: [700, 780, 810, 870, 960, 940, 1010] }
    ],
    labels: ['一', '二', '三', '四', '五', '六', '日']
  },
  breakdown: {
    title: '各通路營收占比',
    xLabel: '營收(萬元)',
    items: [
      { name: '官網', value: 52 },
      { name: '門市', value: 38 },
      { name: '合作夥伴', value: 24 },
      { name: '其他', value: 9 }
    ]
  },
  table: {
    title: '最新項目',
    columns: [
      { key: 'name', label: '項目' },
      { key: 'owner', label: '負責人' },
      { key: 'count', label: '數量', num: true },
      { key: 'status', label: '狀態' }
    ],
    rows: [
      { name: '專案甲', owner: '小明', count: 1240, status: '正常' },
      { name: '專案乙', owner: '小華', count: 860, status: '注意' },
      { name: '專案丙', owner: '小美', count: 302, status: '正常' },
      { name: '專案丁', owner: '小強', count: 95, status: '落後' }
    ]
  }
};
