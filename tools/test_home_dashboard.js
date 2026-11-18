// 首页「最近扫描批次 / 服务器端批阅进度」渲染回归测试（无需起服务）
//
// 做法：从 templates/index.html 里把三个渲染函数原样抽出来，在 node 里用假 DOM 跑一遍。
// 目的有两个：
//   1. 首页这两块不再读 /api/folders（旧流程目录，新扫的卷子不在那儿），改成读
//      /api/home/scan-summary；接口字段一变，这里会立刻失败。
//   2. 接口失败/空数据时必须给降级文案，不能白屏——这是「首页看着像坏了」的直接原因。
//
// 跑法：node tools/test_home_dashboard.js
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'templates', 'index.html'), 'utf-8');

const start = html.indexOf('function renderHomeScanBatches');
const end = html.indexOf('// 加载班级中心视图');
if (start < 0 || end < 0 || end <= start) {
  console.error('FAIL  没在 templates/index.html 里找到首页渲染函数（是否被重命名/删除了？）');
  process.exit(1);
}
const code = html.slice(start, end);

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
const api = new Function('escapeHtml', code + '\nreturn { renderHomeScanBatches, renderHomeGrading };')(escapeHtml);

function fakeEl() {
  return { innerHTML: '', insertAdjacentHTML(p, h) { this.innerHTML += h; } };
}

let fail = 0;
function check(name, cond, extra) {
  console.log((cond ? 'PASS  ' : 'FAIL  ') + name + (cond ? '' : '  → ' + (extra || '')));
  if (!cond) fail++;
}

// —— 用例 1：正常数据（照真实接口返回的结构造） ——
const summary = {
  batches: [
    { id: 'b1', created_at: '2026-10-30 14:23:13', subject_name: '地理', total: 3, done: 3, status: 'done', archived: 3, failed: 0, collect_only: false },
    { id: 'b2', created_at: '2026-10-30 14:21:44', subject_name: '地理', total: 30, done: 30, status: 'done', archived: 30, failed: 0, collect_only: false }
  ],
  grading: {
    running: [], running_count: 0, running_papers: 0, running_done: 0,
    today_batches: 2, today_papers: 33, records_total: 93, graded: 92,
    pending: 1, today_graded: 51, pending_confirm: 0
  }
};
const p1 = fakeEl(), r1 = fakeEl();
api.renderHomeScanBatches(p1, summary);
api.renderHomeGrading(r1, summary);

check('批次卡有「今日已扫 N 张」', p1.innerHTML.indexOf('今日已扫') >= 0);
check('今日进卷张数取接口值', p1.innerHTML.indexOf('>33</strong>') >= 0, p1.innerHTML.slice(0, 200));
check('每批写出扫描张数', p1.innerHTML.indexOf('扫描 30 张') >= 0);
check('每批写出已归档份数', p1.innerHTML.indexOf('已归档 30 份') >= 0);
check('已完成批次打标', p1.innerHTML.indexOf('已完成') >= 0);
check('进度卡有档案库总数', r1.innerHTML.indexOf('93 份') >= 0);
check('进度卡有待批阅数', r1.innerHTML.indexOf('1 份') >= 0);
check('无运行批次时说清楚', r1.innerHTML.indexOf('当前没有正在批阅的批次') >= 0);

// —— 用例 2：接口失败（summary=null）必须降级，不许白屏 ——
const p2 = fakeEl(), r2 = fakeEl();
api.renderHomeScanBatches(p2, null);
api.renderHomeGrading(r2, null);
check('接口失败时批次卡给引导', p2.innerHTML.indexOf('还没有扫描批次') >= 0);
check('接口失败时进度卡提示查服务', r2.innerHTML.indexOf('暂时取不到服务端进度') >= 0);

// —— 用例 3：正在批阅的批次，百分比与进度条要对 ——
const runningSummary = {
  batches: [{ id: 'b3', created_at: '2026-10-30 15:00:00', subject_name: '地理', total: 8, done: 2, status: 'running', archived: 2, failed: 0 }],
  grading: {
    running: [{ message: '正在批第 3 张…' }], running_count: 1, running_papers: 8, running_done: 2,
    today_batches: 1, today_papers: 8, records_total: 100, graded: 90, pending: 10,
    today_graded: 5, pending_confirm: 2
  }
};
const p3 = fakeEl(), r3 = fakeEl();
api.renderHomeScanBatches(p3, runningSummary);
api.renderHomeGrading(r3, runningSummary);
check('运行中批次显示 25%', p3.innerHTML.indexOf('批阅中 25%') >= 0, p3.innerHTML.slice(0, 300));
check('进度条宽度 25%', r3.innerHTML.indexOf('width:25%') >= 0);
check('带出后端给的那句进度文案', r3.innerHTML.indexOf('正在批第 3 张') >= 0);
check('待确认张数提示出现', r3.innerHTML.indexOf('2</strong> 张卷子没认出学号') >= 0);

// —— 用例 4：只扫描不批阅的批次，文案不能说「已归档」 ——
const collectSummary = {
  batches: [{ id: 'b4', created_at: '2026-10-30 16:00:00', subject_name: '政治', total: 5, done: 5, status: 'done', archived: 5, failed: 0, collect_only: true }],
  grading: { running: [], running_count: 0, running_papers: 0, running_done: 0, today_batches: 1, today_papers: 5, records_total: 5, graded: 0, pending: 5, today_graded: 0, pending_confirm: 0 }
};
const p4 = fakeEl();
api.renderHomeScanBatches(p4, collectSummary);
check('只扫描批次用「登记入库」措辞', p4.innerHTML.indexOf('已登记入库 5 份') >= 0, p4.innerHTML.slice(0, 300));

console.log(fail ? ('\n' + fail + ' 例失败') : '\n全部通过');
process.exit(fail ? 1 : 0);
