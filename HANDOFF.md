# QuickJudge 交接文档

> 生成时间：2026-10-26 19:55　｜　适用对象：接手的 AI agent 或开发者
> 项目根目录：`C:\Users\Administrator\WorkBuddy\2026-10-15-21-55-06\QuickJudge`
> 本文件是**当前状态的唯一准确描述**，请先读完第 1、2、9 节再动手。

---

## 1. 一句话现状

老师在本机点「开始扫描」→ 系统自动识别学号、按学号分班、逐篇批阅英语作文、生成话题语料与同学佳句、归档；
学生在任意设备浏览器登录后能看到自己的卷子、AI 评语、同学佳句。
**这条链路已经跑通并实测过 5 份真实手写卷**，当前处于「打磨识别质量」阶段，不是「从零开发」阶段。

---

## 2. 五分钟跑起来

```bash
cd "C:/Users/Administrator/WorkBuddy/2026-10-15-21-55-06/QuickJudge"
"C:/Users/Administrator/.workbuddy/binaries/python/envs/quickjudge/Scripts/python.exe" app.py
```

- 端口 **5001**，`host='0.0.0.0'`，`debug=True`（见 `app.py` 末行 `if __name__ == '__main__'`）
- 依赖装在专用 venv：`C:\Users\Administrator\.workbuddy\binaries\python\envs\quickjudge`
  （**没有 requirements.txt**；关键包：flask、flask_cors、openai、pillow、numpy、opencv-python 5.0、python-dotenv）
- 入口页：扫描端 `/station`（仅本机）、用户端 `/portal`（局域网/任意设备）
- 配置：`.env`（含真实密钥，已在 `.gitignore` 中排除，**不要提交、不要写进文档**）
- 首次起服务若 5001 被占用，先结束旧进程（历史上存在手工起的临时进程）

---

## 3. 系统形态：两端分离

| 端 | 路由 | 权限闸门 | 职责 |
|---|---|---|---|
| 扫描端 | `/station` | `_is_local_request()`：`request.remote_addr` 必须是 `127.0.0.1/::1`，否则 403 | 扫描、临时区整理、触发自动流水线、看进度 |
| 用户端 | `/portal` | 登录态（教师 / 学生） | 教师看本班、学生看自己的卷子与 AI 评语 |

`STATION_ALLOW_REMOTE=1` 可放开扫描端给远程（默认关闭，别随便开）。
学生账号 = 学号，密码 = 学号（`_ensure_student_registry` 自动创建）。

---

## 4. 主链路：点一次扫描，后面全自动

```
POST /api/station/scan  (auto=1 默认)
  └─ _run_scanner()            NAPS2.Console.exe，_SCAN_LOCK 串行，失败分诊
  └─ _new_staging_item()       入临时区，顺手做卷面体检（本地，免费）
  └─ _start_auto_batch()       建批次 + 起后台线程
       └─ _auto_pipeline()
            1. _vision_grade_essay()  读表头(学号/姓名) + 转写全文 + 批阅（一次调用）
            2. _ensure_student_registry()  自动建班 / 进名单 / 建账号
            3. _get_topic_corpus()    话题语料，跨批缓存
            4. _pick_awards()         全批一次横向评选佳句
            5. _commit_staging_item() 写入档案库 data/records
            6. 清空已入库的临时区条目
      进度写 data/batches/<batch_id>.json，前端轮询 /api/station/jobs/<bid>
```

**设计原则（改动时请守住）**：老师点完扫描后不应再有任何必填动作。
任何"让老师选一下模型/方案/参数"的设计都要先质疑。

---

## 5. 代码地图（行号会漂移，以函数名为准）

`app.py` 共 7669 行，单文件 Flask，`templates/station.html`(1005) / `templates/portal.html`(858)。

| 关注点 | 函数 / 常量 | 约行号 |
|---|---|---|
| 模型档位解析 | `resolve_vision_model()` | 49 |
| 视觉/文本模型关思考 | `_vision_model_kwargs` / `_text_model_kwargs` / `_vision_chat` | 44-96 |
| 学号→班级 | `student_id_to_class_code()`（4 位取前 2、6 位取前 4） | 1160 |
| 自动建班建号 | `_ensure_student_registry()` | 1225 |
| 学号抽取 | `extract_student_id()` + `_squash_digit_gaps()`（"1 2 3 4"→"1234"） | 3344 |
| 扫描执行 | `start_scan()` / `_run_scanner()` / `_SCAN_LOCK` | 2592 / 5340 / 120 |
| 临时区条目 | `_new_staging_item()` | 5423 |
| 入库 | `_commit_staging_item()` | 5484 |
| 对外序列化 | `_public_record()`（挂 essay / batch / quality） | 6083 |
| 作文提示词 | `ESSAY_PROMPT_TEMPLATE`（语篇改错 + 词汇/句式/衔接升级 + 个性化范文） | 6443 |
| 语料提示词 | `CORPUS_PROMPT`（points 分组 + 逐条翻译） | 6558 |
| 佳句提示词 | `AWARD_PROMPT` | 6578 |
| 送图压缩 | `_img_to_data_url(max_size=2400)` | 6425 |
| 卷面体检（本地） | `image_quality_metrics()` / `quality_badge()` | 6454 |
| 字迹体检（模型） | `assess_handwriting()` + `HANDWRITING_PROMPT` | 6575 |
| 读不准过滤 | `_norm_essay_result()`（含 `illegible` 过滤） | 6924 |
| 批阅结果归一化 | `_clip` / `_locate_quote`（quote 定位三级回退） | 6894 / 6902 |
| 范文超词数压缩 | `_essay_words` / `_fit_model_essay` | 7117 / 7129 |
| 作文批阅 | `_vision_grade_essay(page, requirements, max_score, model, quality)` | 7166 |
| 语料归一化 | `_norm_corpus()`（points[].items[text/trans/usage]） | 7330 |
| 语料缓存 | `_get_topic_corpus()`（精确指纹 + Jaccard≥0.55 模糊复用，只认 v2） | 7373 |
| 佳句评选 | `_pick_awards()` | 7421 |
| 自动流水线 | `_auto_pipeline()` / `_start_auto_batch()` | 约 7560 / 7710 |
| 用户端接口 | `/api/portal/*` | 7388 起 |

---

## 6. 数据与文件格式

```
data/
├── staging/index.json         临时区索引（含 quality 体检结果）
├── staging/images/<id>.png    扫描原图
├── records/index.json         档案库（含 essay / quality / handwriting）
├── records/images/<id>.png    归档图（与 records 同 id）
├── batches/<batch_id>.json    批次：item_ids / topic / corpus / awards / job / results
├── config/users.json          账号（role: teacher / student，学生密码=学号）
├── config/class_owners.json   班级归属
├── config/corpus_cache.json   话题语料缓存（按题干指纹）
├── config/naps2_config.json   扫描参数（driver/device/dpi/pagesize/source/bitdepth）
├── classes/<班级>/            班级目录
└── students/<班级>.json       班级学号列表
```

**批次 JSON 的关键字段**（前端直接消费）：
`job{status,stage,done,total,message}`、`topic{title,requirements,key}`、`corpus{points[{point,items[{text,trans,usage}]}],pitfalls[{point,wrong,right}]}`（v2，带 `version:2`；老批次仍是 phrases/sentences/patterns）、
`awards{awards[],standard,pool_size}`、`results[{student_id,name,score,level,model,quality,illegible,error}]`。

---

## 7. 模型配置与档位（本轮重点，别推翻重来）

`.env` 现状：

| 变量 | 值 | 用途 |
|---|---|---|
| `LLM_MODEL` | `Qwen/Qwen3-VL-8B-Instruct` | 看图主力（识别 + 批阅） |
| `LLM_TEXT_MODEL` | `Qwen/Qwen3.5-9B` | 纯文本（语料/佳句/译文/结构解析） |
| `LLM_MODEL_FAST` | 同 8B | 极速档 |
| `LLM_MODEL_STRONG` | `Qwen/Qwen3-VL-32B-Instruct` | 难卷兜底 |
| `LLM_QUALITY_MODE` | `auto` | auto / fast / standard / strong |

`resolve_vision_model(mode, hard)`：auto 时只有卷面被判"难"才升 strong。

### 已经踩过的坑（改模型前必读）
1. `Qwen2.5-VL-7B` 已下线（错误码 20012）。
2. `enable_thinking` **只有 32B-Instruct / \*-Thinking 认**；8B、30B-A3B 传了直接 400。按模型名判断 + 失败兜底重试，代码已处理。
3. 32B 默认走思考链，单页 141.7s；关掉后 6.2s。**不关就是不可用状态**。
4. `.env` 注释里写"推荐 32B"是旧结论，**已作废**，实测见第 9 节。

---

## 8. 卷面 / 字迹体检（本轮新增）

设计上刻意**不让模型给"字迹美丑"打 1-5 分** —— 实测同一张图反复打分会在 3/4 之间漂移，
风险点永远是「1 和 7 形似 / e 与 c 连笔 / 字压线」三句套话，等于没看图。

现在拆成两层：

1. **本地量化**（`image_quality_metrics`，cv2，毫秒级、零成本）
   有效 dpi、等效字高 `char_mm`、笔画宽 `stroke_px`、拉普拉斯方差、倾斜、墨量、噪点。
   结果写进 staging item 的 `quality`，扫描后立刻显示在卡片上。
2. **模型自报"读不准清单"**（`assess_handwriting`，或批阅时顺带产出 `illegible`）
   让它逐行转写、把没把握的词标出来（词级、带位置和原因），用**不确定率**折算 A/B/C/D。

阈值（按真实卷子校准过，别随手改）：
- `hard`（触发升档）仅当 `char_mm<1.4` 或 `sharp<1000` 或 `ink<1.2%`
- 等级：>12% 很难 / >8% 偏难 / >3% 较清晰 / ≤1% 清晰；本地 flag 额外扣分（small 扣 20）

接口：`POST /api/station/handwriting`，body `{"item_id": ...}` 或 `{"record_id": ...}`。

---

## 9. 实测基线（换模型/改参数后请复现对比）

**样本**：高一(14)班英语作文 5 份，Book1 Unit2「Discovering China」邀请信，满分 25。
学号样例：同班 5 个四位学号（已脱敏，形如 `14xx`）。

### 模型档位对比（3 图 × 8B/32B × 两种送图分辨率 = 12 次调用）
| 项 | 结果 |
|---|---|
| 学号识别 | **12/12 全对**，8B 与 32B 无差别 |
| 姓名识别 | 两边都不稳。同一张图 8B 给「王某勃 / 王某世」、32B 给「王某勤」 |
| 耗时 | 8B 6–13s；32B 11–53s（波动大） |
| 送图 1427px→1657px | prompt token 3114→4068（+30%），**结果无提升** |

结论：**8B 是性价比最优解，32B 不是无脑升级项**。默认走 8B，只对极端难卷升档。

### 卷面体检（本地指标）
| 图 | 字高 mm | 笔画 px | 清晰度 | 判定 |
|---|---|---|---|---|
| f0925f64 | 1.52 | 1.0 | 2975 | 字号偏小 |
| 589c10ee | 1.77 | 1.0 | 3495 | — |
| 14166b82 | 1.91 | 1.0 | 3574 | — |
| bc9657cc | 2.03 | 1.0 | 3965 | — |
| 1a4fed7b | 2.15 | 1.0 | 3378 | — |

字高 1.52mm 那张：模型转写 102 词、报 6 处读不准（`Jinggan`/`Hanzhou`/`Leifeng Tower`…）→ 等级 C。

### 批阅质量
得分 12/14/14/15/16（满分 25），5 份不同分；语料 8 短语 + 6 句全批共用并跨批缓存；佳句选出 3 条。
单页约 45s（8B），5 份含语料佳句约 4.5 分钟。

---

## 10. 已知限制（别当成 bug 反复查）

1. **一篇 = 一页**。跨页作文的第二页没有表头，读不到学号，会变成一条无主记录。
2. 多区域批阅时满分 = 每块满分 × 块数，**不支持不等分值分块**。
3. 全批得分完全相同时会写 `warning` 提示抽查（这是模型"抄示例"的信号，不是正常现象）。
4. 语料缓存靠题干指纹 + 字符二元组 Jaccard≥0.55 模糊命中，题干大改会重新生成（正常）。
5. `data/` 整目录被 `.gitignore` 忽略（含学生个人信息），别把它提交上去。
6. **报告 v2 让单篇批阅从 ~45s 涨到 ~200s**（8B，输出量 3 倍）。40 篇串行要 2 小时以上。
   要提速只有三条路：压输出配额、并发批阅（当前是串行）、或对「只求分数」的场景走 fast 档。
7. **8B 的衔接判断偏弱**：容易把「Moreover → Furthermore」这种同义替换当成衔接升级，偶尔还会把
   `Indeed` 标成因果。提示词已加限制（只报"该有却没有 / 关系用错"），但仍会漏，别指望它替代老师。
8. 范文超词数靠二次压缩兜底（`_fit_model_essay`）：8B 实测会写到 2–3 倍，压一次能落回要求词数，
   但要多吃一次文本模型调用（约 +5s、几分钱）。

---

## 11. 待办（按优先级）

### P0 — 影响可用性
1. ~~**git 未提交**~~ **已解决（2026-10-26 晚，commit `a46d35c`）**：`app.py`、两个模板、`tools/`、
   本文件已全部入库；`temp/`、`.scan-probe/` 已忽略（内含真实姓名）。
2. **扫描实际分辨率是 200dpi，不是配置的 300dpi**（已实测确认，见第 14 节）。
   出图固定 1656×2321，笔画只有 1 像素，这是"字小就看不清"的根因，比换模型重要得多。
   待验证：档位是否被 TWAIN 驱动钳死——需在馈纸器放纸才能测（`temp/probe_dpi.py`）。
3. `debug=True` 直接对外监听 `0.0.0.0`，正式使用前改掉。

### P1 — 质量提升
4. ~~**姓名识别不稳**~~ **已解决（2026-10-26 晚）**：姓名一律以名单为准，OCR 姓名不再写进账号；
   学号解析不出完整学号时**不建账号**，卷子归入「待确认」。详见第 14 节。
5. 跨页作文合卷（同学号多页自动合并）。
6. 多区域支持不等分值。

### P2 — 体验
7. 扫描端可加"本批 N 张卷面偏难"的汇总提示（流水线里已统计 `hard_n` / `illeg_n`，只在完成消息里用）。
8. `temp/probe_quality.py`、`temp/probe_models.py` 是本轮的实验脚本，可复用做对比，也可以删。

---

## 12. 本轮（2026-10-26）改了什么

| 文件 | 改动 |
|---|---|
| `app.py` | 新增 `resolve_vision_model` / `image_quality_metrics` / `quality_badge` / `assess_handwriting`；`_vision_grade_essay` 支持 `model`/`quality` 参数；批阅提示词新增 `illegible`；`_norm_essay_result` 保留并过滤该字段；`_new_staging_item` 写入 `quality`；`_auto_pipeline` 传 `quality` 并把 model/quality/illegible 记入批次结果；新增 `POST /api/station/handwriting`；`_llm_parse_paper_structure` 改用文本模型；`_img_to_data_url` 默认 2400 |
| `templates/station.html` | 卡片显示卷面体检徽章（字高/笔画/dpi）+「🔍 字迹体检」按钮 + 结果详情 |
| `templates/portal.html` | 作文顶部新增「🖊 字迹可靠度」区块（读不准清单 + 卷面标签） |
| `.env` | 新增 `LLM_QUALITY_MODE` / `LLM_MODEL_FAST` / `LLM_MODEL_STRONG` |

验证状态：`py_compile` 通过；`/station`、`/portal`、`/api/station/staging`、`/api/station/context` 均 200；
`/api/station/handwriting` 对真实记录返回 200 且结构正确；两个模板的 `<script>` 段经 Node 语法校验通过。

### 12.1 学生报告 v2（2026-10-26 20:20 起）

按老师给的 4 份范例 PDF（2023级1/2班 × 应用文 / 高三教学测试）重排学生报告页。
设计文档：`docs/report-v2-design.md`；真实输出预览：`docs/preview-report-v2.html`。

| 文件 | 改动 |
|---|---|
| `app.py` | 三个提示词重写：`errors` 改中文类型 + 强制可定位、`rewrites` 拆成 `vocab_upgrades` / `pattern_upgrades` / `cohesion`、新增 `content_upgrade`（读后续写）、`model_essay` + `model_essay_note`、`word_count`；`CORPUS_PROMPT` 改成按要点分组且逐条带翻译 |
| `app.py` | 新增 `_clip` / `_locate_quote` / `_essay_words` / `_fit_model_essay` / `_norm_corpus`；`_norm_essay_result` 增清洗、去重、按原文位置排序；`CORPUS_VERSION=2`（老缓存不再命中，会重生成一次）；`_vision_grade_essay` 加 20015 重试 + `max_tokens=6000`；`corpus_ready` 兼容 v2 |
| `templates/portal.html` | `essayHtml()` 重排为 8 段：字迹可靠度 → 要点覆盖 → **语篇批改（原文内联划改 + 编号 + 错因速查）** → 本篇佳句 → 高级词汇 → 高级句式 → 逻辑衔接/内容提升 → 个性化范文；新增 `inlineMarks()` / `annotateHtml()`；`batchHtml()` 支持 `points` 并**修掉 sentences 是对象却被 `esc()` 打印成 `{'sentence': …}` 的老 bug** |

**实测（高一某份卷子，满分 25）**：得分 18；9 处错误 **9/9 可在原文定位**；产出 3 条词汇升级 / 2 条句式 / 2 条衔接 / 一篇个性化范文。
耗时 **209.6s**（原链路约 45s）——输出量涨了 3 倍，这是 v2 的主要代价，见第 10 节。

**兼容策略**：老记录没有新字段时自动降级——`pattern_upgrades` 缺失就走原来的 `rewrites` 列表，
语料没有 `points` 就走 v1 的 phrases/sentences/patterns。不需要重批历史数据。


---

## 13. 交接前最后确认

- [x] `git add -A && git commit`（2026-10-26 晚完成：`a46d35c` 主链路 + `d9f91b7` 名单工具/测试）
- [ ] `.env` 里是真密钥，确认没被提交（`.gitignore` 已忽略，已核验）
- [ ] 新机器上先跑 `python app.py`，打开 `/station` 看临时区是否空、体检徽章是否正常
- [ ] 有扫描仪的那台机器才测得了扫描链路，其他机器只能测导入图片

> ✅ **已合并（2026-10-26 20:52）**：`app.py` 里学号口径重构与学生报告 v2 两轮并发改动已一并入库
> （`f43c24b`），其后的登录页账号提示、报告记录卡加宽、文档脱敏由本次合并提交收尾。
> `app.py`、`templates/portal.html`、`templates/index.html`、`static/css/scan.css` 当前均无未提交改动。

---

## 14. 2026-10-26 晚：名单换成真实数据 + 学号口径重构（重要）

### 14.1 名单已导入（取代原先 6 个测试账号）

用 `tools/import_roster.py` 导入学校下发的《学生姓名与学号对照表.xlsx》：
**1850 名学生 / 41 个班（26xx 高一 14 班、25xx 高二 14 班、24xx 高三 13 班）**。
学生账号（密码=学号）、`data/students/<班>.json`、`data/classes/<班>/` 三者已对齐重建；
`admin` 保留；旧测试账号（1402/1412/14xx/14xx/14xx/230205）与其班级 14、2302 已清除；
失效的班级归属（2302）已清理。账号表备份在 `data/config/users.json.bak-*`。

⚠️ **旧档案已清理（2026-10-26 20:37）**：`data/records` 里原 7 条旧测试记录
（学号 14xx/1412/1402/14xx/14xx/230205，学号 230205 为两条）及其图片、
只含这些记录的批次 `33543bf6…` 已删除，档案库现为空、门户从零开始。
删除前已备份到 `temp/backup-roster-work/deleted-orphans-<时间戳>/`（含 index.json、7 张图、批次）。
清理脚本 `temp/purge_orphan_records.py`（默认 dry-run，`--apply` 才动手）。

重新导入或换学校：`python tools/import_roster.py 新表.xlsx`（先加 `--dry-run` 看变更）。

### 14.2 学号口径（本轮最重要的认知）

- 学号 **6 位** = 年级 2 位 + 班号 2 位 + 座号 2 位（`260101` = 26 级 1 班 01 号），班级编码取前 4 位（`2601`）。
- **只有当前高一这一届卷面上只写「后四位」**（`260101` 写成 `0101`），高二高三写完整六位。
- 四位号在三个年级之间**完全撞号**（`0101` = 240101 / 250101 / 260101），
  而且**班号也完全相同**（都是 01），所以**班号无法用来消歧**——这是本轮实测确认的，
  不要退回「4 位前两位是班号」的旧思路，那样会把 `0101` 判成 1 班。
- 唯一可行的补全方式：**年级前缀 + 后四位**（`26` + `0101` → `260101`）。
  年级码配在 `.env` 的 `FOUR_DIGIT_GRADE=26`，学年初换届改成新码（如 27）。
  代码里由 `_four_digit_grade_prefix()` 读取，留空则回落 26。

### 14.3 新增/改动的函数

| 关注点 | 函数 | 说明 |
|---|---|---|
| 学号索引 | `_student_index()` | 学号集合、学号→姓名、后四位→[完整学号]；按 users.json 的 mtime+size 缓存 |
| 学号回查 | `resolve_student_id(observed, class_no)` | 返回 `(学号, 说明)`；说明为空=唯一命中，非空=待确认 |
| 建班建号 | `_ensure_student_registry(sid, name, class_no)` | 返回 4 元组 `(班级, 姓名, 完整学号, 说明)`；**姓名一律以名单为准** |

行为变化（有意为之，别改回去）：
1. OCR 姓名**不再**写进账号；名单有名字就用名单的。
2. 学号解析不出完整学号 → **不建账号、不建班**，卷子归入「待确认」班级，记录带
   `needs_review=true` 与 `sid_note`（原因），批次结果里同时给出 `student_id`（落库值）
   与 `student_id_seen`（卷面真实值）。
3. 临时区预识别（`/api/station/ocr` 那条路径）也会先回查，返回里多 `student_id_seen` / `sid_note`。

回归测试：`python test_student_id_resolve.py`（13 条用例，覆盖班号一位/两位、撞号、无匹配、空值）。

### 14.4 dpi 实测修正

配置 300dpi，实扫落盘 **1656×2321 ≈ 200dpi**（按 A4 宽 210mm 反推）——配置未生效已确认。
但「400/600 是否被钳制」**尚未测出**：那两次报 `No scanned pages` 是因为馈纸器里没纸
（用已知可用的 300 档做对照，同样报无输出，故排除档位原因）。
**该设备是馈纸式，一次命令会连续吸入多页**，`_collect_scan_pages` 收多页的逻辑是对的。
复测脚本：`temp/probe_dpi.py <dpi>`（必须放纸）。

---

## 15. 2026-10-26 21:20：接口鉴权收口 + 报告 v2 导出 + 衔接过滤

本轮针对「全项目扫描」发现的问题做了 5 处修复，**均已改代码并通过独立实例（5011）实测**。
⚠️ **5001 上的正式服务仍是旧代码，必须重启才生效**（`启动服务.bat`，或结束旧进程后重跑）。

### 15.1 旧版接口统一鉴权（P0，安全）
历史遗留的 `/api/image`、`/api/file`(DELETE)、`/api/vision_grade`、`/api/grade`、`/api/ocr`、
`/api/class_center` 等接口原本**完全没有登录校验**，项目对局域网开放后任何设备都能不登录就读图、
删文件、白烧模型额度（已用 curl 实测确认）。

修法：`app.py` 顶部新增 `@app.before_request _guard_api_auth()`：
- 公开：`/api/login`、`/api/logout`、`/api/me`、`/api/register_teacher`、`/api/invite_code/validate`；
- 豁免：`/api/station/*`（自带「仅本机」闸门）、`/api/portal/*`（逐条校验角色）；
- 其余 `/api/*`：未登录 → 401；非教师/管理员 → 403（学生仅保留 `/api/my_reports`、`/api/change_password`）。

实测：匿名 `/api/image/logo/1.png` 401、`/api/class_center/2601` 401；学生 403；教师 200；`/portal`、`/station`、`/api/station/context` 均 200。

### 15.2 报告 v2 可导出打印（P0，功能）
旧 `export_class_report` 只读 `TASK_RESULTS_DIR`，v2 流水线写的是 `records/` + `batches/`，两条链路不通。
- 新增 `utils/report_v2_export.py`：消费 `essay` v2 结构（要点覆盖/语篇批改/词汇/句式/衔接/范文）+ 批次 `corpus.points`，
  每生一页正文 + 一页语料，纯标准库、打印友好（`@media print`）。
- 新增 `GET /api/portal/export/report?class_name=&format=html|pdf`（教师/管理员；`pdf` 尝试 weasyprint，失败提示用 HTML 打印）。
- `templates/portal.html` 教师工具栏新增「导出学生报告」按钮（新标签打开，Ctrl/Cmd+P 存 PDF）。

实测：test_client 注入一条 v2 记录 → 200、`text/html`、17KB、七个区块齐全；无记录时 400 且给中文提示。

### 15.3 衔接「同义替换」过滤（P1，质量）
`_norm_essay_result` 新增 `_COHESION_CLUSTERS`：`before`/`after` 落在同一同义簇
（如 Moreover→Furthermore、However→Nevertheless）直接丢弃，不再当成「逻辑衔接升级」。
纯规则、零额外调用/延迟；仍在簇外的真实升级保留。

### 15.4 上传大小上限
`app.config['MAX_CONTENT_LENGTH'] = QJ_MAX_UPLOAD_MB`（默认 64MB），防止超大文件打满内存/磁盘。

### 15.5 旧文件接口路径包含校验（防御）
`/api/file` DELETE 与 `/api/image` 增加 `_path_within()`：目标必须落在项目根、班级目录、临时目录或扫描输出目录内，
防止将来框架行为变化时 `../` 拼出项目外路径。实测穿越尝试仍 404。

### 15.6 本轮**未做**（需要决策 / 需 A/B）
1. **model_essay 迁到文本模型并行**：实测范文只占输出约 10–15%，收益有限；改提示词有质量回归风险，建议先做单卷 A/B。
2. **200dpi 上采样**：放大 1.5× 会让图像 token 约翻倍（+成本 +延迟），且不增加真实信息；真正的解在扫描 dpi。
3. **跨页合卷**、**多区域不等分值**：都需要产品定义（合并策略、区域分值 UI），只改后端无意义。

### 15.7 验证与遗留
- `py_compile` 通过；`portal.html` 内联脚本 `node --check` 通过。
- 工作区未提交：`app.py`、`templates/portal.html`、`utils/report_v2_export.py`（新增），
  另有并发会话遗留的 `templates/index.html` 4 行改动，未由本轮触碰。
- 测试脚本：`temp/_test_report_export.py`（渲染器）、`temp/_test_route_export.py`（路由 test_client），均 gitignore。

---

## 16. 2026-10-26 21:30：班级认领可以退出（教师）

反馈：「认领之后要能调整，现在领了不能退」。定位到 **3 个叠加原因**，均已修：

1. **前端没有退出入口**：只有「认领」按钮，没有「退出」。现在自己名下的班级卡片上多了「退出认领」按钮
   （在卡片 meta 行右侧，点击不触发卡片跳转；带确认框）。
2. **已认领但 0 份试卷的班级不显示**：`overview` 的 `classes` 原先只由「有记录」的班级生成，
   刚认领还没扫描的班级连卡片都没有，自然没有退出按钮。现在教师的已认领班级即使 0 份也会出现（`total: 0`）。
3. **后端 `owned or None` 的越权/退不掉**：`_teacher_visible_classes` 在「退掉最后一个班」后 `owned` 变空，
   原逻辑会返回 `None`（不限范围）→ 反而能看到**所有班级**。改为：教师始终严格按认领范围，仅管理员不限。

同时把未认领班级从「提示行里挤 6 个按钮」改成独立面板 `#claimBar`：默认前 8 个，可「展开全部 N 个…」，
40+ 个班也不会撑爆页面；未认领时班级区提示改为「你还没有认领班级，请在下方认领」。

接口不变（`POST /api/portal/classes/claim`，`action=claim|release`，原先就支持 release，只是没人调用）。

验证（test_client + 内存伪造归属，未写真实数据）：
`teacher01` 只见 2401/2402/2403（0 份也显示）；`teacher05` 未见任何班、未认领 38 个；认领 2601 后可见，退出后消失并回到未认领；
退出全部后**不再**变成「看全部」。真实 `data/config/class_owners.json` 未被测试改动。

---

## 17. 2026-10-26 21:40：卷面方向自动纠正（上下颠倒 / 横放）+ 空白页跳过

### 17.1 有关系吗？——实测很有关系
同一张真实卷（`.scan-probe/dpi300.1.png`）走**真实批阅链路**对比 0° 与 180°（`temp/probe_grade_rot.py`）：

| 指标 | 0° | 180° |
|---|---|---|
| 学号读出 | **4 位** | **0 位（读不出）** |
| 全文转写 | 1333 字符 | 1112 字符 |
| 与 0° 文本相似度 | — | **0.25** |
| 分数 | 16/25 | 17/25 |
| 错误/词汇/句式/衔接 | 7/3/2/2 | 8/2/2/1 |

原因：`_vision_grade_essay` 会把**固定裁在顶部 10% 的表头特写**一起送模型（`ESSAY_HEADER_RECT`）。
页面一旦倒置，这个特写裁到的是页脚，学号读不到 → 卷子进「待确认」；转写质量也明显下降。

### 17.2 修法：本地检测 + 入库即转正
`_detect_page_rotation(img)`（app.py，纯 cv2，零 API 成本）判据：
- 正立文字行水平 → 行投影起伏大；横放时行投影接近平坦。
- 这类答题卷页眉有手写信息、页脚是空白页边 → 正立时上部墨量 > 下部。

在 `_new_staging_item` 里检测并**直接把图片转正**（`auto_rotation` 记档、扫描端显示徽章），
后续表头识别 / 批阅 / 归档 / 预览全部用正立图。置信度 <0.15 不转（宁可不转，不转错）。
实测 5 张真实卷 × 4 个方向（0/90/180/270）= **20/20 全对**，转正后与原图像素相似度 100%。

### 17.3 顺带修掉的 bug
`_vision_grade_essay` / `_read_header` 原先**完全不读 `item['rotation']`**——老师手动点「↻ 旋转」
只影响预览，批阅仍用原图。现在两者都接受 `rotation` 并先转正，流水线会把条目的 rotation 传进去。

### 17.4 和「正反」的关系
- **背面（空白页）**：馈纸式扫描会把背面也扫进来。本地识别 `ink < 0.3%` 判为空白页，
  流水线**直接跳过**（不批阅、不建档），完成消息注明「N 张空白页已跳过（可能是背面）」。
- **镜像（纸面朝下平扫）**：本地检测判不了镜像（镜像不改变投影/页眉分布），
  实测模型自报 `mirrored` 也不可靠（镜像时仍答 false）。这类目前仍需人工翻面/重扫；
  若要支持，建议在扫描端加一个「水平翻转」按钮（接口已有 rotate，仿写一个 flip 即可），暂未做。

### 17.5 代价与验证
- 入库时每张多 ~50–150ms 本地计算（1000px 下采样 + 4 方向），**不增加模型调用**。
- `py_compile` 通过；portal/station 内联脚本 `node --check` 通过；认领、导出路由回归通过。
- 测试脚本：`temp/probe_local_rot.py`（判向）、`temp/probe_rot_margin.py`（阈值标定）、
  `temp/probe_grade_rot.py`（真实 A/B），均 gitignore。


---

## 18. 2026-10-26 21:35：扫描端可以选「满分」

反馈：「扫描界面，目前还是没有让我选择分值的地方」。

### 17.1 根因：满分是写死的常量
`_auto_pipeline` 里批阅调的是 `_vision_grade_essay(path, requirements_text, ESSAY_MAX_SCORE, ...)`，
`ESSAY_MAX_SCORE = 25` 是模块级常量。所以即使老师在「批阅方案」里选了数学（12 分）或默写（10 分），
卷子照样按 25 分制评——提示词里的 `{max_score}` 永远是 25，学生端 `5/25` 也是错的。
方案下拉里虽然显示了「满分 X」，但没有任何代码消费它。

### 17.2 改法：满分随批次走
- 新增 `_subject_max_score(subject)`：读 `grading_config` 里该方案自己的 `max_score`（作文 25 / 语文作文 50 /
  数学 12 / 默写 10 / 通用 100），查不到回落到 25。
- 新增 `_coerce_max_score(value, subject='')`：校验前端传来的值（必须 0 < v ≤ 1000），非法就回落到方案默认值。
- 新增 `_subject_display_name(subject)`：档案里的 `subject_name` 不再硬写「英语作文」。
- `_new_batch(..., max_score=None)` / `_start_auto_batch(..., max_score=None)`：批次里存 `max_score`。
  **重试复用同一批次**（`/api/station/jobs/<id>/retry` 直接重跑 `_auto_pipeline`），所以整批口径天然一致。
- `_auto_pipeline` 开头算一次 `batch_max`，预热、并行、失败补跑全用它；入库的 `max_total` 也用它。
- `_grade_essay_record`（教师端「重新生成」）同样沿用原口径：批次 `max_score` > 该记录已有 `max_total` > 方案默认。
  否则 12 分制的卷子重批一次会变成 25 分制。
- `station_scan` / `station_import` 接收 `max_score`，并在响应里回带 `batch.max_score`。

### 17.3 前端
`#subjectSel` 右侧新增「满分」数字输入框（`#maxScore`）：
- 默认值 = 当前方案的 `max_score`，切换方案会跟着换；
- 老师改过的值按**学科**记在 `localStorage['qj_station_max']`，下次扫同一学科不用重填；
- 改回等于方案默认值时自动清掉记忆，避免「改过一次永远生效」；
- 非法输入（空/0/负数/超大）失焦时纠正回默认值；发送前 `readMaxScore()` 再兜一层。
- 扫描请求体新增 `max_score`；进度条首行显示「已扫描 N 页（满分 X）」，让老师当场确认口径。

### 17.4 验证
- `temp/test_max_score2.py`：六个方案默认值、非法输入回落、`'18'`/`30.5` 正常透传，全部符合预期。
- `temp/test_scan_api.py` / `_2`：test_client 打 `/api/station/import`（`_auto_pipeline` 打桩，不烧模型额度），
  传 `max_score=18` → 批次落盘 `18.0`；不传 → `25.0`；`subject=math_solution&max_score=12` → `12.0`。测试数据已清理。
- `station.html` 内联脚本 `node --check` 通过；5001 实际渲染出 `id="maxScore"`（页面 `maxScore` 出现 10 次）。

### 17.5 遗留（未做，需要产品定义）
自动流水线目前**只走英语作文契约**：选了数学/默写，满分是对的，但提示词仍是作文那套（要点覆盖/语言错误/句子改写）。
要做学科分流，得给 staging 条目加「读图 → 通用评分契约」的分支，属于独立改造，未在本次范围内。

> 服务已于 21:36 用新代码重启（旧进程 15648/1864 已结束），`/station`、`/portal` 均 200。

---

## 19. 2026-10-26 21:45：管理员默认看全校 + 先堵住非英语作文的自动批阅

两条反馈，一条是权限诉求，一条是「宁可不开也别批错」。

### 19.1 管理员默认看到全校所有试卷
原来管理员虽然**后端不限制**（`_teacher_visible_classes()` 对 admin 返回 `None`），但首页的班级卡片
只由「有记录」的班级生成——没扫过的班根本不出现，管理员打开门户看到的是空页面，也没法点进去翻。
改动：

- `portal_overview`：`role == 'admin'` 时把 `_load_classes_roster()` 里的**全部班级**补进列表（0 份也显示），
  有卷的班按最近提交排前面，空班按名单顺序垫底。教师分支不动（仍只看自己认领的班）。
- 前端 `renderTeacher`：管理员标题变「全校班级」，卡片徽章不再标「他人班级 / 未认领」（对管理员没意义），
  统一标「全校」；0 份的卡片淡化并显示「还没有扫描件」。
- 新增「查看全部试卷（N 份）」按钮（仅管理员可见）：不选班，直接按时间倒序列出**跨班全部试卷**，
  每条记录头部补一个班级号（`VIEW.all` 时才显示），避免看不出是哪个班的。

验证（test_client，真实名单 41 个班）：admin → `classes` 41 个、`total` 0（当前无记录）；
`teacher01` → 8 个（认领范围），未被带偏。

### 19.2 非英语作文先禁止走「扫完即批」
自动流水线的提示词和输出结构都是英语作文专用（要点覆盖 / 语言错误 / 句子改写 / 佳句），
选数学（12 分）或默写（10 分）走这条链路会产出一堆对不上的字段。学科技能还没做，所以先挡住。

- 新增 `AUTO_PIPELINE_SUBJECTS = ('english_essay',)`、`auto_pipeline_supports()`、`auto_pipeline_reject_message()`。
- `station_scan` / `station_import`：**在驱动扫描仪、落盘之前**就 400 返回，附中文说明
  （说清为什么不行 + 两个替代做法），不让老师扫完一叠纸才发现批不了。
- `_start_auto_batch` 里再挡一次（防将来新增调用点绕过校验）。
- `_auto_pipeline` 也挡：老批次（放开学科前建的）点「重试」时直接置 error 并给同样的提示，
  不会拿作文契约再跑一遍。
- `grading_subject_options()` 每个方案多返回一个 `auto_ok`，前端不硬编码学科名。
- 扫描端下拉只列 `auto_ok=True` 的方案；被藏起来的方案在下面黄色提示条里列名 + 说明，
  并告诉老师「要批这些科目，扫进来后在用户端逐份手动批」。

⚠️ 只限制「自动批阅」。教师在用户端对已入库卷子点「评语 / 批阅」走的是 `_grade_record` 通用契约，
数学/默写仍可正常逐份手动批——那是另一条（正确的）链路，没有一并禁掉。

验证（test_client + `_auto_pipeline` 打桩）：
`general/chinese_essay/math_solution/dictation` 的 `auto_ok` 均为 False，只有 `english_essay` 为 True；
`import?subject=math_solution&auto=1` → 400，且**批次目录条数不变**（没有建批、也没有落盘）；
`english_essay&max_score=20` → 200 且批次 `max_score=20.0`；直连 `_start_auto_batch(subject='dictation')` → ValueError。

> 服务已于 21:44 用新代码重启，`/station`、`/portal` 均 200。

---

## 20. 2026-10-26 22:25：39 张卷面「凭空消失」——索引被覆盖，已整批恢复

反馈：「我之前曾经扫过三十几张进去的，为什么现在一张识别的也没有？是正在识别中吗」
**答案：不是在识别，当时也没有任何任务在跑——是临时区索引被别的进程写空了，图片一张都没丢。**

### 20.1 现场
- `data/staging/images/` 里躺着 **39 张真实卷面**（1655×2323，1.6–2.5MB，mtime 全部是 21:38:14~21:38:18），
  另加 6 张 300×400/1KB 的测试图。
- `data/staging/index.json` 是 `[]`，`data/records/index.json` 只有 2 条（2614 班，21:36 那批）。
- 批次记录里**没有任何 39 张的批次** → 那次扫描只完成了「写图片」，没走到建批。
- `D:/ScanTests` 只剩 `scan1.39.png`（与临时区里某张同尺寸），说明 39 页是同一批吸入、
  逐张被 move 进临时区后，最后一张没来得及搬完。

### 20.2 根因
`data/staging/index.json` 是典型的「读-改-写」：`load_staging()` → 改 → `save_staging()`。
进程内的 `_STAGING_LOCK` 只能挡住**同一个进程**里的线程；**第二个进程**（测试脚本、
另一个服务实例、另一个 agent 会话的工具）拿着过期快照写回来，就能把整批条目覆盖掉，
而 `images/` 下的文件不动——于是「列表空了、图还在」。

⚠️ 我自己的排查脚本（`temp/test_scan_api*.py`、`temp/test_admin_and_guard.py`）就属于
**第二个进程**，它们在 21:41/21:43/21:53 都调用过 `save_staging`。按时间线它们嫌疑最大，
以后任何脚本都不该对**线上数据目录**做 `save_staging`，要在副本上跑。

### 20.3 处理
1. **恢复**：`tools/restore_staging.py`（新增，只读图片、不删不移），按 mtime 重建 39 条，
   保留 `id`（=文件名去扩展名）、`quality` 体检结果，`original_name` 记为 `restored-NN.png`，
   `note` 标「从图片恢复（索引曾丢失）」。原空索引已备份为 `staging/index.json.bak-*`。
   复核：扫描端 `/api/station/staging` 返回 39 条，取图 200 / image/png / 1.67MB。
2. **兜底**：`save_staging()` 写之前，若新列表比旧列表**更短**，先把旧索引存成
   `staging/index.prev.json`；条目骤减到不足一半时在控制台打警告。
   实测：39 → 3 时备份里是 39，再写回 39 无异常。

### 20.4 遗留 / 待办
- 恢复出来的条目**没有学号/班级**（原本就没识别过），走「全部提交」会被跳过（它只提交已填学号的）。
  要对这批出成绩，需要把 39 个 item_id 交给 `_start_auto_batch()` 重跑一遍批阅（会自动读表头分班）。
- 那个「写图片成功、建批没发生」的 21:38 请求为什么断在半路，缺服务端日志（当时的进程已被换手），
  无法定论；但 `station_scan` 里「先落盘再建批」的顺序本身是合理的，不必改。
- 更彻底的做法（未做）：给 `save_staging` 加「调用方必须持锁」的断言，或改成按 id 增删的 append-only 存储。
