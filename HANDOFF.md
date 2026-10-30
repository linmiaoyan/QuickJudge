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

## 21. 2026-10-27 07:00：卷子全倒着进「待确认」，而且没法手动处理

反馈（附截图）：「有一篇逆向的进入了"待确认" 1.逆向应该被正确处理 2.没有看到手动处理它的方法」

这两条其实是**同一批 39 张的连带伤害**，只是先被倒置和一张错学号暴露出来。

### 21.1 倒置：为什么自动纠正没生效
`_detect_page_rotation()` 本身没坏，拿那 28 张复核全部返回 180°（置信度 0.39~0.66，远高于
0.15 的门槛）。真正的原因是**入库路径绕过了体检**：

- `_new_staging_item()` 会同时做「体检 + 定向 + 旋转原图」，正常扫描走的都是它。
- §20.3 的恢复走的是 `tools/restore_staging.py`，那个脚本只重建索引**没有做定向那一步**，
  于是 39 条全是 `rotation: 0`，图片也保持倒着的原始像素。
- 前端是按 `rotation` 渲染的，`0` 就等于「告诉界面这张是正的」——倒着显示就成了必然。

处理：
1. 28 条已入库记录把 `rotation` 改回 `180`（备份 `temp/backup-20261027/index.before-rotate.json`）。
2. 临时区剩的 10 张直接旋转原图像素，复核 `_detect_page_rotation` 全部回到 0。
3. `tools/restore_staging.py` 补上定向 + 旋转（和 `_new_staging_item` 对齐）。
   **以后再从图片恢复，不会再产出一批倒着的条目。**

### 21.2 「待确认」：以前只能看，不能改
`/api/portal/records/<id>/comment` 原来收到 `student_id` 只是**原样写进去**，而班级只在
「本来为空」时才补——`待确认` 是个非空字符串，所以班级永远不会被改，卷子就永久卡在待确认里。
另外 `_filter_records_for_session` 严格按班级范围过滤，`待确认` 不属于任何班，
**任课老师连看都看不到**，只有管理员能瞄一眼。

处理：
- 接口改为**认名单**：`resolve_student_id()` 解析后不在名单里就返回 400 并带未匹配原因
  （实测 `999999` → 400「学号「999999」不在名单里」，数据零改动）；匹配上就同步改正
  `class_name`、`student_name`，清掉 `needs_review` / `sid_note` / `student_id_seen`。
- 同一个接口接受 `rotation`（0/90/180/270），老师肉眼看出倒了可以直接转正，不动原图。
- 门户编辑弹窗加「学号」「卷面方向」两个字段 + 待确认提示条；记录卡加 `待确认` 徽章和
  「指定学号」按钮。
- `_filter_records_for_session` 对 `待确认` 单独放行，任课老师也能认领自己班的。

实测那张：表头 OCR 读成 `2416`，名单里 `江丞博` 唯一对应 `261416`
（`1416` 的 1 被读成 2）。提交 `1416` → `261416 / 江丞博 / 2614`，`needs_review` 归 False。
提交后 `待确认` 计数 **0**。

### 21.3 顺带补的归档幂等（同一批暴露出来的）
批次 `95e573c678da4045` 卡在 `committing` 时留下 31 条记录，续跑会把已入库的再建一遍。
- 归档按 `staging_item` 建幂等表 `_records_by_staging_item()`：同一条临时区条目只允许一条档案。
- 每入库一条就 `_drop_staging_items()` 落盘一次，中断最多丢一条（原来是整批跑完才清）。
- 批次加 `runner.boot`（每进程一个 `_BOOT_ID`）：进程换了就把 `running` 自愈成 `interrupted`
  + 提示「还有 N 张留在临时区」，不再永久转圈。
- `_start_auto_batch()` 加了重叠检测 `_find_running_batch()`，同一批图片不会被起两个批次；
  `/api/station/jobs/<id>/retry` 对正在跑的批次返回 409（连点两下就是重复归档的源头）。
- 建批时若调用方没带满分，先从临时区条目回读 `_staging_max_score()`，不再一路回落到 25。

### 21.4 教训
**任何「重建索引」的工具都必须复刻建索引时的全部副作用**（体检、定向、旋转原图），
否则数据看起来是对的，只是重要字段静默变成默认值——比直接报错难查得多。
`rotation: 0` 和「图片是正的」是两件事，前者只是后者的**断言**。

## 22. 2026-10-27 07:15：批阅断点续跑（一张批完就落盘）

### 22.1 为什么加
批次 `95e573c678da4045` 在 40 分钟内被**连续打断三次**，每次都是进程在批阅途中被换掉
（07:11:52 那次新进程日志里直接写着「批次 95e573c678da4045 状态已修正为中断」，说明是
**另一个进程重启了服务**，把正在跑的 daemon 线程一起带走了，不是自己崩的）。
而批阅结果原本只在内存里（`results = [None] * n`），进程一没就全丢——
每次续跑都要把那 10 张从头再烧一遍，一次约 10 次视觉调用。

### 22.2 做法
- `data/batches/partial/<batch_id>/<item_id>.json`：**一张批完立刻原子写一个文件**
  （`_save_partial` 在 `_grade_one` 里，不是循环外——这是「不重烧额度」的全部保证）。
  放在子目录而不是 `batches/*.json`，否则 `_find_running_batch` 会把它当批次文件扫。
- `_auto_pipeline` 开跑先 `_load_partials()`：已有结果的直接填进 `results`，
  `state['done']` 从复用数起算，预热循环里 `continue` 跳过——不会重复调用模型。
- 复用结果里若已带 `task.requirements`，直接拿来当本批题干，连预热那张都省了。
- 批次 `status='done'` 时 `_clear_partials()` 收尾，不占磁盘。
- 中断提示语跟着改：有断点就显示「已批完的 N 份会直接复用，不会重新烧额度」。

### 22.3 边界
断点只保证**同一批次**续跑不重复花钱；换了批次 id（重新建批）就复用了别的目录，用不上。
模型调用本身没有幂等键，所以「批到一半的某一张」还是会重来一次——这是可接受的，
一张的量而已。


## 23. 2026-10-27 07:20：视图改成「列表 → 整页」，学生端可导 PDF

### 23.1 为什么改
以前两种视角都是「一屏塞很多张卡」：教师一行三张班级卡、点开后在下方续一段列表（网格还挂在上面），
学生端干脆把**所有**记录一次性全展开在同一条长页里。报告 v2 之后每份卷子有 8 个区块，
塞进 360px 的格里就是互相打断——想看一份完整的，得来回滚。

改成两层：**一层只负责选（网格 / 批次），一层只负责看（整页）**。

### 23.2 教师端
- `#teacherView` 拆成 `#classHome`（统计 + 班级网格 + 未认领栏）与 `#detailBox`（整页详情）。
  点班卡 → `classHome` 加 `hidden`，整页只剩这个班；「← 返回班级」复原。
- 整页顶部 `.pagebar` 吸顶（`top` 由 `--tb` 实时校准顶栏高度），承载：返回、标题+共N份/已批/待批、
  状态筛选、学号筛选、为本批生成评语、导出报告 PDF、导出清单 CSV。原来两行工具栏挤在网格下方，
  现在全在这一条里。
- 记录列表 `.records.full` 单列整宽 + `.rec.wide`：**左栏卷面图 sticky**，右栏报告，
  往下读长报告时原卷一直在视野里（对照划改时很关键）。
- `#navStrip` 名单跳转条：一个班几十份卷子，横排「学号 + 分数/待批」，点一下滚到那份。
- `recCardHtml()` 抽成教师/学生共用的卡片函数。原来两处各写一遍模板，这次要加第三种布局，
  再抄一遍必然走样；顺手把并发会话加的「待确认 / 指定学号」并进了同一份。

### 23.3 学生端
- `#studentView` 拆成 `#studentHome`（统计 + 批次卡 `#studentBatches`）与 `#studentPage`（整页）。
- `groupBatches()` 按 `batch_id → task_id → 标题+日期` 回退分组，**最新批次在前**；
  卡面给任务名、提交时间、份数、批阅状态、得分。
- 点批次 → 整页渲染该批全部卷子（含本题语料 / 同学佳句），页头「← 返回列表 / 导出 PDF」。
- 只有一批的学生也照样走列表，不搞分支——少一条只在特定数据下才跑到的路径。

### 23.4 「导出 PDF」的实际做法
导出的 HTML 以前**没有任何打印入口**，点开只是一张只能看的网页，等于导了个寂寞。
- `utils/report_v2_export.py` 顶部加 `.pbar` 工具条（含「打印 / 存为 PDF」，`@media print` 里隐藏）。
- 本机**没有 weasyprint 也没有 pdfkit**（实测两个 import 都失败），`format=pdf` 这条路走不通，
  所以按钮打开的是打印友好页，靠浏览器的「另存为 PDF」。别再对着 `html_to_pdf` 猜为什么 500。
- 学生也能调 `/api/portal/export/report`：`_filter_records_for_session` 对 student 先按本人学号过滤，
  `class_name` 参数对学生会清空，`batch_id`/`task_id` 只能收窄范围，构造不出越权。
- 顺手给整页视图加了打印样式：Ctrl+P 直接出干净版面（隐藏顶栏/工具条/名单条，单列、区块不跨页拆）。

### 23.5 一个差点踩进去的坑
`.rec{overflow:hidden}`（裁图片圆角用的）会让内部元素的 `position:sticky` **完全失效**——
卷面图不会报错，只是安安静静跟着页面滚走。宽卡改成 `overflow:visible`，圆角交给 `.pic` 自己裁。

### 23.6 验证
- `py_compile`、内联脚本 `node --check` 通过。
- HTML 里 74 个 id 与 JS 引用的 74 个 id **完全对齐、无重复**（重构最容易死在这一步）。
- DOM 桩跑真实脚本：学生 2 张批次卡 → 点入整页 8 个区块齐全；教师 1 张班级卡 → 3 张整页卡 + 3 个名单跳转。
- Flask `test_client`：学生导出 200 / 18KB 且不含他人；伪造 `class_name` 不漏；
  `batch_id` 收窄生效；教师按班导出 2 份 / 31KB；未登录 403。

### 23.7 注意 / 遗留
- **5001 正式服务要重启**才带上新的导出参数（`batch_id`、学生角色）；模板改动靠 `TEMPLATES_AUTO_RELOAD` 刷新即见。
- `#statScope` 那条「当前查看 XXX」现在命中不了（开详情时它所在的 `classHome` 已隐藏），
  统计信息挪到 `.pagebar` 标题里了。代码留着没删，别以为它还在工作。
- 服务端一键出 PDF 需要装 weasyprint；不装就走浏览器打印，功能等价、体验差一步。

## 24. 2026-10-27 09:38：满分口径折算成 15 + 同学号撞车不再静默建档

### 24.1 为什么要折算
整班 41 条里 39 条按 `/25` 出分、2 条按 `/15`，口径不统一。老师当时扫这批填的确实是 15，
但那个值在临时区被覆盖、用恢复工具捞回来后就丢了（见 §18 / §21）。重批要再烧一轮视觉调用，
而分数本身是线性的 —— 老师选了**按比例折算**。

新增 `tools/rescale_scores.py`（默认只预览，`--apply` 才写）：
- 一次改全一条记录的 8 个字段：`max_score`、`score_text`、`grading.max_total`、
  `grading.total`、`grading.items[*].max_score`、`grading.items[*].score`、
  `essay.max_score`、`essay.score`。**漏一个界面就自相矛盾**，所以写死在一处。
- `level`（档次名「良好」）不动：定性档位，不随刻度变。
- 保留一位小数：18/25 → 10.8/15、22/25 → 13.2/15、16/25 → 9.6/15。
- 写盘前自动比对备份，并提示**先停 5001 服务**（索引是读-改-写）。
- 顺手修了两条「字段写着 25、分数其实是 /15」的错标记录（261412、261402）。

批次的 `max_score` 也同步改成 15：否则老师点「重新生成」时
`_grade_essay_record` 会读批次的 25 又按 25 分制重出一遍（这条链路见 §18）。

### 24.2 同学号撞车
那批留下的 261431 有 4 条（22/17/22/17）、261435 有 2 条（16/19），
**四张原图哈希各不相同**——不是重复归档，是有别人的卷子被 OCR 认成了这个学号。
本次按老师要求不动这些数据，但以后不能再静默产生：

- `_auto_pipeline` 归档前先建 `seen_sid`（本批已有记录的学号表），第二条起
  `class_name = 待确认`、`needs_review = True`、`sid_note` 写「本批已有学号 X 的档案…」，
  直接落到待确认里等老师用「指定学号」改，而不是混在班里把人数顶高。
- 并发会话的 `temp/flag_conflicts.py` 已把现存 6 条打上同样的标记（措辞统一成
  「请核对后改成正确学号」），门户里会显示「待确认」徽章。

## 25. 2026-10-27 09:45：预览里的卷子还是倒的——角度没落到图片上

反馈：「预览界面看到的试卷，还是逆向的」。§21 只把 `rotation` 标成 180，
那是**显示时**转，图片本身还是倒的。

### 25.1 三个原因叠在一起
1. **二次旋转的坑**：`_commit_staging_item()` 归档时已经把 rotation 和 crop
   烘进 `dest_path`（`_transform_image(...).save(dest_path)`），却又把 `rotation`
   原样存进记录。取图时再转一次 → 180° 转两遍正好倒回去。
   自动批阅固定传 `rotation=0`，所以这条链路一直没暴露，但谁传一次非 0 就中招。
2. **恢复出来的图本身是倒的**：那 28 张走 `tools/restore_staging.py`，跳过了
   `_new_staging_item` 的定向步骤。标 `rotation=180` 只修好了 `?processed=1`
   的预览；「看原图」（`/image` 原样返回）、导出、重新批阅拿到的仍是倒字节。
3. **浏览器缓存**：`?processed=1` 这个地址在 `rotation` 还是 0 时就返回过倒图，
   字节被缓存住；后来改成 180，地址没变，于是怎么改都还是看到旧的倒图。

### 25.2 处理
- `tools/bake_rotation.py`（新增）：把角度真正转进像素，然后 `rotation` 归零，
  记 `rotation_baked` 留痕。只处理 `rotation ∈ {90,180,270}` 且**没有**
  `crop`/`regions` 的记录（区域框是在旋转前的画面上量的，转了会错位，跳过并提示）。
  改前整批备份图片到 `temp/backup-<日期>/records-images/`，
  索引备份 `index.json.bak-bakerot-<时间>`，图片用 `.tmp` + `os.replace` 原子替换。
  已处理 28 张，失败 0。
- `_commit_staging_item()`：既然烘了就存 `rotation: 0` / `crop: None`，杜绝二次旋转。
- `_public_record()`：取图地址加 `v=<图片mtime>-<rotation>-<crop>`。
  图片换过、角度改过地址就变，缓存不会再骗人。
- `portal_record_image()`：没有可变换的东西就直接发原文件——以前不管有没有变换都会
  把 2MB PNG 重编码一遍，一个班预览下来纯属白烧 CPU。同时补 `Cache-Control`。
- 老师改「卷面方向」也直接转图（按 `delta = (新 - 旧) % 360` 转再归零），
  不再只改字段，避免「预览正了、看原图还是倒的」。

### 25.3 复核
41 条记录的**原图**方向检测全部 0°（`temp/diag_rotation.py`）；
经真实路由取 `?processed=1` 与不带参数的 `/image`，两者字节数一致、都是正立，
说明短路生效、没有重复编码。

### 25.4 教训
`rotation` 这类「显示时才生效」的字段，只要有一个消费方忘了带参数就不一致。
**能烘进原始数据就不要留在渲染层**；而且烘了就必须归零，半烘半存比不烘更难查。



## 26. 2026-10-27 11:06：首页印出教师注册邀请码（管理员不显示）

反馈：「管理员账号不需要在首页提示 教师注册邀请码：wzkjgz」＋「后台设定一下这个值，
目前没出现过」＋「只在管理员下隐藏」。

### 26.1 现状
`data/config/invite_codes.json` **根本不存在**，`load_invite_codes()` 返回 `[]`，
也就是当时一个可用邀请码都没有——注册接口必然报「邀请码无效」。
页面上也确实没有任何「邀请码：xxx」的字样。

### 26.2 处理
- 用管理员接口 `POST /api/admin/invite_codes {code:"wzkjgz", max_uses:50}` 建码，
  走的是正常校验路径（重名会拒）。改次数/换码在「配置 → 邀请码管理」。
  这个文件在 `data/` 下，被 `.gitignore` 忽略，属于运行时数据，不进版本库。
- 登录页 `#loginBox` 里「教师注册（需邀请码）」下面加一行
  `教师注册邀请码：<code>`；码值由 `_home_invite_hint()` 每次渲染实时读配置，
  换码后页面自动跟着变，不需要改模板。
- `_home_invite_hint()` 三条规则：
  ① `session.user_role == 'admin'` 直接返回空（管理员有「配置 → 邀请码管理」，印首页只是噪音）；
  ② 只挑 `used_count < max_uses` 的码，用完的码印出来只会让人注册失败；
  ③ 没有可用码时整行不渲染。
- 前端再加一道兜底：`/api/me` 回来若 role 是 admin，把 `#inviteHintLine` 隐藏，
  防止缓存页面把旧 HTML 露出来。

### 26.3 验证
test_client 三个身份取 `/`：
- 未登录 → 有提示行、含 `wzkjgz`
- admin → 无提示行、页面里不含码值
- teacher → 有提示行、含 `wzkjgz`

### 26.4 注意
邀请码印在公开登录页上，等于**任何能打开首页的人都能注册成老师**（上限 50 人）。
学校内网这么用没问题；要收紧就把 `max_uses` 调小，或者把这一行删掉、改为线下发码。

## 27. 2026-10-27 11:12：教师新手教程从登录页移到教师端

登录页那块（「教师注册」「第一次使用？打开教师新手教程」）是给还没有账号的人看的，
把教程入口放在那儿，等于让已经登录的老师去登录页找文档。

- 登录页删掉教程链接（`templates/index.html` 的 `#loginBox`）。
- 教师端首页 `#homeTeacherBlock` 在「📌 使用流程」下面加一张卡片：
  一句话说明 + 「打开教师新手教程 ↗」按钮。
- 侧边栏「其他入口 → 使用教程」补上 `sidebar-teacher`，`applyRoleLayout()` 会把
  学生身份下的它隐藏掉 —— 教师教程对学生没有意义。

静态页 `/guide`（`使用说明/教师新手教程.html`）本身没动，只是入口位置变了。

## 28. 2026-10-27 11:20：旧版阅卷流程收进折叠组（与自动批阅的合并建议）

### 28.1 诊断：这套旧流程当前是零使用
- `data/config/tasks.json` **不存在**，`load_tasks()` 返回 `[]` —— 一个阅卷任务都没建过；
- `data/task_papers/`、`data/task_results/` 都是空的；
- 所有批次里**没有一个**带 `task_id`。
批记录（`data/records/`）全部来自自动流水线。所以「藏起来」不承担任何数据风险。

### 28.2 为什么不能只砍「组卷布置」
`阅卷管理` 的「创建阅卷任务」向导第一步列的是 `/api/tasks` 里的**已发布任务**，
而发布任务的唯一入口就是组卷布置。单独砍掉它，旧链第一步就断了。
→ 结论：**整链一起收进折叠组**，而不是掐头。

### 28.3 做了什么
- 侧边栏顺序：概览 → 文件与扫描 → 旧版阅卷（备用）→ 其他入口。
- 旧版整组默认收起，标题右侧一个「展开」小按钮，状态记 `localStorage.qj_legacy_open`。
- 首页「📌 日常就三步」+ 一个小按钮「旧版阅卷流程 →」（`openLegacyFlow()`：
  展开分组 + 跳阅卷管理）。
- `switchTab()` 切到 assign/grading/ocr/center 时自动展开该组，否则侧边栏看不出自己在哪。
- 三处空状态文案里「请先在组卷布置中选题」改成指向「创建阅卷任务」/「旧版阅卷 → 组卷布置」。

### 28.4 更彻底的合并（建议，尚未做）
两套流程真正的差别只有一处：**评分依据从哪来**。
- 旧流程：老师先在「组卷」里写题干 + 评分细则 → 建任务 → 收卷 → 批。
- 自动流程：直接从第一张卷面推断题干（`b['topic']`），细则是通用契约。

也就是说旧流程的「任务」本质上是自动批阅的一个**可选输入**，不是另一条流水线。
现状里批次已经带 `task_id`，但**只当归档标签用**（`_commit_staging_item` 拿它填
`task_title`、往 `task_papers/` 拷一份），没有参与评分：
`_auto_pipeline` 的 `requirements_text` 始终来自卷面推断。

要真合并，改动点很小，只有一处：
```python
# app.py _auto_pipeline 里
requirements_text = '\n'.join('· ' + str(x) for x in ((b.get('topic') or {}).get('requirements') or []))
# 改成：批次绑了 task_id 就优先用任务里的题干/要求
_task = _find_task(b.get('task_id'))
if _task:
    requirements_text = _task_requirements_text(_task)   # 题干 + 各小题要求
```
这样「按指定题目精批」就不再是另一条流程，而是自动批阅的一个开关：
建好任务 → 扫描时选上它 → 后面全自动。旧版那三个页面退化成任务管理，可以彻底隐藏。

## 29. 2026-10-27 14:0x：旧流程真正并进自动批阅（task_id 参与评分 + 批改备注）

按 §28.4 的建议落地。核心变化：**「按指定题目精批」不再是一条独立流程，而是自动批阅
的一个输入**——老师在扫描端选个题目、写一句备注，点扫描后仍然一路自动到最终结果。

### 29.1 题干来源的优先级
`_auto_pipeline` 里 `requirements_text` 不再是「卷面推断」一条路：

1. **绑定任务的题干**（`task_id` → `find_task` → `task_requirement_lines`）：最高优先，
   同时把任务的 title/requirements 写进 `b['topic']`，话题语料也按它缓存；
2. **断点复用的旧结果**里带的 `essay.task`；
3. 以上都没有才走老办法：串行预热第 1 张，读它的 `task.requirements`，其余复用。

附带效果：**题干已知时跳过预热，整批直接并行**。以前每批都要先等一张读完才知道题干，
现在选了题目就省掉这段等待。

### 29.2 老师批改备注（`teacher_note`）
- 批次字段 `teacher_note`（`_new_batch` 新参数），由 `station_scan` / `station_import` 的
  请求体带进来，长度上限 300 字；
- 拼进 `ESSAY_PROMPT_TEMPLATE` 末尾新增的一段 `{teacher_note}`，并**在提示词里声明优先级
  高于卷面上的印刷要求**（老师会写「只看第 2 段」「按中考标准，不扣书写分」这类话）；
- 归档时写进记录的 `teacher_note`，事后能追溯「这批为什么这么判」；
- 重试沿用同一个批次，所以备注不会丢。

### 29.3 轻量任务 `POST /api/tasks/quick`
旧「组卷布置」对英语老师无用（他们在别的平台组卷），但「这次批哪道题、要点有哪些」
仍然需要。所以给了一个只有两个输入框的入口：
- body `{title, requirements（多行文本或数组）, subject?}`，`class_names` 允许为空
  （扫描端按学号自动分班，班级不在这里定）；
- 带 `id` 就是改已有的那张（只改 title / requirements）；
- 生成的 task 带 `kind:'prompt'`、`requirements: [...]`；`/api/station/context` 里
  额外给出 `has_reqs`，扫描端只列**有要点**的任务（没要点的选了也帮不上批阅）。

### 29.4 前端
- `templates/station.html`：新增 `<details id="promptBox">`「📝 本批题目与批改备注」，
  **默认收起**，扫卷主按钮位置不变。内含题目下拉（含「＋新建题目」内嵌表单：名称 + 一行一条
  的评分要点）和备注 textarea；选过题目/填过备注就自动展开、标题旁显示「已设置」。
  题目选择与备注都记在 `localStorage`（同一次作业连续扫多批不用重复填）。
- 导入本地图片改成与扫描同一条链路（`auto=1` + 题目 + 备注 + 满分），成功后同样盯进度；
  勾了「先只扫描」才退回只进临时区。
- `templates/index.html`：侧栏删掉「组卷布置」入口（页面 `#tab-assign` 仍在，可
  `switchTab('assign')` 打开）；三处空状态文案改指向扫描端的新入口。

### 29.5 验证
用 mock 掉 `_vision_grade_essay` 的方式跑完整 `_auto_pipeline`：3 张 → 3 次调用全部带
`requirements='· 说明推荐理由\n· 提到时间地点'`、`note='只批第二段'`、`max_score=15`；
`b['topic']` 是任务标题；3 条记录的 `teacher_note` 都在。`/api/tasks/quick` 新建 + 修改
均 200，`/api/station/context` 的 `has_reqs` 为 true。前后端脚本 `node --check` 通过。

### 29.6 还没做
- 记录里的 `teacher_note` 目前只在 JSON 里，用户端界面没展示（`_public_record` 未透出）；
  老师若想在报告里看到「本批按什么口径批的」，需要再补一行展示；
- 阅卷管理「创建阅卷任务」向导对 `class_names` 为空的任务表现未验证。

## 30. 2026-10-27 16:35：批改备注在试卷门户里露出来（教师视角）

§29.6 那条遗留补掉了。备注以前只躺在记录的 JSON 里，老师扫完根本看不到自己写的口径
到底被用上了没有。

### 30.1 后端：按角色下发
`_public_record` 开头新增一句：
```python
if session.get('user_role') not in ('teacher', 'admin'):
    out.pop('teacher_note', None)
```
**必须 pop 而不是靠模板不渲染**：`out = dict(rec)` 是浅拷贝，字段本来就会随
`/api/portal/records` 一起发给学生端，只是界面上没画。老师写的备注可能带班级学情判断
（"这个班普遍差"），不该出现在学生手机上。

### 30.2 前端：一行琥珀色说明
`templates/portal.html`
- `.rec .tnote`：12px、`#8a5300` 字色 + `#fff8ec` 底 + 虚线边框，和 `.task`（题目行）区分开；
- `recCardHtml()` 在题目行下面插一行 `🗒 本批批改备注：…`，带 title 提示说明它的优先级；
- 教师端列表里**同一批只在该批第一张卡上出现**（`noteSeen` 按 `batch_id || task_id || id`
  去重）。整批 40 份卷子刷 40 个一样的框没有意义。

学生端 `renderStudentBatchPage` 走同一个 `recCardHtml`，但后端已把字段剥掉，条件自然为假。

### 30.3 验证
`_public_record` 三种角色：student → 无该键，teacher/admin → 有值；`/portal` 渲染出的
HTML 里 `.rec .tnote`、`本批批改备注`、`r.teacher_note` 均在。脚本 `node --check` 通过，
服务已重启（`stale: false`）。

### 30.4 注意
- 老批次没有这个字段，卡片上就不显示，不会有空行；
- 记录编辑弹窗里还改不了这个备注（改了要重新批），要改只能重扫一批或改批次 JSON。

## 31. 2026-10-29 10:35：API token 加密（设备移交不泄漏）

### 31.1 要解决的具体问题
`.env` 里两个值是明文：`LLM_API_KEY`（学校的模型额度）和 `FLASK_SECRET_KEY`（会话签名密钥）。
后者泄漏比前者更糟——拿到它就能自己签一个 `user_role=admin` 的 Cookie 直接进后台。
设备以后可能交给第三方厂家，一个能读文件的人一个 grep 就全拿走了。

### 31.2 方案：Windows DPAPI（本机 + 本账户 + 应用熵）
新增 `secretbox.py`，用系统自带的 `CryptProtectData`（ctypes 调用，**零第三方依赖**，
装不了 `cryptography`/`pywin32` 也能用）：

```
python tools/seal_secrets.py --set LLM_API_KEY     # 交互录入（不回显、不进 shell 历史）
        ↓
data/config/secrets.dpapi.json                      # {items:{LLM_API_KEY:"dpapi-user:AQAA…"}}
        ↓ 服务启动时
secretbox.get_secret('LLM_API_KEY')                 # 自动解密，无需人输口令
```

- 加密密钥不在项目里：是当前 Windows 用户在 LSA 里的主密钥，由账户口令保护；
- 密文绑定 **本机 + 本账户 + 应用熵** `QuickJudge/secrets/v1`。把 `data/` 整个拷到别的机器、
  或换个 Windows 账户登录，`CryptUnprotectData` 直接返回错误（实测 `GetLastError=13`）；
- `LSA` 那层平时是透明的，所以服务仍然可以无人值守启动（`启动服务.bat` 双击即跑）。

### 31.3 代码改动
- `app.py`：`_llm_key` / `app.secret_key` 改走 `secretbox.get_secret()`；
  取值顺序 **密文 → 明文环境变量 → 默认值**（密文优先，避免"以为加密了其实一直读明文"）；
  启动横幅多打一行体检：
  `密钥保险箱：已加密 2 项（LLM_API_KEY、FLASK_SECRET_KEY）`。
- `secretbox.py`：`encrypt_text` / `decrypt_text` / `get_secret` / `set_secret` / `status` /
  `startup_report` / `fingerprint`。解不开时**不崩**：记 warn、回退明文、`--status` 里能看见原因。
- `tools/seal_secrets.py`：
  | 命令 | 用途 |
  |---|---|
  | `--status`（默认） | 加密了哪些、解不解得开、`.env` 里还剩没剩明文 |
  | `--apply` | 把 `.env` 明文搬进保险箱，并把那几行换成注释（其余配置一字不动） |
  | `--set NAME` | 录入/更换（getpass，不回显） |
  | `--unseal NAME` | 打印明文，换机器时导出用 |
  | `--verify` | 密文 vs `.env` 明文比对**指纹**（sha256 前 8 位，不打印明文） |
  | `--list-plaintext` | 找出磁盘上还含明文密钥的文件 |
  | `--shred-backups` | 删 `.env.bak*` 等副本（默认预演，`--yes` 才删；删前把旧值存成 `<名字>_OLD`） |

### 31.4 已完成迁移（2026-10-29 10:32）
- `LLM_API_KEY` 51 字符 → 密文，指纹 `392c1d59`；`FLASK_SECRET_KEY` 64 字符 → 密文，指纹 `81f6e426`；
- `.env` 里两行已换成注释，其余 11 项配置与顺序未变；
- 服务重启后启动横幅显示"已加密 2 项"，`--verify` 全绿；
- **实测**：用解出来的 token 真调了一次 `chat/completions`（max_tokens=1），返回 200，
  说明不是"能解开但值坏了"。

### 31.5 威胁模型（必须跟接手方讲清楚）
| 场景 | 是否防住 |
|---|---|
| 拷走项目目录（含 `data/`） | ✅ 拿到的是密文，没有本机主密钥解不开 |
| 同一台机器、**换一个 Windows 账户**登录 | ✅ DPAPI 用户作用域不同，解不开 |
| 拿到硬盘/备份镜像，换机器挂载 | ✅ 主密钥在 LSA，不在磁盘明文里 |
| 用**同一个** Windows 账户登录（账户连同设备一起交出去） | ❌ 能解开——DPAPI 保护的是账户，不是人 |
| 进程运行中读内存 | ❌ 明文在内存里（调模型必须如此） |
| 有本机管理员权限、能 dump LSA 主密钥 | ❌ 有这权限基本等于拿到账户 |

最后两条**无法靠代码消除**，只能靠流程。所以交接时必须做下面这套动作。

### 31.6 设备移交清单（交给第三方厂家前）
1. `python tools/seal_secrets.py --list-plaintext` —— 确认没有残留明文（现在只剩一个旧 `.env.bak`，见 §31.7）；
2. 到服务商后台**吊销并重发** token（这一步才是"让旧密钥失效"，本地删文件做不到）；
3. 新账户/新机器上 `--set LLM_API_KEY` 重新灌入，或干脆删掉 `data/config/secrets.dpapi.json`
   让对方用他们自己的 token；
4. 删掉 `data/config/users.json` 里的账号，或至少改掉所有管理员口令；
5. 交接前把 `QuickJudge-BACKUP*` 目录一并清掉（那里是 `data/` 的完整镜像，密文虽解不开，
   学生信息却是明文）。

### 31.7 遗留
- `.env.bak-20261016_113150` 仍含一个**更早的** token（指纹 `d4c62caf`，与当前不同）。
  若它还有效，请先在服务商后台吊销；确认无效后跑
  `python tools/seal_secrets.py --shred-backups --yes` 删掉。
- 换 Windows 账户/机器后 `--status` 会明确报"解不开"，此时的修复动作是 `--set` 重灌，
  不是去找回旧密文（旧密文在新账户下永远打不开）。

## 32. 2026-10-30 08:35：政治 / 地理「要点判分」模块（学科路由 + 逐点给分）

### 32.1 需求原文与拆解
> 1. 教师会给每一个练习输入简单的参考答案和答题要点。
> 2. 我传入卷子，教师输入改卷要点。
> 3. 达到要点则得分，不达到或者要点中有错别字不得分。
> 后续还会有很多学科，所以每张卷子进来的时候你要区分好具体是哪个学科，做好简单的路由。

拆成四件事：**练习（评分标准）** → **学科识别** → **按 kind 路由** → **逐点判分与归档**。
关键是第 3 句：判分不是"整体印象给分"，而是"逐条要点命中 → 分值累加"，
这与英语作文那套（整篇打等级）在数据结构、提示词、归档字段上全都不一样，
所以用**批阅契约 kind** 区分，而不是在代码里写 `if subject == 'politics'`。

### 32.2 批阅契约 kind（路由的唯一依据）
| kind | 含义 | 学科 | 流水线 |
|---|---|---|---|
| `essay` | 整篇作文，按要点打等级 | 英语作文 | `_auto_pipeline` |
| `points` | 逐条要点判命中，按要点分值累加 | 政治/道法、地理（后续历史、生物…） | `_points_pipeline` |
| `generic` | 只能手动逐份批 | 数学、默写、语文作文、通用 | 不开放自动批 |

- 定义在 `DEFAULT_GRADING_CONFIG` 每个方案里，`load_grading_config()` 对老配置
  `setdefault('kind', 'generic')`，所以**老 data/config/grading_config.json 不用迁移**；
- `subject_kind(key)` / `auto_pipeline_supports(key)` 取代了原来的
  `AUTO_PIPELINE_SUBJECTS = ('english_essay',)` 元组；
- `grading_subject_options()` 把 `kind` / `need_exercise` / `auto_ok` 下发前端，
  前端按 kind 切 UI，**不再硬编码学科名**。

**新增一门要点判分学科 = 在 `DEFAULT_GRADING_CONFIG` 加一项 `kind='points'` + 写提示词**，
前端（扫描端「本批练习」、门户「要点判分明细」）自动生效，不用改代码。

### 32.3 练习（评分标准）数据层
- 文件：`data/config/exercises.json`（`EXERCISES_FILE`，原子写）；
- 结构：`{id, subject, title, body(老师粘的原文), questions[], max_score, point_count, created_at, updated_at}`；
- 一题：`{no, stem, answer, points:[{text, score}], max}`；`max` = 该题要点分值合计；
- 接口：`GET/POST /api/exercises`（POST 支持 `dry_run`）、`GET/DELETE /api/exercises/<id>`；
- POST 只收 `kind='points'` 的学科，其余 400；解析不出题目也 400（带 `warnings`）。

**解析器 `parse_exercise_body()`**（老师手写的东西必须宽容）：
| 认得出 | 例子 |
|---|---|
| 题号 | `1.` `1、` `（1）` `第1题` `一、` |
| 参考答案 | `答：` `答案：` `参考答案:` |
| 要点 | `要点：`、行首 `- － — · • * +`、或题号后的普通行 |
| 分值 | `（2分）` `2分` `| 2` `: 2` ` 2`（空格结尾，最后才试） |
| 多要点同行 | `要点A（1分）；要点B（1分）` |

分值优先级：**要点分值合计** > 题中标称分（不一致时报警告并以后者为准）。
只有参考答案没有要点 → 退化成"对照答案整体判"并警告；两者都没有 → 跳过该题并警告。
`dry_run` 就是给老师看这份解析结果（几题/几点/总分/警告）用的——分值写在行尾这种写法，
机器猜错了必须让他一眼看见。

### 32.4 判分：分数由服务端算，绝不信模型算术
`_norm_points_result(data, exercise)`：
- 按**老师的要点表**逐条遍历（不是按模型返回的顺序），模型只负责判 `hit/typo/evidence`；
- `points` 数组按下标一一对应；数量不一致 → 记 `align_warnings`；
- **错别字规则落到代码里**：`hit && typo` → 强制 `hit=False`、该要点 0 分，并记警告。
  不这样做模型会"两头讨好"（说命中又报错字），而老师的要求是"要点中有错别字不得分"；
- 每题得分 `min(各命中要点分值之和, 题满分)`，总分 `min(Σ, 卷满分)`，全部 `round(x, 2)`；
- 输出：`{questions[{no,stem,answer,max,score,points[],hit_count,point_count,note,comment,missing}],
  total, max_total, level, comment, strengths, issues, illegible, text, student, align_warnings}`。

`POINTS_PROMPT_TEMPLATE` 里把三件事写死：要点表原样喂进去、`points` 顺序/数量必须与标准一致、
`hit` 只填 true/false（分数不让它算）。老师的批改备注 `{teacher_note}` 拼接在末尾、优先级最高。

### 32.5 流水线与路由
`_run_pipeline(batch_id)` 是**唯一入口**（扫描/导入/重试/崩溃恢复都从这里进）：
```
subject == 'auto' → _resolve_auto_subject() 定成具体学科
kind = subject_kind(subject)
  essay  → _auto_pipeline
  points → _points_pipeline
  其他   → 落盘 error（附 auto_pipeline_reject_message）
```
`_points_pipeline`：
1. 取练习（没有 → 报错，见 32.8）；`batch_max = exercise_total_max(exercise)`；
2. `_read_headers_and_assign()`（与作文共用，读学号/姓名/班级）；
3. 并行 `_vision_grade_points()`，**每判一份 `_save_partial` 落盘**，重启/中断可续批；
4. 失败补跑一次（并发下多为限流/超时）；
5. 归档：`points`、`points_meta{exercise_title,point_count,hit_count,align_warnings}`、
   `max_score=练习总分`、`essay=None`、`grading{items,total,max_total,level,exercise_id}`；
6. 幂等：`_records_by_staging_item()` 命中就复用原档案，不建第二份。

### 32.6 学科自动识别（`SUBJECT_AUTO`）
- 扫描端方案下拉里多一项「自动识别学科」（`subject='auto'`）；
- `_detect_subject()`：只认**第一张非空白页**（同一批几乎必然同科，逐页识别等于成本翻倍），
  长边 1400、`max_tokens=300`、temperature=0，返回 `(key, confidence, evidence)`；
- 别名表 `_DETECT_ALIAS` 兜住模型爱写的同义词；认不出/认错学科 → 记 `subject_detect_error`，
  `confidence < 0.5` 时**仍然采信**但把"可能认错"挂到批次上，扫描端和门户都能看到；
- 结果写回 `b['subject']` / `b['subject_detected']`，之后一切照常。

### 32.7 扫描端交互（`templates/station.html`）
- 「📝 本批X与批改备注」块按 kind 一分为二：
  `#essayTaskWrap`（作文选题干）/ `#exerciseWrap`（要点判分选练习）；`syncSubjectUi()` 切换；
- `#exerciseSel` 只列当前学科的练习；`#exerciseMeta` 显示「N 题 / M 个要点 / 总分 X」+ 查看修改；
- 「＋新建练习」：学科 + 名称 + 正文（`解析预览` 走 `dry_run`，`保存并使用` 落盘并选中）；
- **满分输入框在要点判分学科里自动上锁**并显示练习总分——后端会用练习总分覆盖，
  不锁住就会重演"填了满分却不生效"的困惑（§29 同类问题）；
- 请求体带 `exercise_id`；`_precheck_auto_batch()` 在**驱动扫描仪之前**就挡掉"没选练习"。
- 任务列表里要点判分显示 `得分/满分`，完成文案带「判分依据：《练习名》N 个要点 / 满分 X」。

### 32.8 出问题时的两条可操作路径
| 情形 | 提示 |
|---|---|
| 自动认成政治但没选练习 | 「这批卷子认成「政治 / 道法」，按要点判分需要先选一份练习…选完点**重试本批**——不用重新扫」 |
| 绑定的练习被删了 | 「本批绑定的练习已不存在（可能被删了）…重新选一份后点重试」 |

为此 `POST /api/station/jobs/<id>/retry` 接受可选 `exercise_id`：补选练习后直接换绑重试，
**不要求老师把纸重新放回扫描仪**（卷子已经扫进来了，重扫纯属折腾）。

### 32.9 门户展示（`templates/portal.html`）
- `pointsHtml(r)`：按题逐点摊开，`✓ 得分 / ✗ 未得分`、每条要点右侧 `得/满分`、
  错别字要点带红色 `错别字 当家做主 → 当家作主` 标签、`依据：`显示卷面原文；
  末尾附「识别到的作答」供对照；
- 学生端/教师端都渲染（老师要核对，学生要知道自己哪条没写）；
- `errStats()` 扩到要点判分：平均分 + **要点命中率** + **「未得分最多的要点」Top5**
  （某条要点 30 份里 28 份没写，多半是没讲到，讲评比逐个学生看更有用）。

### 32.10 验证记录（2026-10-30）
隔离环境（把 `app.py` 复制到 `temp/pttest/`，`DATA_DIR` 因此落在临时目录，不碰真实档案）跑
`temp/pttest_test.py`，30 项全绿：
- 解析：2 题 / 5 要点 / 10 分；每要点 2 分；提示词含题号、参考答案、`—— 2 分`；
- 算分：模型报 `total=99` → 服务端算成 **8**；错别字那份 `hit` 被压成 0 分 → 该题 2 分、
  总分 **6**，并产生「已按不得分处理」警告；
- 归档：2 份记录、`score_text` `8/10` 与 `6/10`、`max_score=10`、`essay=None`、
  `points` 两题、`points_meta.exercise_title`、`grading.items`、`teacher_note` 均正确；
- 路由：`subject='auto'` → 识别 `politics` → 走要点流水线 → done；
- 缺练习：状态 error，提示含「练习」「重试」；
- 真实服务（`temp/smoke_live.py`，只读）：`/api/station/context` 里政治/地理 `kind=points`、
  `need_exercise=true`，作文仍 `kind=essay`；`dry_run` 解析不落盘；作文学科 POST 被拒 400；
  `/portal`、`/station` 均含新渲染代码；`_code_stale_info()` 不再报陈旧。

### 32.11 遗留与注意
- **要点分值写错不会自动纠正**：`max_score` 完全由老师给的要点分值决定，
  这也是为什么 `dry_run` 预览必须让老师确认一眼；
- `points_meta.hit_count` 只统计"这份卷子命中几点"，用于门户顶部与共性分析，
  不是分数依据（分数看 `total`）；
- 「要点中有错别字不得分」只对**该条要点**生效，不影响同题其他要点；
  若老师的本意是"整题作废"，应该在要点里把该点写成独立要点再判；
- 同一份练习被多批共用时，`max_score` 随之固定——改练习分值会让新旧批次口径不同，
  门户按记录里的 `max_score` 各自显示，不会串；
- 方案配置里的 `prompt` 只服务**手动批阅**那条路（`_get_grading_subject` → `grade_paper`），
  自动要点判分用 `POINTS_PROMPT_TEMPLATE` + 老师的要点表，**不叠加**学科 prompt：
  两段都写着"评分要点"，叠在一起模型容易把通用口径当成真的要点去判命中。

## 33. 2026-10-30 08:50：答案图片识别 + 门户按科目区分班级里的卷子

### 33.1 需求
> 批改备注里头，要允许教师传入答案的图片，然后返回识别后的结果；
> 由于现在有多个班级，因此改卷完之后的界面，进入班级后，也要显示不同的科目。

两件事：**（A）参考答案拍照就能用**、**（B）一个班里混着多科时要分得清**。

### 33.2 A. 答案图片识别（`POST /api/ocr/answer_key`）
老师手里的参考答案多是打印稿/手写稿，抄一遍要点又慢又容易抄错分值。
入口放在扫描端「📝 本批练习与批改备注」块里（`#ocrWrap`），**识别 ≠ 保存**：

1. 点「📷 上传答案图片识别」选图（最多 8 张）→ 后端逐张视觉识别；
2. 模型只负责**读出来并拆成结构化 JSON**（题号 / 题干 / 参考答案 / 要点[文本,分值]），
   **格式由服务端生成**（`_render_exercise_text`），所以产出一定是 `parse_exercise_body`
   认的格式；让模型去模仿我们的文本格式，十次有三次会错；
3. 结果面板显示：识别了几张、几题几点多少分、不确定项（`notes`），下面是**可编辑**的
   文本框，老师改完再点：
   - 「填入练习正文」→ 打开新建练习、把文本塞进去、**自动跑一次解析预览**；
   - 「填入题目要点」→ 作文学科走这条，塞进「新建题目」的评分要点；
   - 「追加到批改备注」→ 备注上限 300 字，超了直接拦下并提示改用前两个按钮
     （整份答案塞进一句口径只会误导模型）。

实现要点：
- `ANSWER_KEY_EXERCISE_PROMPT` / `ANSWER_KEY_REQUIREMENTS_PROMPT`：分值优先级
  「图上印的 > 每点 N 分展开 > 没标就填 null（不编）」；看不清写 `[?]` 并进 notes；
- `_merge_ocr_questions()`：**多页题号重复就整体重编号**——两张图各自从「1.」开始是常态，
  直接拼会得到两道第 1 题，解析后要点被并到一题里，分数就乱了；
- `_render_exercise_text()`：**只在所有要点都有分值时**才在题号行写总分，
  否则会出现"题头 2 分 / 要点合计 3 分"这种自相矛盾的警告，老师还得回头查是哪一步错了；
- `_strip_score_tail()`：模型抄题干时经常把「（6分）」一起带出来，我们再补一次合计分就成了
  「…（6分）（6分）」。分值统一由我们按要点算，行尾原有的分值标记一律摘掉（真机验证暴露的）；
- 识别完顺手用 `parse_exercise_body` 回读一遍，几十毫秒换来"立刻看到几题几点总分"。

### 33.3 B. 门户按科目区分
一个班同时有英语作文（满分 25/15）和政治卷（满分 50）是常态，混在一起算平均分没有意义。

后端：
- `_subject_counts(records)`：每科 `{key, name, total, graded, pending, kind}`，按份数降序；
- `/api/portal/overview`：顶层 `subjects` + 每个班级卡片 `subjects`；
- `/api/portal/records?subject=<key>`：支持科目筛选，并返回**不含这次 subject 条件**的
  `subjects` 清单（`base_count` 同理）——否则一键进「政治」就只剩政治，切不回别的科目了。

前端：
- 详情页 `#subjectFilter`：只在当前范围确实有多科时出现，选项带份数；
  选中的科目被清空（卷子删了/换班了）会自动回到「全部科目」并重拉，
  避免停在空列表上而切换条又被隐藏；
- 记录卡 `subjectBadge()`：要点判分类橙色、整篇作文类蓝色；
- 班级卡片直接摊开各科份数（`英语作文 12` `政治 / 道法 8`），点进去之前就知道有什么；
- 详情标题：`共 N 份 · 已批 X · 待批 Y` + 各科份数；
- `errStats(list, title)` 改为**按科目分组各算一块**（平均分/要点命中率/未得分要点 Top5），
  否则 25 分制的作文和 50 分制的政治会算进同一个平均分；
- 学生端批次卡也带科目标签（一批 = 一科），整页详情标题同样标出。

### 33.4 验证（2026-10-30）
| 脚本 | 覆盖 | 结果 |
|---|---|---|
| `temp/pttest_test2.py` | 渲染→回读、多页题号重整、作文模式、接口鉴权/空文件/全失败文案 | 全绿 |
| `temp/pttest_test3.py` | `_subject_counts`、混合班级 overview、`?subject=` 筛选、科目+状态叠加 | 全绿 |
| `temp/pttest_test.py` | 要点判分回归（改 `_render_exercise_text` 后重跑） | 全绿 |
| `temp/smoke_subjects.py` | 真实服务只读：overview/records 带 subjects、筛选后份数一致 | 全绿 |
| `smoke_live.py` | 真实服务只读：context / exercises dry_run / 页面 | 全绿 |

真实数据目前只有英语作文（41 份），多科目 UI 的多科形态是在隔离环境里造混合数据验的；
等政治/地理卷真扫进来，班级卡和筛选条会自动出现。服务已重启，横幅显示"密钥保险箱：已加密 2 项"。

**真机验证**：`temp/ocr_real.py` 用 PIL 画了一张中文「参考答案」图（msyh 字体，13 行），
走真实模型识别一次：HTTP 200，2 题 / 5 要点 / 总分 10 全对，题干分值重复的问题当场发现并修掉。

### 33.5 注意
- 答案图识别走的是**视觉模型**，和批卷共用额度：一次最多 8 张，别拿它整本扫描；
- 识别结果只填输入框，**不落盘**；「保存并使用」才写 `exercises.json`；
- 作文学科上传答案图时 `mode=requirements`，产出是一行一条评分要点，
  会填进「新建题目」的要点框，不是练习。

## 34. 2026-10-30 09:15：练习改成可选（扫进去就能批）+ 答案原文直接 AI 整理

起因是两句老师的话：

> 「我其实只需要添加备注，不需要添加『选择本批练习』，我希望扫进去就能做，
> 现在加题目一直是『没解析出题目』，我不知道这是什么。」

> 「批改要点，我只希望老师输入这样的文本：`16表格内填空题 答案【答案】①坚持以人民为中心 ②…`，
> 而不是每个人都要写很严谨的要点。去除『就选一份练习，或点右边「＋ 新建练习」』这种东西。」

一句话：**前置条件太多，格式要求太细**。这两样都把成本压在老师身上，都该去掉。

### 34.1 练习从「必填」变「可选」（两条判分路子）

要点判分类学科（政治/地理）原来卡着「必须先选一份练习」，扫完一叠才报错，老师得把纸
重新放回扫描仪。现在：

| 情形 | 行为 |
|---|---|
| 选了练习 | 逐点判命中，按老师给的要点分值累加（不变） |
| **没选练习** | **按学科标准整体给分 + 评语**（`_vision_grade_free`），不逐点给分 |
| 选过但练习被删 | 仍然报错（`本批绑定的练习已不存在…`）——不能悄悄降级，否则老师以为还在按要点给分 |

改动点：
- `_precheck_auto_batch()` 不再因「没选练习」拦人，只挡「学科压根不能自动批」；
- `_start_auto_batch()`：有练习 → `max_score = exercise_total_max(exercise)`；
  没练习 → `exercise = None`，满分用老师填的（`_coerce_max_score`）；
- `_points_pipeline()` 用 `freeform = exercise is None` 分流（同一套归档、断点续批、表头逻辑）；
- 新增 `_vision_grade_free(page, subject, max_score, note=…, quality, rotation)`：
  调用 `_vision_grade_one`，返回与 `_vision_grade_points` 同构的 dict（`questions: []`、`freeform: True`）；
- 归档：`points = []`、`points_meta = None`、`grading.mode = 'free'`，完成文案
  「已完成 N 份整体判分（满分 X，未用要点表）」；
- 批次公开字段新增 `mode`（`points` / `free`），扫描端据此把阶段文案换成「整体判分」，
  并在完成卡里写明「本批没选要点表，所以只有总分和评语」；
- 扫描端：下拉默认项改成 `— 不用要点表（按学科标准整体判分）—`，`#promptHint` 改成
  「（可选：选了按要点逐条判，不选就整体判分）」，`syncExerciseMaxScore()` 在「从锁住变回可填」
  时把满分还给学科默认值（否则会留着上一份练习的合计分）；
- 门户：`recCardHtml` 在 `grading.mode === 'free'` 时插一条 `.freebar`
  「按 X 标准整体判分（本批未使用要点表，所以只有总分与评语）」，避免老师以为明细没加载出来。

### 34.2 答案原文直接粘，格式由系统（和 AI）兜

老师粘的是答案文档里的句子，不是我们的格式。三层保障：

1. **本地解析器先认**（`parse_exercise_body`）：
   - 新增**圈码要点**：一行里的 `①②③…`（`\u2460-\u2493`）一条一个要点，题号可以紧贴文字
     （`16表格内填空题` → 第 16 题），`【答案】/答案：` 这类标记自动剥掉；
   - 分值写不写都行：不写每点 1 分，且一题多个未标分值的要点**合并成一条警告**
     （原来 4 个要点刷 4 行，老师只会学会无视警告）；
   - 题号归一化 `_norm_qno()`：「第一问」「第1题」「（1）」「一、」统一成 `1`
     （AI 常把题号写成「第一问」，不归一化会出现两道"第一问"的题）。
2. **认不出就自动叫 AI**：`save_exercise_api` 在本地解析为空时调 `_ai_structure_exercise`，
   成功则落盘存**规范文本**（老师下次打开还能接着改），返回 `ai` 一句说明；
3. **手动兜底**：`POST /api/exercises/ai_parse` + 扫描端「AI 整理」按钮，
   给"认出来了但拆得不对"的情况用。

新增：
- `ANSWER_TEXT_STRUCTURE_PROMPT`（文本版，与图片版的 `ANSWER_KEY_EXERCISE_PROMPT` 同 schema）；
- `_structure_answer_text(raw_text, subject)` → 模型只出 JSON；
- `_ai_structure_exercise(raw_text, subject)` → JSON → `_merge_ocr_questions` 归一化 →
  `_render_exercise_text` 生成规范文本 → `parse_exercise_body` 回读校验（走不通就报错，不丢空结果）。

提示词里写死一条：**原文没标分值就让 score 填 null，绝对不许自己编分值**。编出来的分值会
一路带进学生成绩，比"每点 1 分"危险得多。

### 34.3 文案（把"教学口吻"删掉）

- `renderExerciseMeta()` 无练习时只说一句「不选也能扫：按《X》标准整体给分（分数 + 评语）」，
  删掉了「就选一份练习，或点右边「＋ 新建练习」」；
- 练习框占位符换成老师真实文本，格式说明改成「不用改格式：整段答案原文粘进来即可」；
- 「填个示例」填的就是答案原文（带【答案】和 ①②③），让示例本身证明不用改格式；
- 解析失败文案从「没解析出题目，检查一下格式」改成能照做的一句 + 说明「AI 也试过了」。

### 34.4 验证（2026-10-30）
| 脚本 | 覆盖 | 结果 |
|---|---|---|
| `temp/pttest_test.py` | 新增 [5] 不选练习 → 整体判分（freeform 归档字段、满分取老师填的、备注透传）；[6] 练习被删仍报错 | 40 项全绿 |
| `temp/pttest_test2.py` | 答案图 OCR 回归（改解析器后重跑） | 全绿 |
| `temp/pttest_test3.py` | 多科目统计/筛选回归 | 全绿 |
| `temp/parse_check.py` | 老师原话 6 组：一行 ①②③④ / 带分值 / 多题上下排 / 换行一条一点 / 旧格式 / 非答案文本 | 全对 |
| `temp/_ai_real.py` | **真机**：乱格式文本（本地认不出）→ AI 整理 → 2 题 / 4 要点 / 8 分 | 通过 |
| `temp/smoke_live.py` | 真实服务只读：context / exercises dry_run / 页面 / 代码不陈旧 | 全绿 |

真机那次暴露了一个真问题：AI 把题号写成「第一问 / 第二问」，`_Q_NUM_RE` 只认「1. / 一、」，
于是「第二问.（4分）」被当成上一题的要点，解析成 1 题 5 要点（其中一条叫"第二问."）。
修法两条一起上：解析器认 `第X问`，且出口处 `_norm_qno` 统一归一化。

服务已重启（端口 5001），横幅「密钥保险箱：已加密 2 项」；`/station` 含 `exAiBtn`/`exSampleBtn`，
`/portal` 含 `freebar`，旧提示语「需要先选一份练习」在页面里已为 0 处。

## 35. 2026-10-30 11:05：备注默认展开、「暂不开放」取消、英语作文之外统一叫「通用批阅」

老师的原话：

> 「每个老师过来改，基本都是要打入备注的，那个输入框默认展开。
> 『暂不开放』的那部分提示词去掉……除了英语作文外的那种批阅，其他都叫通用批阅，
> 默认需要老师输入批改备注的，但是不需要表明『本批题目』之类的多余信息，以便捷老师们的操作。」

### 35.1 批改备注默认展开，且排在最前

- `<details id="promptBox" open>`：老师每次来几乎都要写一句口径，藏起来等于让他多点一次；
- 备注 textarea 从折叠区**底部**挪到**顶部**（题目/答案要点这些选配项挪到下面），
  rows 2 → 3，placeholder 直接给三种真实写法；
- 删掉 `<span id="promptMode">` 与标题里的「本批题目 / 本批练习」，summary 只剩「✍️ 批改备注」，
  右边的提示随学科换：作文「（写一句这次怎么判；不写就按卷面识别题干）」、
  通用批阅「（写一句这次怎么给分；不写就按学科标准整体给分）」；
- 删掉折叠区底部的说明「题目/评分标准与备注跟着这一批卷子一起送去批阅，优先级高于卷面上印的要求」；
- 备注**不设硬性拦截**：它是"默认要填的东西"，不是"不填就不让扫"。空着按学科标准批，
  真设成必填只会逼老师写"无"。

### 35.2 「暂不开放」不存在了：英语作文之外一律走通用批阅

以前 `subject_kind()` 里数学 / 默写 / 语文作文是 `generic`（只能逐份手动批），
扫描端因此挂一条「暂不开放：数学、默写…」的橙条，下拉里也把它们过滤掉。

现在改成**按流水线路由**而不是按配置里的原始 kind：

```python
# 配置里有这个学科 → 是 essay 就走作文，其余一律 points（= 通用批阅）
# 配置里压根没有 → 仍按 generic 拒绝（不能拿个不存在的 key 去批卷）
return ESSAY if val.get('kind') == 'essay' else POINTS
```

连带改动：
- `_run_pipeline`：`kind in AUTO_PIPELINE_KINDS` 就进 `_points_pipeline`（原来只认 `== POINTS`）；
- `grading_subject_options()`：对外暴露**实际路由 kind**，`need_exercise` 恒为 False
  （答案要点已改成可选），前端 `isPointsSubject()` 因此对数学/默写也成立；
- `save_exercise_api`：只拒绝英语作文（作文的评分要求写进「题目」），其余学科都能建答案要点；
- `auto_pipeline_reject_message`：去掉"目前接了作文与要点判分两类"的说法；
- 扫描端删掉 `#subjNote` 的「暂不开放」渲染（元素留着，恒隐藏，老页面缓存也不会报错）；
- 门户/扫描端文案统一成「通用批阅」：「📌 通用批阅 · 逐点得分」、
  「通用批阅：按《X》标准整体给分（没给答案要点，所以只有总分和评语）」；
- 后台首页里指路的「📝 本批题目与批改备注」→「✍️ 批改备注」。

### 35.3 顺手修的坑：满分框被整体禁用

`syncMaxEnabled()` 原来按 `isPointsSubject()` 一刀切禁用满分输入——那是"要点判分必须用练习总分"
时代的规定。现在没选答案要点时满分该由老师决定，禁用会让 freeform 批次永远只能用学科默认满分。
改成 `locked = isPointsSubject() && !!currentExerciseSummary()`（选了要点表才锁），
并在 `#exerciseSel.onchange` 里同步调用。

### 35.4 验证
| 脚本 | 覆盖 | 结果 |
|---|---|---|
| `temp/pttest_test.py` | 新增 [7]：数学/默写/语文作文都走通用批阅、下拉无 `auto_ok=false`、数学批次 free 归档、数学能建答案要点、英语作文仍拒要点表 | 全绿（48 项） |
| `temp/pttest_test2.py` / `test3.py` | OCR、多科目统计回归 | 全绿 |
| `temp/smoke_live.py` | 真实服务：`/station` 200、`id="promptBox" open`、标题「✍️ 批改备注」、「暂不开放」0 处、代码不陈旧 | 全绿 |

服务已重启（5001）。`/station` 返回的 HTML 里 `id="promptBox" open` 1 处、`✍️ 批改备注` 1 处，
「本批题目 / 本批练习」只剩注释里的历史说明。
