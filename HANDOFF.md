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
| 作文提示词 | `ESSAY_PROMPT_TEMPLATE` | 6325 |
| 语料提示词 | `CORPUS_PROMPT` | 6392 |
| 佳句提示词 | `AWARD_PROMPT` | 6405 |
| 送图压缩 | `_img_to_data_url(max_size=2400)` | 6425 |
| 卷面体检（本地） | `image_quality_metrics()` / `quality_badge()` | 6454 |
| 字迹体检（模型） | `assess_handwriting()` + `HANDWRITING_PROMPT` | 6575 |
| 结果归一化 | `_norm_essay_result()`（含 `illegible` 过滤） | 6725 |
| 作文批阅 | `_vision_grade_essay(page, requirements, max_score, model, quality)` | 6838 |
| 语料缓存 | `_get_topic_corpus()`（精确指纹 + Jaccard≥0.55 模糊复用） | 6981 |
| 佳句评选 | `_pick_awards()` | 7030 |
| 自动流水线 | `_auto_pipeline()` / `_start_auto_batch()` | 7090 / 7240 |
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
`job{status,stage,done,total,message}`、`topic{title,requirements,key}`、`corpus{phrases,sentences,patterns,pitfalls}`、
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

---

## 13. 交接前最后确认

- [x] `git add -A && git commit`（2026-10-26 晚完成：`a46d35c` 主链路 + `d9f91b7` 名单工具/测试）
- [ ] `.env` 里是真密钥，确认没被提交（`.gitignore` 已忽略，已核验）
- [ ] 新机器上先跑 `python app.py`，打开 `/station` 看临时区是否空、体检徽章是否正常
- [ ] 有扫描仪的那台机器才测得了扫描链路，其他机器只能测导入图片

> ⚠️ **待办**：`app.py` 里本轮学号逻辑的改动与另一并发会话（学生报告 v2）的改动交错，
> 尚未提交。接手后先确认两者都完整，再一起提交。

---

## 14. 2026-10-26 晚：名单换成真实数据 + 学号口径重构（重要）

### 14.1 名单已导入（取代原先 6 个测试账号）

用 `tools/import_roster.py` 导入学校下发的《学生姓名与学号对照表.xlsx》：
**1850 名学生 / 41 个班（26xx 高一 14 班、25xx 高二 14 班、24xx 高三 13 班）**。
学生账号（密码=学号）、`data/students/<班>.json`、`data/classes/<班>/` 三者已对齐重建；
`admin` 保留；旧测试账号（1402/1412/1425/1431/1440/230205）与其班级 14、2302 已清除；
失效的班级归属（2302）已清理。账号表备份在 `data/config/users.json.bak-*`。

⚠️ **旧档案已清理（2026-10-26 20:37）**：`data/records` 里原 7 条旧测试记录
（学号 1431/1412/1402/1440/1425/230205，学号 230205 为两条）及其图片、
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

