# 借鉴记录：文风词表（style）

- 模块范围：`shared/style/ai-tone.md`、`shared/style/defensive.md`、`shared/scripts/style_patterns.json`
- 阅读方式：2026-10-16 通过 GitHub API 在线读取，未克隆、未下载到本仓库。下表"读取版本"为当日默认分支的提交号。
- 借鉴方式分三种：**思路**（分类方法、判定原则，不涉及文字）；**改写**（参考原文的规则或例句，用自己的公文语境例句重写）；**数据**（引用原项目公开的统计结论，用于定级别和阈值）。本模块所有例句均为自拟，未整段照录任何来源。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的规则 / 文件 |
|---|---|---|---|---|---|
| [larashero3-dotcom/lieflat-less-ai-tone](https://github.com/larashero3-dotcom/lieflat-less-ai-tone) | MIT | 27d2923 | 11 项有量化证据的中文 AI 特征（段首零主语评论 4.4 倍、翻案腔 3.4 倍、破折号 3.0 倍、提示性冒号 3.8 倍、空转句引列表 9.4 倍、拟人化喻体 7.3 倍、相邻句同构 2 倍、数字密度人类为 AI 的 2.8 倍）；翻译腔从 18 种筛到 5 种的结论；"触发标记可定位才立规则""冒号按用途拆分""句中总说分说冒号不改""句内排比不改、只管句间"等判定原则；`check-structure.py` 用逗号数和长度做句子指纹的思路 | 数据、思路、改写 | AI-001、003～007、010～021、023、044、045、049、051，DEF-010；ai-tone.md 第一部分（一）～（七）（十二）、白名单第 7、13 条 |
| [op7418/Humanizer-zh](https://github.com/op7418/Humanizer-zh) | MIT | f4518a8 | 中文规则骨架（铺垫、公式化节奏、拔高、排版、残留五组）；"四字格、被动句、名词化在商务和学术文体中可以保留"的边界；§15 句尾拔高、§30"随着……的发展"、§31 套话收尾的判定 | 思路、改写 | AI-009、026、028、032、036、043、044、047；ai-tone.md 白名单第 4、13 条 |
| [LifelongLazyLearner/qu-ai-wei](https://github.com/LifelongLazyLearner/qu-ai-wei) | MIT | 1d32e80 | 每类特征都写"什么时候保留"的组织方式；README 明确"公文、条款、步骤需要平行结构"；"限定词和安全垫串联"只删重复限定、保留证据等级；"假揭示"（看似……本质） | 思路 | AI-007、028、040，DEF-013；ai-tone.md 第二部分整体结构 |
| [MrGeDiao/shuorenhua](https://github.com/MrGeDiao/shuorenhua) | MIT | e7c2b86 | 保持事实、责任主体、限定和判断强度不变的改写约束 | 思路 | ai-tone.md 使用原则第 3 条；defensive.md 第八部分 |
| [VincentOld/stop-slop-zh](https://github.com/VincentOld/stop-slop-zh) | MIT | 42305b6 | 按"开场套话、强调虚壳、黑话、元评论、鸡汤收尾、比喻库存、三件套排比"分组的词表组织方式；各组的候选词 | 思路、改写（逐词筛选） | AI-001、002、009、019、022、025、026、029～031、033～036、041～043、046、050 |
| [blader/humanizer](https://github.com/blader/humanizer) | MIT | 225a6f3 | §9 限定词堆叠（保留适用范围、法律和安全提示）；§22 聊天残留；§23 知识边界免责与猜测填充（写明来源未记载，或删去） | 思路 | DEF-001、003、006、008、010、013、025；defensive.md 第三、八部分 |
| [hexiecs/talk-normal](https://github.com/hexiecs/talk-normal) | MIT | d89cf32 | 正反两种否定对举（不是 X 而是 Y / X 而不是 Y）；总结标签（总而言之、一句话总结）；比较类问题直接给推荐、不写平衡论述 | 思路 | AI-004、006、008、038、039，DEF-011、012、016、023 |
| [molang163/chinese-natural-reply](https://github.com/molang163/chinese-natural-reply) | MIT | 8f4ea91 | 中文防御性回复的 10 类毛病，其中"划边界""提前纠正对方没有的误解""汇报没动什么"迁移到公文场景；"没把握又影响对方行动的，说一次；不影响的，不说"的判定 | 思路 | DEF-001、004、007、009、016、026、027；defensive.md 第三部分 |
| [MohamedAbdallah-14/unslop](https://github.com/MohamedAbdallah-14/unslop) | MIT | 29d4360 | 三档严重度（单次即报 / 组合才报 / 过密才报），对应本模块"P1 单次提示、P2 加密度阈值"的分级；"Whether you're X, Y, or Z"平行假设句 | 思路 | 全部 `density_per_1000` 规则的分级思路；AI-049，DEF-003、011 |

## 二、只学思路的来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 处理方式 | 影响的规则 |
|---|---|---|---|---|---|
| [larashero3-dotcom/lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen) | PolyForm Noncommercial 1.0.0 | e0a5aba | 其公开的真实公文语料统计结论：顿号密度每千字约 14～31 个、破折号极少、冒号稀少、"进一步"为最高频副词、"必须/应当"每篇只有几处 | 只参考统计结论来校准阈值和白名单方向，未复制任何文字、表格或代码；阈值数值由本项目自行确定 | AI-010、051 的阈值，DEF-019、028 的阈值；ai-tone.md 白名单第 8 条；defensive.md 第六部分"必须要省着用" |

## 三、有意没有采用的内容

以下内容在来源项目中是去 AI 味规则，但会误伤规范公文，本模块不收录或改为白名单：

1. stop-slop-zh 把"总分总结构""首先……其次……最后""一方面……另一方面"列为 AI 套路：在公文中属于正当结构，列入白名单或不设规则。
2. stop-slop-zh 的"高频 AI 成语"（与时俱进、攻坚克难、砥砺前行等）和"比喻库存"中的基石、引擎、桥梁：多为公文常用语或政治性提法，不收录。
3. lieflat-less-ai-tone 第 6 条"序数词当小标题"：公文层次序号由 GB/T 9704-2012 规定，反列入白名单（WL-002）。
4. lieflat-less-ai-tone 第 2 条"顿号罗列过密"的原始阈值：公文顿号密度本来就高，改为每千字 35 个，并豁免实施方案、工作要点、计划。
5. 各来源中的 Markdown、emoji、AI 身份自述、对话引导语：归 `lint_rules.json` 的 residue 类，本模块不收录。
