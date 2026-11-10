# tender-doc-agent · 采购文件编制与合规审查

> 《协同采购：人工智能时代政府采购治理变革》**第 6 章**配套开源项目。
> 五章对应五个独立仓库，本仓库聚焦「采购文件编制与合规审查」场景，clone 后即可按下方步骤复现该章实践案例。

Flask + DeepSeek + Dify 工作流主备双路由，集成 PyMuPDF 与百度 OCR；需求结构化提取、招标文件流式生成、字段级防幻觉校验、四块 42 项合规审查与 25 项避错专项审查。

- **技术栈**：Python 3.9+ / Flask / PyMuPDF / python-docx / Dify
- **服务端口**：5001（Flask Web）
- **许可证**：MIT（见 `LICENSE`）

## 场景能力

1. PDF/DOCX/TXT 与 300DPI 扫描件的混合文本提取
2. 按 JSON Schema 抽取需求字段并逐项回指原文（check_field_in_source）
3. 模板骨架 + 字段填充 + 章节级生成 + 全局一致性校验
4. 关键字段正则防幻觉与审查结果行号锚点定位
5. 67 项采购检查与分析维度中的 42 项合规 + 25 项避错规则

## 快速开始

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env                       # 填入 DEEPSEEK / Dify / 百度 OCR 密钥
python app.py                              # http://localhost:5001
```

> 所有 API 密钥、数据库口令与签名密钥均通过 `.env` 注入，源码不含真实凭据。
> 复制 `.env.example` 为 `.env` 后填写；`.env` 已被 `.gitignore` 排除，详见 `SECURITY.md`。
>
> **密钥配置（部署前必读）**：完整申请入口、字段表与校验方法见 [docs/密钥配置指南.md](docs/密钥配置指南.md)；
> 复制模板填值后运行 `python3 scripts/check_env.py` 校验，返回 `[OK]` 再启动服务。
> 绝不把 `.env` 提交到 Git，也不要把密钥回填进源码，详见 `SECURITY.md`。

## 本章核心模块

| 路径 | 作用 |
|---|---|
| `app.py` | /recognize_tender_info、/generate_tender、/review_tender |
| `app.py::_validate_hallucination` | 关键字段防幻觉校验 |
| `app.py::_sanitize_review_result` | 审查结果行号锚点净化 |
| `docs/67项采购检查与分析维度完整清单.csv` | 四块审查与避错规则明细 |
| `bid_collusion_workflow.yml` | Dify 工作流导入样例（第7章亦使用） |

完整流程、接口参数与演练步骤见 `docs/场景指南.md`；可直接运行的调用示例见 `examples/`。

## 完整部署文档

- `docs/场景指南.md`：本章场景的操作流程、接口示例与验证点
- `docs/`：安装说明、架构说明、规则清单与系统截图说明
- 原系统架构文档位于 `docs/ARCHITECTURE.md`（若适用）

## 安全与数据

请勿把 `.env`、运行日志、业务数据库或含真实供应商信息的文件提交到公开仓库。
涉及合同、发票、投标文件的运行数据应置于校内网络或本地环境。

---

*本项目由本书团队以 MIT 许可证公开发布，欢迎 Issue 与 Pull Request。*
