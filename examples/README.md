# 第 6 章 · 可运行示例

### 1. 健康检查

```bash
curl http://localhost:5001/
```

### 2. 招标文件信息识别

```bash
curl -X POST http://localhost:5001/recognize_tender_info \
  -F "file=@samples/需求论证报告.pdf"
```

### 3. 招标文件生成与合规审查

```bash
curl -X POST http://localhost:5001/generate_tender \
  -H "Content-Type: application/json" \
  -d '{"tender_info": {...}, "correction_hints": ""}'

curl -X POST http://localhost:5001/review_tender \
  -H "Content-Type: application/json" \
  -d '{"tender_content": "...", "company_name": "示例采购单位"}'
```


> 请把 `samples/` 中的示例文件替换为自己的测试数据；示例数据不含任何真实业务信息。
