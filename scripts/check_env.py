#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检查 .env 是否配置完整。只输出字段名与状态，绝不打印密钥值。"""
import os
import re
import sys

ENV_FILE = ".env"
REQUIRED = ['FLASK_SECRET_KEY', 'DEEPSEEK_API_KEY']
OPTIONAL = ['DIFY_COMPLIANCE_API_KEY', 'BAIDU_OCR_API_KEY']
PLACEHOLDER = re.compile(
    r"your|please-generate|change-me|dev-only|你的|示例|xxx|placeholder|<[^>]+>", re.I)


def load(path):
    if not os.path.exists(path):
        return None
    data = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            data[k.strip()] = v.strip().strip('"').strip("'")
    return data


def main():
    data = load(ENV_FILE)
    if data is None:
        print(f"[X] 未找到 {ENV_FILE}，请先执行：cp .env.example .env")
        return 1
    missing = [k for k in REQUIRED if not data.get(k)]
    fake = [k for k in REQUIRED
            if data.get(k) and PLACEHOLDER.search(data[k])]
    print(f"== 密钥配置检查（{ENV_FILE}，共 {len(data)} 项）==")
    for k in REQUIRED:
        if k in missing:
            state = "缺失"
        elif k in fake:
            state = "仍是占位符"
        else:
            state = "已配置"
        print(f"  [{'+' if state == '已配置' else '-'}] {k}: {state}")
    for k in OPTIONAL:
        print(f"  [~] {k}: {'已配置' if data.get(k) and not PLACEHOLDER.search(data[k]) else '未配置/可选'}")
    if missing or fake:
        print("\n[X] 配置未通过：请补齐必填项并替换全部占位符，详见 docs/密钥配置指南.md")
        return 1
    print("\n[OK] 必填项已配置。安全提醒：{ENV_FILE} 已被 .gitignore 排除，切勿提交、截图或粘贴到聊天/文档中。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
