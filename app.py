from flask import Flask, render_template, request, jsonify, redirect, url_for, flash
import requests
import base64
import logging
import docx
import PyPDF2
import os
import json
import re  # 正则表达式模块，用于行号/文本提取
import uuid
from datetime import datetime  # 新增：用于记录项目创建时间
from io import BytesIO
import fitz  # pymupdf

# ===================== 基础配置 =====================
# 配置日志（优化格式，输出到文件+控制台）
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tender_service.log', encoding='utf-8'),  # 日志文件
        logging.StreamHandler()  # 控制台输出
    ]
)
logger = logging.getLogger(__name__)

# 初始化Flask应用
app = Flask(__name__)
# 新增：设置Flask的secret_key（用于flash消息提示）
app.secret_key = 'change-me-flask-secret'

# ===================== 配置（全部从环境变量读取，切勿硬编码密钥） =====================
# 复制 .env.example 为 .env 后填写；生产环境建议用密钥管理服务注入。
def _load_dotenv(path=".env"):
    """无第三方依赖的 .env 加载器；已存在的环境变量优先。"""
    if not os.path.exists(path):
        return
    for line in open(path, encoding="utf-8"):
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

_load_dotenv()

def env(name, default=""):
    return os.environ.get(name, default)

# Dify 五个工作流（信息识别 / 招标生成 / 合规审查 / 合同评审 / 围标串标）
DIFY_RECOGNIZE_API_KEY     = env("DIFY_RECOGNIZE_API_KEY")
DIFY_GENERATE_API_KEY      = env("DIFY_GENERATE_API_KEY")
DIFY_COMPLIANCE_API_KEY    = env("DIFY_COMPLIANCE_API_KEY")
DIFY_ARGUMENT_API_KEY      = env("DIFY_ARGUMENT_API_KEY")
DIFY_BID_COLLUSION_API_KEY = env("DIFY_BID_COLLUSION_API_KEY")
DIFY_BASE_URL              = env("DIFY_BASE_URL", "https://api.dify.ai/v1").rstrip("/")
DIFY_RECOGNIZE_URL         = DIFY_BASE_URL + "/workflows/run"
DIFY_GENERATE_URL          = DIFY_BASE_URL + "/workflows/run"
DIFY_COMPLIANCE_URL        = DIFY_BASE_URL + "/workflows/run"
DIFY_ARGUMENT_URL          = DIFY_BASE_URL + "/workflows/run"
DIFY_BID_COLLUSION_URL     = DIFY_BASE_URL + "/workflows/run"
DIFY_PROXIES               = {"http": None, "https": None}

# DeepSeek 直连与百度 OCR
DEEPSEEK_API_KEY      = env("DEEPSEEK_API_KEY")
DEEPSEEK_BASE_URL     = env("DEEPSEEK_BASE_URL", "https://api.deepseek.com/chat/completions")
BAIDU_OCR_API_KEY     = env("BAIDU_OCR_API_KEY")
BAIDU_OCR_SECRET_KEY  = env("BAIDU_OCR_SECRET_KEY")

# 可选的外部 AI Agent 服务
AI_AGENT_BASE_URL     = env("AI_AGENT_BASE_URL", "http://127.0.0.1:8000")

# Flask 会话密钥（放在 .env 加载之后，避免重启后会话失效）
app.secret_key = env("FLASK_SECRET_KEY") or os.urandom(24).hex()

# 支持的文件类型（新增.txt）
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'txt'}
# 默认招标公司名称（传给Dify的{{company_name}}参数）
DEFAULT_COMPANY_NAME = env("DEFAULT_COMPANY_NAME", "示例采购单位")

# 新增：项目记录存储（开发环境用内存，生产环境建议用数据库）
# 扩展项目结构，增加 data 字段存储内容
projects = []


# ===================== 新增：项目记录管理工具函数 =====================
def save_project(project_type, project_name, mode, created_at):
    """
    保存项目记录到内存列表
    :param project_type: 模块名称（如tender_generate）
    :param project_name: 项目名称
    :param mode: 功能模式（generate/compliance）
    :param created_at: 创建时间字符串
    :return: 项目ID
    """
    project_id = len(projects) + 1
    projects.append({
        "id": project_id,
        "type": project_type,
        "name": project_name,
        "mode": mode,
        "created_at": created_at,
        "data": {}  # 新增：存储项目数据
    })
    return project_id


def update_project_data(project_id, data):
    """更新项目数据（修复：兼容字符串ID）"""
    # 统一转换为字符串比较
    project_id_str = str(project_id)
    for project in projects:
        if str(project['id']) == project_id_str:
            project['data'] = data
            logger.info(f"保存项目 {project_id} 内容，长度：{len(str(data)) if data else 0}")
            return True
    return False


def get_project_data(project_id):
    """获取项目数据（修复：兼容字符串ID）"""
    project_id_str = str(project_id)
    for project in projects:
        if str(project['id']) == project_id_str:
            return project['data']
    return {}


# ===================== 核心工具函数（保留你原有所有代码，仅修改extract_file_content） =====================
def clean_empty_fields(data):
    """
    递归清理字典/列表中的空值字段
    - 清理范围：空字符串""、None、空字典{}、空列表[]
    - 用途：仅保留有效数据，减少前端渲染冗余
    """
    if isinstance(data, dict):
        cleaned_dict = {}
        for key, value in data.items():
            cleaned_value = clean_empty_fields(value)
            if cleaned_value not in ["", None, {}, []]:
                cleaned_dict[key] = cleaned_value
        return cleaned_dict
    elif isinstance(data, list):
        cleaned_list = [clean_empty_fields(item) for item in data if item not in ["", None, {}, []]]
        return cleaned_list
    else:
        return data


def _sanitize_recognize_result(parsed_data, original_text):
    """
    对信息识别结果做后处理校验，清除明显幻觉字段：
    1. 电话号码：必须在原文中出现，且归属字段正确（招标人/代理机构不混填）
    2. 日期：必须在原文中出现
    3. 金额：必须在原文中出现
    4. 招标编号：必须在原文中出现
    返回 (cleaned_data, issues列表)
    """
    issues = []

    def check_field_in_source(section, field, value):
        """检查字段值是否能在原文中找到"""
        if not value or not isinstance(value, str):
            return True
        # 取值的核心部分（去掉括号说明等）做匹配
        core = value.split('（')[0].split('(')[0].strip()
        if len(core) >= 4 and core not in original_text:
            return False
        return True

    # 1. 检查电话号码字段
    phone_pattern = re.compile(r'1[3-9]\d{9}|0\d{2,3}[-\s]\d{7,8}|\d{11}')
    contact = parsed_data.get("联系方式", )

    # 收集原文中所有电话
    phones_in_text = set(phone_pattern.findall(original_text))

    for field in ["招标人联系电话", "代理机构联系人"]:
        val = contact.get(field, "")
        if val:
            found_phones = set(phone_pattern.findall(val))
            fake_phones = found_phones - phones_in_text
            if fake_phones:
                issues.append(f"联系方式.{field} 含原文不存在的电话：{fake_phones}")
                contact[field] = ""

    # 2. 检查日期字段
    date_pattern = re.compile(r'\d{4}年\d{1,2}月\d{1,2}日|\d{4}-\d{2}-\d{2}')
    date_fields = [
        ("文件获取信息", "获取时间"),
        ("投标文件递交信息", "投标截止时间"),
    ]
    for section_name, field_name in date_fields:
        section = parsed_data.get(section_name, {})
        val = section.get(field_name, "")
        if val:
            dates_in_val = set(date_pattern.findall(val))
            dates_in_text = set(date_pattern.findall(original_text))
            fake_dates = dates_in_val - dates_in_text
            if fake_dates:
                issues.append(f"{section_name}.{field_name} 含原文不存在的日期：{fake_dates}")
                section[field_name] = ""

    # 3. 检查金额字段
    core_info = parsed_data.get("基础核心信息", {})
    for field in ["预算金额", "资金来源"]:
        val = core_info.get(field, "")
        if val:
            # 提取数字部分做校验
            nums = re.findall(r'\d+', val)
            for num in nums:
                if len(num) >= 4 and num not in original_text:
                    issues.append(f"基础核心信息.{field} 含原文不存在的数字：{num}")
                    core_info[field] = ""
                    break

    # 4. 检查招标编号
    bid_no = core_info.get("招标编号", "")
    if bid_no and bid_no not in original_text:
        issues.append(f"基础核心信息.招标编号 '{bid_no}' 在原文中不存在")
        core_info["招标编号"] = ""

    if issues:
        logger.warning(f"信息识别后处理发现 {len(issues)} 个疑似幻觉字段已清空：{issues}")

    return parsed_data, issues


def extract_file_content(file):
    """
    提取PDF/Word/TXT文件的文本内容（修改版：返回纯文本+带行号文本）
    :param file: Flask上传的文件对象
    :return: 提取的纯文本内容, 带行号的文本 / 错误提示字符串
    """
    filename = file.filename
    file_ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''

    try:
        if file_ext == 'pdf':
            # 解析PDF（PyPDF2）
            pdf_reader = PyPDF2.PdfReader(file)
            text = ''
            for page in pdf_reader.pages:
                page_text = page.extract_text() or ''
                text += page_text
            # 统一换行符，保留所有空行
            text = text.replace('\r\n', '\n').replace('\r', '\n')
            lines = text.split('\n')
            # 生成带行号的文本（保留所有行，包括空行）
            text_with_line = ''
            for idx, line in enumerate(lines, 1):
                text_with_line += f"【行{idx}】{line}\n"
            return text, text_with_line

        elif file_ext == 'docx':
            # 解析docx（python-docx）
            doc = docx.Document(file)
            text = '\n'.join([para.text for para in doc.paragraphs])
            # 统一换行符，保留所有空行
            text = text.replace('\r\n', '\n').replace('\r', '\n')
            lines = text.split('\n')
            # 生成带行号的文本（保留所有行，包括空行）
            text_with_line = ''
            for idx, line in enumerate(lines, 1):
                text_with_line += f"【行{idx}】{line}\n"
            return text, text_with_line

        elif file_ext == 'doc':
            # .doc格式需额外依赖（如python-docx2txt/antiword），暂不支持
            return "暂不支持.doc格式，请转换为.docx或PDF后上传", None

        elif file_ext == 'txt':
            # 新增：解析TXT文件
            text = file.read().decode('utf-8', errors='ignore')
            # 统一换行符，保留所有空行
            text = text.replace('\r\n', '\n').replace('\r', '\n')
            lines = text.split('\n')
            # 生成带行号的文本（保留所有行，包括空行）
            text_with_line = ''
            for idx, line in enumerate(lines, 1):
                text_with_line += f"【行{idx}】{line}\n"
            return text, text_with_line

        else:
            return f"不支持的文件类型：{file_ext}（仅支持PDF/Word/TXT）", None

    except Exception as e:
        logger.error(f"文件解析失败：{str(e)}，文件名：{filename}")
        return f"文件解析失败：{str(e)}", None


def format_tender_info(data):
    """
    将结构化字典转换为LLM易解析的键值对字符串
    :param data: 识别后的招标文件结构化数据
    :return: 格式化后的文本字符串
    """
    info_str = "以下是完整的招标文件结构化信息，请严格按照这些信息生成招标文件：\n\n"

    def recursive_format(obj, prefix=""):
        nonlocal info_str
        if isinstance(obj, dict):
            for key, value in obj.items():
                if isinstance(value, dict):
                    info_str += f"{prefix}【{key}】\n"
                    recursive_format(value, prefix + "  ")
                else:
                    value = value if value is not None and value != "" else "无"
                    info_str += f"{prefix}{key}：{value}\n"
        elif isinstance(obj, list):
            info_str += f"{prefix}列表内容：{obj}\n"
        else:
            info_str += f"{prefix}内容：{obj}\n"

    recursive_format(data)
    return info_str


# ===================== 核心接口（保留你原有所有代码） =====================
def _call_deepseek_extract(content: str) -> dict:
    """
    直接调 DeepSeek 提取招标信息，严格禁止推断/编造。
    返回解析后的 dict，失败抛异常。
    """
    system_prompt = """你是招标文件信息提取助手。

【铁律，违反则输出无效】
1. 所有字段的值必须逐字来自下方原文，禁止推断、补全、联想。
2. 原文中找不到的字段，值必须填空字符串 ""，绝对不能编造。
3. 禁止把其他项目、其他地区、其他公司的信息填入。
4. 数字、金额、日期、地址、联系方式必须与原文完全一致，一个字不能改。
5. 只输出 JSON，不输出任何解释文字。

【字段映射说明，必须遵守】
- "招标人名称" = 原文"采购人信息"或"招标人"下的"名称"字段
- "招标人地址" = 原文"采购人信息"下的"地址"字段
- "招标人联系人" = 原文"项目联系方式"下的"项目联系人"字段
- "招标人联系电话" = 原文"采购人信息"下的"联系方式"字段（注意区分采购人和代理机构）
- "代理机构名称" = 原文"采购代理机构信息"下的"名称"字段
- "代理机构地址" = 原文"采购代理机构信息"下的"地址"字段
- "代理机构联系电话" = 原文"采购代理机构信息"下的"联系方式"字段
- "代理机构电子邮件" = 原文"采购代理机构信息"下的邮箱地址
- "采购方式" = 原文标题或正文中出现的"公开招标"/"竞争性谈判"等字样
- "开标时间" = 原文"开标时间"字段，与投标截止时间可能相同

输出格式（严格按此结构，找不到的字段填""）：
{
  "基础核心信息": {
    "项目名称": "",
    "招标编号": "",
    "预算编号": "",
    "采购方式": "",
    "项目所在地区": "",
    "资金来源": "",
    "预算金额": "",
    "最高限价": "",
    "项目实施地点": "",
    "采购需求概述": "",
    "合同履约期限": "",
    "是否接受联合体投标": ""
  },
  "采购标的详细信息": {
    "标的名称": "",
    "采购数量": "",
    "简要技术规格": "",
    "是否允许进口产品": "",
    "产品交货要求": ""
  },
  "投标人资格要求": {
    "基础资格": "",
    "专项资质要求": "",
    "信用记录要求": "",
    "关联关系限制": "",
    "分支机构要求": "",
    "行贿犯罪记录": ""
  },
  "文件获取信息": {
    "获取时间": "",
    "获取地点": "",
    "获取方式": "",
    "招标文件售价": "",
    "报名资料清单": "",
    "发票开具": ""
  },
  "投标文件递交信息": {
    "投标截止时间": "",
    "投标地点": "",
    "递交方式": "",
    "投标文件要求": "",
    "开标时间": "",
    "开标地点": "",
    "放弃投标通知": ""
  },
  "联系方式": {
    "招标人名称": "",
    "招标人地址": "",
    "招标人联系人": "",
    "招标人联系电话": "",
    "招标人电子邮件": "",
    "代理机构名称": "",
    "代理机构地址": "",
    "代理机构联系人": "",
    "代理机构联系电话": "",
    "代理机构电子邮件": ""
  },
  "其他法定必填信息": {
    "公告发布媒介": "",
    "公告期限": "",
    "监督部门": "",
    "政策落实要求": ""
  }
}"""

    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"请从以下招标文件原文中提取信息：\n\n{content}"}
        ],
        "temperature": 0.0,
        "stream": False
    }
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    resp = requests.post(DEEPSEEK_BASE_URL, headers=headers, json=payload, timeout=120)
    resp.raise_for_status()
    raw = resp.json()["choices"][0]["message"]["content"]

    # 去掉 <think> 标签和 markdown 代码块
    raw = re.sub(r'<think>.*?</think>', '', raw, flags=re.DOTALL).strip()
    raw = re.sub(r'^```(?:json)?\s*', '', raw).strip()
    raw = re.sub(r'\s*```$', '', raw).strip()

    return json.loads(raw)


@app.route('/recognize_tender_info', methods=['POST'])
def recognize_tender_info():
    """
    招标文件信息识别接口
    - 支持文本/文件输入
    - 优先调 DeepSeek 直连（严格提取，无幻觉），失败降级 Dify
    """
    recognize_result = None
    try:
        # 1. 获取输入数据（文本/文件）
        text = request.form.get('text', '').strip()
        file = request.files.get('file')
        content = text

        # 2. 提取文件内容（如有文件）
        if file and file.filename != '':
            file_content_result = extract_file_content(file)
            if isinstance(file_content_result, tuple):
                file_content = file_content_result[0]
            else:
                file_content = file_content_result

            if file_content.startswith(('不支持', '文件解析失败')):
                return jsonify({"success": False, "error": file_content})
            content = f"{file_content}\n{text}" if text else file_content

        if not content:
            return jsonify({"success": False, "error": "未获取到可识别的文本内容"})

        # 3. 优先调 DeepSeek 直连提取
        parsed_data = None
        try:
            logger.info("开始调用 DeepSeek 直连信息识别")
            parsed_data = _call_deepseek_extract(content)
            logger.info("DeepSeek 信息识别成功")
        except Exception as e:
            logger.warning(f"DeepSeek 识别失败，降级 Dify：{e}")

        # 4. DeepSeek 失败则降级 Dify
        if parsed_data is None:
            headers_dify = {
                "Authorization": f"Bearer {DIFY_RECOGNIZE_API_KEY}",
                "Content-Type": "application/json"
            }
            dify_data = {
                "inputs": {"tender_text": content},
                "response_mode": "blocking",
                "user": f"flask_client_{uuid.uuid4().hex[:8]}"
            }
            logger.info("降级调用 Dify 信息识别工作流")
            response = requests.post(
                DIFY_RECOGNIZE_URL, headers=headers_dify, json=dify_data,
                timeout=120, proxies=DIFY_PROXIES
            )
            response.raise_for_status()
            result = response.json()
            recognize_result = result.get("data", {}).get("outputs", {}).get("recognized_data")
            if not recognize_result:
                return jsonify({"success": False, "error": "识别服务未返回结果"})
            recognize_result = recognize_result.strip().strip('`').strip('json').strip()
            parsed_data = json.loads(recognize_result)

        # 5. 后处理校验 + 清理空字段
        parsed_data, sanitize_issues = _sanitize_recognize_result(parsed_data, content)
        cleaned_data = clean_empty_fields(parsed_data)

        logger.info(f"信息识别完成，有效数据长度：{len(str(cleaned_data))}")
        return jsonify({"success": True, "data": cleaned_data})

    except requests.exceptions.HTTPError as e:
        error_msg = f"请求失败（状态码：{e.response.status_code}）：{e.response.text[:500]}"
        logger.error(error_msg)
        return jsonify({"success": False, "error": error_msg})
    except json.JSONDecodeError as e:
        raw_preview = recognize_result[:500] if recognize_result else "无"
        error_msg = f"识别结果解析失败（JSON格式错误）：{str(e)}"
        logger.error(f"{error_msg}，原始内容：{raw_preview}")
        return jsonify({"success": False, "error": error_msg})
    except Exception as e:
        error_msg = f"信息识别异常：{str(e)}"
        logger.error(error_msg, exc_info=True)
        return jsonify({"success": False, "error": error_msg})


def _call_dify_generate(tender_info_str, correction_hints=""):
    """调用Dify生成工作流，支持带纠错提示的重试"""
    input_text = tender_info_str
    if correction_hints:
        input_text = (
            tender_info_str
            + "\n\n【上一次生成存在以下幻觉问题，本次必须修正】\n"
            + correction_hints
            + "\n严禁再次出现上述问题，所有内容必须100%来自上方结构化信息。"
        )
    headers = {
        "Authorization": f"Bearer {DIFY_GENERATE_API_KEY}",
        "Content-Type": "application/json"
    }
    dify_data = {
        "inputs": {"tender_info": input_text},
        "response_mode": "blocking",
        "user": "flask_client"
    }
    response = requests.post(DIFY_GENERATE_URL, headers=headers, json=dify_data, timeout=600, proxies=DIFY_PROXIES)
    response.raise_for_status()
    result = response.json()
    logger.info(f"Dify生成原始返回结构：{str(result)[:300]}")
    # 兼容两种输出字段名
    outputs = result.get("data", {}).get("outputs", {})
    return outputs.get("tender_document", "") or outputs.get("text", "")


def _validate_hallucination(tender_content, original_data):
    """
    校验生成内容是否存在幻觉：检查生成文本中出现的关键实体是否都能在原始数据中找到。
    返回 (is_clean: bool, issues: list[str])
    """
    # 将原始数据展平为一个大字符串，作为"事实库"
    original_text = json.dumps(original_data, ensure_ascii=False)

    issues = []

    import re

    # 0. 检查项目名称（最关键，防止 Dify 知识库污染）
    project_name = None
    if isinstance(original_data, dict):
        # 兼容嵌套结构
        for key in ['项目名称', 'project_name', 'name']:
            val = original_data.get(key)
            if not val:
                # 尝试从嵌套 dict 里找
                for v in original_data.values():
                    if isinstance(v, dict):
                        val = v.get(key)
                        if val:
                            break
            if val:
                project_name = str(val).strip()
                break
    if project_name:
        # 生成文本里出现的"项目名称：XXX"或"标的名称：XXX"必须和原始数据一致
        name_patterns = re.findall(r'(?:项目名称|标的名称)[：:]\s*(.+?)(?:\n|$|，|。)', tender_content)
        for found_name in name_patterns:
            found_name = found_name.strip()
            if found_name and found_name not in original_text:
                issues.append(f"项目名称'{found_name}'与原始数据不符，疑似幻觉（原始：{project_name}）")

    # 1. 检查金额（如：100万、50,000元、预算金额等数字）
    amounts = re.findall(r'[\d,，]+(?:万|千|百)?(?:元|万元|千元)', tender_content)
    for amt in amounts:
        clean_amt = re.sub(r'[,，]', '', amt)
        if clean_amt not in original_text and amt not in original_text:
            issues.append(f"金额'{amt}'在原始数据中不存在，疑似幻觉")

    # 2. 检查日期（如：2024年3月15日、2024-03-15）
    dates = re.findall(r'\d{4}年\d{1,2}月\d{1,2}日|\d{4}-\d{2}-\d{2}', tender_content)
    for d in dates:
        if d not in original_text:
            issues.append(f"日期'{d}'在原始数据中不存在，疑似幻觉")

    # 3. 检查电话号码
    phones = re.findall(r'1[3-9]\d{9}|0\d{2,3}-\d{7,8}', tender_content)
    for p in phones:
        if p not in original_text:
            issues.append(f"电话'{p}'在原始数据中不存在，疑似幻觉")

    # 4. 检查招标编号（必须同时含字母和数字，且长度>8，避免误报普通词）
    codes = re.findall(r'[A-Z]{2,}[0-9\-]{3,}[A-Z0-9\-]*|[0-9]{3,}[A-Z]{2,}[A-Z0-9\-]*', tender_content)
    for c in codes:
        if len(c) > 8 and c not in original_text:
            issues.append(f"编号'{c}'在原始数据中不存在，疑似幻觉")

    is_clean = len(issues) == 0
    return is_clean, issues


@app.route('/generate_tender', methods=['POST'])
def generate_tender():
    """
    招标文件生成接口
    优先走 LangGraph（FastAPI /generate/tender）：
      planner → router → knowledge_search（检索规范模板）→ bid_generate_node → verifier → END
    LangGraph 不可用时降级到 Dify + Python 幻觉校验兜底
    """
    try:
        data = request.get_json()
        if not data:
            return jsonify({"success": False, "error": "未接收到生成参数"})

        # 兼容前端传来的 {form_data: {...}} 和直接传 {...} 两种格式
        form_data = data.get('form_data', data)
        logger.info(f"接收到生成请求，数据长度：{len(str(form_data))}")

        # ===== 优先：LangGraph 生成 =====
        try:
            agent_response = requests.post(
                f"{AI_AGENT_BASE_URL}/generate/tender",
                json={"tender_info": form_data},
                timeout=300,
                proxies=DIFY_PROXIES
            )
            agent_response.raise_for_status()
            agent_result = agent_response.json()

            if agent_result.get("success") and agent_result.get("tender_document"):
                logger.info("LangGraph 生成成功")
                return jsonify({"success": True, "data": agent_result["tender_document"]})
            else:
                logger.warning(f"LangGraph 返回失败：{agent_result.get('error')}，降级到 Dify")

        except requests.exceptions.ConnectionError:
            logger.warning("FastAPI 服务未启动，降级到 Dify 生成")
        except requests.exceptions.Timeout:
            logger.warning("LangGraph 生成超时，降级到 Dify 生成")
        except Exception as e:
            logger.warning(f"LangGraph 调用异常：{str(e)[:100]}，降级到 Dify 生成")

        # ===== 降级：Dify + Python 幻觉校验 =====
        logger.info("使用 Dify 兜底生成")
        tender_info_str = format_tender_info(form_data)
        best_result = ""
        best_issue_count = float('inf')
        correction_hints = ""
        max_attempts = 3

        for attempt in range(1, max_attempts + 1):
            logger.info(f"Dify 第{attempt}次生成招标文件")
            tender_content = _call_dify_generate(tender_info_str, correction_hints)

            if not tender_content:
                correction_hints = "上次生成内容为空，请重新生成完整招标文件。"
                continue

            is_clean, issues = _validate_hallucination(tender_content, form_data)
            if is_clean:
                logger.info(f"Dify 第{attempt}次生成通过校验")
                return jsonify({"success": True, "data": tender_content})

            logger.warning(f"Dify 第{attempt}次发现{len(issues)}个幻觉问题")
            # 保留幻觉最少的那次结果
            if len(issues) < best_issue_count:
                best_issue_count = len(issues)
                best_result = tender_content
            correction_hints = "\n".join(issues[:10])

        logger.warning(f"Dify 3次均未通过校验，返回幻觉最少的结果（{best_issue_count}个问题）")
        return jsonify({"success": True, "data": best_result or "生成内容为空，请重试"})

    except requests.exceptions.HTTPError as e:
        error_msg = f"生成请求失败（状态码：{e.response.status_code}）：{e.response.text[:500]}"
        logger.error(error_msg)
        return jsonify({"success": False, "error": error_msg})
    except Exception as e:
        error_msg = f"招标文件生成异常：{str(e)}"
        logger.error(error_msg, exc_info=True)
        return jsonify({"success": False, "error": error_msg})


def _call_dify_review(tender_content_with_line, company_name, feedback=""):
    """调用Dify审核工作流，支持传入feedback用于重试纠错"""
    headers = {
        "Authorization": f"Bearer {DIFY_COMPLIANCE_API_KEY}",
        "Content-Type": "application/json"
    }
    dify_data = {
        "inputs": {
            "tender_document": tender_content_with_line,
            "company_name": company_name,
            "feedback": feedback
        },
        "response_mode": "blocking",
        "user": "flask_client"
    }
    response = requests.post(DIFY_COMPLIANCE_URL, headers=headers, json=dify_data, timeout=600, proxies=DIFY_PROXIES)
    response.raise_for_status()
    result = response.json()
    logger.info(f"Dify审核原始返回结构：{str(result)[:500]}")
    logger.info(f"Dify审核完整text（前2000字）：{str(result.get('data', {}).get('outputs', {}).get('text', ''))[:2000]}")
    # Dify workflow blocking 模式返回结构：{"data": {"outputs": {"text": "..."}}}
    text = result.get("data", {}).get("outputs", {}).get("text", "")
    if not text:
        # 兼容部分版本直接返回 outputs 在顶层
        text = result.get("outputs", {}).get("text", "")
    return text.strip()


def _parse_review_text(review_text):
    """解析审核文本，返回 (review_result列表, item_count字典)"""
    review_result = []
    current_block = ""
    item_count = {
        "招标需求文件审查": 0,
        "合同文件审查": 0,
        "投标文件审查": 0,
        "招标文件撰写避错专项审查": 0
    }

    # 去掉 <think>...</think> 思考过程，只保留实际输出
    clean_text = re.sub(r'<think>.*?</think>', '', review_text, flags=re.DOTALL)
    all_lines = [line.strip() for line in clean_text.split('\n') if line.strip()]
    for line in all_lines:
        if line.startswith('一、招标需求文件审查'):
            current_block = "招标需求文件审查"
            continue
        elif line.startswith('二、合同文件审查'):
            current_block = "合同文件审查"
            continue
        elif line.startswith('三、投标文件审查'):
            current_block = "投标文件审查"
            continue
        elif line.startswith('四、招标文件撰写避错专项审查'):
            current_block = "招标文件撰写避错专项审查"
            continue

        if not current_block:
            continue

        item_match = re.match(r"第(\d+)条\. ([^：]+)：", line)
        item_num = item_match.group(1) if item_match else ""
        problem_type = item_match.group(2).strip() if item_match else "未知审查项"

        desc_match = re.search(r"第\d+条\. [^：]+：([^（]+)", line)
        problem_desc = desc_match.group(1).strip() if desc_match else ""
        if "该审查项无问题" in line:
            problem_desc = "该审查项无问题"

        original_match = re.search(r'对应原文[：:][\"\u201c\u300c]([^\"\u201d\u300d]+)[\"\u201d\u300d]', line)
        original_text = original_match.group(1).strip() if original_match else ""

        line_num_match = re.search(r"（行号：(\d+)）", line)
        line_num = line_num_match.group(1) if line_num_match else "未匹配到"

        suggest_match = re.search(r"整改建议：(.+)", line)
        suggestion = suggest_match.group(1).strip() if suggest_match else ""

        review_result.append({
            "block": current_block,
            "itemNum": item_num,
            "problemType": problem_type,
            "problemDesc": problem_desc,
            "originalText": original_text,
            "suggestion": suggestion,
            "lineNum": line_num
        })
        item_count[current_block] += 1

    return review_result, item_count


def _sanitize_review_result(review_result, tender_content_with_line):
    """
    对审核结果做宽松清洗：
    - 保留所有问题描述（信任 LLM 的判断）
    - 只对行号做存在性校验：行号不存在则清空，但保留问题本身
    - 原文片段做模糊匹配：找不到则清空原文字段，但保留问题本身
    - 返回清洗后的结果和虚假行号数量（用于判断质量）
    """
    line_map = {}
    for raw_line in tender_content_with_line.split('\n'):
        m = re.match(r"【行(\d+)】(.*)", raw_line)
        if m:
            line_map[m.group(1)] = m.group(2)

    fake_count = 0
    cleaned = []
    for item in review_result:
        item = dict(item)  # 浅拷贝，不修改原始数据
        if item["problemDesc"] == "该审查项无问题":
            cleaned.append(item)
            continue

        line_num = item.get("lineNum", "")
        original_text = item.get("originalText", "")

        # 行号校验：不存在就清空行号，记录为虚假
        if line_num and line_num != "未匹配到":
            if line_num not in line_map:
                item["lineNum"] = "未匹配到"
                fake_count += 1

        # 原文校验：找不到就清空，但保留问题描述
        if original_text:
            # 先精确匹配，再去除空格后模糊匹配
            normalized_orig = original_text.replace(" ", "").replace("\u3000", "")
            normalized_doc = tender_content_with_line.replace(" ", "").replace("\u3000", "")
            if original_text not in tender_content_with_line and normalized_orig not in normalized_doc:
                item["originalText"] = ""
                fake_count += 1

        cleaned.append(item)

    return cleaned, fake_count


@app.route('/review_tender', methods=['POST'])
def review_tender():
    """
    招标文件智能审核接口
    优先走 LangGraph 图（FastAPI /review/tender）：planner → review_node → verifier → end
    失败时降级调用 Dify 审核工作流（Plan B）
    """
    try:
        data = request.get_json()
        tender_content = data.get('content', '').strip()
        tender_content_with_line = data.get('contentWithLine', tender_content)
        company_name = data.get('company_name', '上海大学')

        if not tender_content:
            return jsonify({"success": False, "error": "未获取到招标文件内容"})

        # ===== Plan A：LangGraph 图 =====
        review_text = ""
        used_langgraph = False
        try:
            logger.info("调用 FastAPI LangGraph 审核图...")
            resp = requests.post(
                f"{AI_AGENT_BASE_URL}/review/tender",
                json={"tender_content_with_line": tender_content_with_line},
                timeout=600,
            )
            resp.raise_for_status()
            review_text = resp.json().get("review_text", "")
            if review_text:
                used_langgraph = True
                logger.info(f"LangGraph 审核成功，返回长度：{len(review_text)}")
            else:
                logger.warning("LangGraph 审核返回空文本，降级 Dify")
        except requests.exceptions.RequestException as e:
            logger.warning(f"LangGraph 审核请求失败，降级 Dify：{e}")

        # ===== Plan B：Dify 降级 =====
        if not review_text:
            try:
                logger.info("降级调用 Dify 审核工作流...")
                review_text = _call_dify_review(tender_content_with_line, company_name)
                if review_text:
                    logger.info(f"Dify 审核成功，返回长度：{len(review_text)}")
                else:
                    return jsonify({"success": False, "error": "LangGraph 和 Dify 均未返回有效审核内容"})
            except Exception as dify_err:
                logger.error(f"Dify 审核也失败了：{dify_err}")
                return jsonify({"success": False, "error": f"审核服务不可用：{str(dify_err)}"})

        logger.info(f"审核来源：{'LangGraph' if used_langgraph else 'Dify'}，前200字：{review_text[:200]}")
        review_result, item_count = _parse_review_text(review_text)
        total_parsed = sum(item_count.values())
        logger.info(f"解析到{total_parsed}条审核结果")

        cleaned_result, fake_count = _sanitize_review_result(review_result, tender_content_with_line)
        problem_count = len([x for x in cleaned_result if x['problemDesc'] != '该审查项无问题'])
        logger.info(f"审核完成，识别问题数：{problem_count}，虚假字段数：{fake_count}")

        return jsonify({"success": True, "result": cleaned_result})

    except Exception as e:
        error_msg = f"招标文件审核异常：{str(e)}"
        logger.error(error_msg, exc_info=True)
        return jsonify({"success": False, "error": error_msg})


@app.route('/upload', methods=['POST'])
def upload_file():
    """
    文件上传接口（纯文本提取，移除冗余的Dify审核调用）
    - 仅提取文件文本并添加行号，不提前审核
    - 返回纯文本+带行号文本，供前端使用
    """
    try:
        # 1. 校验文件
        if 'file' not in request.files:
            return jsonify({"success": False, "error": "未上传文件"})
        file = request.files['file']
        if file.filename == '':
            return jsonify({"success": False, "error": "文件名不能为空"})

        file_ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
        if file_ext not in ALLOWED_EXTENSIONS:
            return jsonify({"success": False, "error": f"仅支持{ALLOWED_EXTENSIONS}格式文件"})

        # 2. 提取文件文本并添加行号（保留所有行，和前端行号规则统一）
        file_content_result = extract_file_content(file)
        if isinstance(file_content_result, tuple):
            file_content, tender_content_with_line = file_content_result
        else:
            # 处理错误情况
            return jsonify({"success": False, "error": file_content_result})

        if file_content.startswith(('不支持', '文件解析失败')):
            return jsonify({"success": False, "error": file_content})

        # 3. 仅返回文本内容，不调用Dify审核（核心修改）
        logger.info(f"文件上传成功：{file.filename}，文本长度：{len(file_content)}")
        return jsonify({
            "success": True,
            "filename": file.filename,
            "content": file_content,  # 纯文本（前端显示用）
            "contentWithLine": tender_content_with_line  # 带行号文本（后续审核用）
        })

    except Exception as e:
        error_msg = f"文件上传异常：{str(e)}"
        logger.error(error_msg, exc_info=True)
        return jsonify({"success": False, "error": error_msg})


# ===================== 新增：页面路由（核心修改部分） =====================
@app.route('/')
def index():
    """平台主页（替换原有首页，指向新的index.html）"""
    return render_template('index.html')


@app.route('/create_record/<module_name>')
def create_record(module_name):
    """创建与记录页面（通用）"""
    # 过滤出当前模块的项目记录
    module_projects = [p for p in projects if p["type"] == module_name]
    return render_template('create_record.html', module_name=module_name, projects=module_projects)


@app.route('/select_mode/<module_name>')
def select_mode(module_name):
    """功能选择页面（仅招标文件模块）"""
    return render_template('select_mode.html', module_name=module_name)


@app.route('/create_project', methods=['POST'])
def create_project():
    """创建项目接口"""
    project_type = request.form.get('project_type')
    project_name = request.form.get('project_name')
    mode = request.form.get('mode')
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not project_name:
        flash("项目名称不能为空！", "error")
        return redirect(url_for('select_mode', module_name=project_type))

    # 保存项目记录
    project_id = save_project(project_type, project_name, mode, created_at)

    # 根据模式跳转到对应工作页面
    if project_type == "tender_generate" and mode == "generate":
        return redirect(url_for('tender_generate', project_id=project_id))
    elif project_type == "tender_generate" and mode == "compliance":
        return redirect(url_for('compliance', project_id=project_id))
    elif project_type == "tender_compliance":
        return redirect(url_for('compliance', project_id=project_id))
    else:
        return redirect(url_for('create_record', module_name=project_type))


@app.route('/save_project', methods=['POST'])
def save_project_endpoint():
    """保存项目内容接口"""
    data = request.get_json()
    project_id = data.get('project_id')
    content = data.get('content')

    # 更新项目数据
    if update_project_data(project_id, data):
        flash("已全部保存", "success")
        return jsonify({"success": True})
    else:
        return jsonify({"success": False, "error": "项目不存在"})


# 新增：获取项目数据接口
@app.route('/get_project_data', methods=['GET'])
def get_project_data_endpoint():
    project_id = request.args.get('project_id')
    if not project_id:
        return jsonify({"success": False, "error": "缺少project_id"})

    data = get_project_data(project_id)
    return jsonify({"success": True, "data": data})


# ===================== 保留你原有页面路由（修改：增加重定向逻辑） =====================
@app.route('/compliance')
def compliance():
    """合规审查页面"""
    project_id = request.args.get('project_id')
    if not project_id:
        # 如果没有项目ID，重定向到项目管理页
        return redirect(url_for('create_record', module_name='tender_generate'))
    return render_template('compliance.html', project_id=project_id)


@app.route('/tender_generate')
def tender_generate():
    """招标文件生成页面"""
    project_id = request.args.get('project_id')
    if not project_id:
        # 如果没有项目ID，重定向到项目管理页
        return redirect(url_for('create_record', module_name='tender_generate'))
    return render_template('generate.html', project_id=project_id)

# ===================== AI辅助论证模块 =====================
# 新增：AI辅助论证工作流配置

@app.route('/ai_argument')
def ai_argument():
    """AI辅助论证模块首页（创建与记录）"""
    module_name = "ai_argument"
    # 只筛选出类型为 "ai_argument" 的项目
    module_projects = [p for p in projects if p["type"] == module_name]
    return render_template('ai_argument_record.html', module_name=module_name, projects=module_projects)

@app.route('/create_ai_argument_project', methods=['POST'])
def create_ai_argument_project():
    """创建AI辅助论证项目"""
    project_type = request.form.get('project_type')
    project_name = request.form.get('project_name')
    mode = "argument"  # 固定为评审模式
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not project_name:
        flash("项目名称不能为空！", "error")
        return redirect(url_for('ai_argument'))

    project_id = save_project(project_type, project_name, mode, created_at)
    return redirect(url_for('ai_argument_detail', project_id=project_id))

@app.route('/ai_argument_detail')
def ai_argument_detail():
    """AI辅助论证详情页（合同评分）"""
    project_id = request.args.get('project_id')
    if not project_id:
        return redirect(url_for('ai_argument'))
    return render_template('ai_argument_detail.html', project_id=project_id)


@app.route('/extract_scoring_criteria', methods=['POST'])
def extract_scoring_criteria():
    """从招标文件中提取评分维度，返回8个维度的名称和满分"""
    try:
        data = request.get_json()
        file_content = data.get('content', '').strip()
        if not file_content:
            return jsonify({"success": False, "error": "文件内容为空"})

        system_prompt = """你是招投标专家。从招标文件的评标办法/评分细则中提取评分维度。

要求：
1. 提取所有评分维度，合并为恰好8个（不足8个时保留原有，多于8个时合并相近项）
2. 每个维度给出：key（英文字段名，snake_case）、label（中文名称）、max_score（满分，整数）
3. 所有维度满分之和必须等于100
4. 第一个字符必须是 {，最后一个字符必须是 }，不输出任何解释

返回格式：
{"dimensions": [{"key": "price_score", "label": "报价得分", "max_score": 30}, ...]}"""

        payload = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"请从以下招标文件中提取评分维度：\n\n{file_content[:8000]}"}
            ],
            "temperature": 0.0,
            "stream": False
        }
        headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
        resp = requests.post(DEEPSEEK_BASE_URL, headers=headers, json=payload, timeout=60,
                             proxies=DIFY_PROXIES)
        resp.raise_for_status()
        raw = resp.json()["choices"][0]["message"]["content"]
        raw = re.sub(r'<think>.*?</think>', '', raw, flags=re.DOTALL).strip()
        raw = re.sub(r'^```(?:json)?\s*', '', raw).strip()
        raw = re.sub(r'\s*```$', '', raw).strip()

        result = json.loads(raw)
        dimensions = result.get("dimensions", [])

        # 校验总分
        total = sum(d.get("max_score", 0) for d in dimensions)
        logger.info(f"提取到{len(dimensions)}个评分维度，总分{total}")
        return jsonify({"success": True, "dimensions": dimensions, "total": total})

    except Exception as e:
        logger.error(f"评分维度提取失败：{e}")
        return jsonify({"success": False, "error": str(e)})


@app.route('/analyze_contracts', methods=['POST'])
def analyze_contracts():
    """投标文件智能评审，直接调用 DeepSeek，绕开 Dify"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"success": False, "error": "未接收到任何数据"})

        contracts = data.get('contracts', [])
        project_id = data.get('project_id')
        custom_criteria = data.get('custom_criteria', '').strip()
        custom_dimensions = data.get('dimensions', [])  # 前端传来的自定义维度

        if not isinstance(contracts, list) or len(contracts) == 0:
            return jsonify({"success": False, "error": "未上传任何投标文件"})
        if not project_id:
            return jsonify({"success": False, "error": "缺少项目ID参数"})

        # 拼接所有文件内容
        combined_content = ""
        for idx, c in enumerate(contracts):
            combined_content += f"\n\n===== 投标文件{idx + 1}：{c.get('name', f'文件{idx+1}')} =====\n"
            combined_content += c.get('content', '')

        file_names = "、".join([c.get('name', f'文件{i+1}') for i, c in enumerate(contracts)])

        # 构建评分维度（优先用前端传来的自定义维度，否则用默认）
        if custom_dimensions and len(custom_dimensions) > 0:
            score_max = {d['key']: float(d['max_score']) for d in custom_dimensions}
            dim_desc = "\n".join([f"- {d['key']}：{d['label']}（满分{d['max_score']}分）" for d in custom_dimensions])
        else:
            score_max = {
                "price_score": 30.0, "tech_score": 45.0, "onsite_service": 5.0,
                "onsite_personnel": 6.0, "after_sales": 3.0, "overall_strength": 4.0,
                "past_performance": 5.0, "doc_quality": 2.0,
            }
            dim_desc = """- price_score：报价得分（满分30分，报价合理性、竞争力）
- tech_score：技术方案（满分45分，含主要技术参数响应15分+一般技术参数响应20分+施工方案10分）
- onsite_service：驻场服务（满分5分，驻场响应、服务承诺）
- onsite_personnel：驻场人员资历（满分6分，人员资质、经验）
- after_sales：售后服务内容（满分3分，售后保障、维保方案）
- overall_strength：综合实力（满分4分，企业规模、资质、案例）
- past_performance：近三年业绩（满分5分，同类项目业绩数量和质量）
- doc_quality：投标文件编制（满分2分，文件完整性、规范性、清晰度）"""

        # 专项评分标准段落（可选）
        custom_criteria_section = ""
        if custom_criteria:
            custom_criteria_section = f"\n\n【本项目背景与特殊要求】\n{custom_criteria}\n请在评分时重点关注以上背景信息。"

        # 构建 JSON 格式示例
        scores_example = "\n".join([f'        "{k}": 0.0,' for k in score_max])
        scores_example = scores_example.rstrip(',')

        system_prompt = f"""你是一名专业的招投标评审专家，负责对多份投标文件进行综合评分，选出最优投标方。

【评分维度及满分】（总分100分）
{dim_desc}{custom_criteria_section}

【评分要求】
1. 每个维度按各自满分给出具体分数（保留1位小数），不得超过该维度满分
2. 评分要有区分度，不同投标方分数不能全部相同
3. 综合评价要指出该投标方的核心优势和主要不足
4. 第一个字符必须是 {{，最后一个字符必须是 }}，不得输出任何 Markdown 或解释文字

请严格按以下 JSON 格式返回：
{{
  "contracts": [
    {{
      "id": "contract_1",
      "name": "投标文件的实际文件名",
      "scores": {{
{scores_example}
      }},
      "evaluation": "综合评价文字，指出优势和不足"
    }}
  ]
}}"""

        payload = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"请对以下{len(contracts)}份投标文件（{file_names}）进行评分：\n{combined_content}"}
            ],
            "temperature": 0.1,
            "stream": False
        }
        headers = {
            "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
            "Content-Type": "application/json"
        }

        logger.info(f"投标评审，直接调用 DeepSeek，项目ID：{project_id}，文件数：{len(contracts)}")
        resp = requests.post(DEEPSEEK_BASE_URL, headers=headers, json=payload, timeout=180)
        resp.raise_for_status()
        raw = resp.json()["choices"][0]["message"]["content"]

        # 去掉 <think> 标签和 markdown 代码块
        raw = re.sub(r'<think>.*?</think>', '', raw, flags=re.DOTALL).strip()
        raw = re.sub(r'^```(?:json)?\s*', '', raw).strip()
        raw = re.sub(r'\s*```$', '', raw).strip()

        logger.info(f"DeepSeek 原始返回前500字：{raw[:500]}")

        try:
            analysis_data = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            analysis_data = {"contracts": []}

        if isinstance(analysis_data, list):
            analysis_data = {"contracts": analysis_data}
        analysis_data.setdefault("contracts", [])

        # 后端强制重算总分，防止 LLM 算错（score_max 已在上方根据自定义维度动态生成）
        # 构建 label→key 反查表，用于 AI 返回中文 key 时的兜底匹配
        label_to_key = {d['label']: d['key'] for d in custom_dimensions} if custom_dimensions else {}

        for idx, contract in enumerate(analysis_data["contracts"]):
            contract.setdefault("id", f"contract_{idx + 1}")
            contract.setdefault("name", contracts[idx].get("name", f"文件{idx+1}") if idx < len(contracts) else f"文件{idx+1}")
            contract.setdefault("evaluation", "暂无评价")
            contract.setdefault("scores", {})

            raw_scores = contract["scores"]

            # 如果 AI 用中文 label 作为 key 返回，先做一次 label→key 的转换
            normalized_scores = {}
            for k, v in raw_scores.items():
                mapped_key = label_to_key.get(k, k)
                normalized_scores[mapped_key] = v

            dimension_scores = []
            for field, max_val in score_max.items():
                try:
                    score = float(normalized_scores.get(field, 0))
                    score = max(0.0, min(max_val, score))
                except (ValueError, TypeError):
                    score = 0.0
                normalized_scores[field] = round(score, 1)
                dimension_scores.append(score)

            normalized_scores["total_score"] = round(sum(dimension_scores), 1)
            contract["scores"] = normalized_scores
            logger.info(f"合同{contract['name']}最终总分：{contract['scores']['total_score']}")

        update_project_data(project_id, {"contracts": contracts, "analysis": analysis_data})
        logger.info(f"投标评审完成，项目ID：{project_id}，处理{len(analysis_data['contracts'])}份文件")

        return jsonify({"success": True, "data": analysis_data})

    except requests.exceptions.Timeout:
        return jsonify({"success": False, "error": "DeepSeek 请求超时（超过3分钟），请检查文件大小或网络"})
    except requests.exceptions.ConnectionError:
        return jsonify({"success": False, "error": "无法连接到 DeepSeek，请检查网络"})
    except Exception as e:
        error_msg = f"投标评审异常：{str(e)}"
        logger.error(error_msg, exc_info=True)
        return jsonify({"success": False, "error": error_msg})


# ===================== 合同评审模块 =====================
@app.route('/contract_review_record')
def contract_review_record():
    """合同评审 - 创建与记录页面"""
    module_projects = [p for p in projects if p["type"] == "contract_review"]
    return render_template('contract_review_record.html', projects=module_projects)


@app.route('/create_contract_review_project', methods=['POST'])
def create_contract_review_project():
    """创建合同评审项目"""
    project_name = request.form.get('project_name')
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not project_name:
        flash("项目名称不能为空！", "error")
        return redirect(url_for('contract_review_record'))
    project_id = save_project("contract_review", project_name, "review", created_at)
    return redirect(url_for('contract_review', project_id=project_id))


@app.route('/contract_review')
def contract_review():
    """合同评审详情页"""
    project_id = request.args.get('project_id')
    if not project_id:
        return redirect(url_for('contract_review_record'))
    return render_template('contract_review.html', project_id=project_id)


def _get_baidu_ocr_token():
    """获取百度 OCR access_token，有效期30天"""
    resp = requests.post(
        "https://aip.baidubce.com/oauth/2.0/token",
        params={
            "grant_type": "client_credentials",
            "client_id": BAIDU_OCR_API_KEY,
            "client_secret": BAIDU_OCR_SECRET_KEY
        },
        proxies=DIFY_PROXIES,
        timeout=10
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def _baidu_ocr_image(img_bytes, token):
    """单张图片调百度通用文字识别，返回识别文字"""
    b64 = base64.b64encode(img_bytes).decode()
    resp = requests.post(
        f"https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token={token}",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={"image": b64},
        proxies=DIFY_PROXIES,
        timeout=30
    )
    resp.raise_for_status()
    result = resp.json()
    if "words_result" not in result:
        logger.warning(f"百度 OCR 返回异常：{result}")
        return ""
    return "\n".join([w["words"] for w in result["words_result"]])


def _extract_pdf_text(file_bytes):
    """提取 PDF 文字：先用 pymupdf，扫描件则用百度 OCR"""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    text = text.strip()

    if text:
        return text

    # 扫描件：用百度 OCR 逐页识别
    logger.info("检测到扫描件，启用百度 OCR")
    token = _get_baidu_ocr_token()
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    ocr_text = ""
    for i, page in enumerate(doc):
        mat = fitz.Matrix(2.0, 2.0)  # 2x 放大提升识别率
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        img_bytes = pix.tobytes("jpeg")
        page_text = _baidu_ocr_image(img_bytes, token)
        ocr_text += page_text + "\n"
        logger.info(f"OCR 第 {i+1} 页，识别 {len(page_text)} 字")
    doc.close()
    return ocr_text.strip()


def _call_deepseek_text_review(contract_text, contract_name):
    """把合同文字发给 DeepSeek，返回审核结果列表"""
    prompt = f"""你是一名专业的政府采购合同审核专家。请对合同文件《{contract_name}》进行合规性审核。

合同内容如下：
{contract_text[:8000]}

请从以下维度逐项检查，以 JSON 数组格式返回结果，不要输出任何其他内容：

审核维度：
1. 合同主体信息（甲乙双方名称、地址、联系方式是否完整）
2. 合同标的与金额（标的描述是否清晰、金额是否明确）
3. 履约期限（交货/服务期限是否明确、是否合理）
4. 付款条款（付款方式、节点、比例是否明确）
5. 验收条款（验收标准、验收方式是否明确）
6. 违约责任（违约情形、赔偿标准是否明确）
7. 争议解决（是否约定争议解决方式）
8. 合同附件（是否引用附件、附件是否完整）

返回格式（严格 JSON，不加 markdown 代码块）：
[
  {{
    "type": "审核维度名称",
    "title": "具体审核项",
    "has_problem": true或false,
    "problem": "问题描述，无问题时填'无问题'",
    "suggestion": "整改建议，无问题时填空字符串"
  }}
]"""

    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-v4-flash",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "max_tokens": 3000
    }
    resp = requests.post(
        DEEPSEEK_BASE_URL,
        headers=headers,
        json=payload,
        proxies=DIFY_PROXIES,
        timeout=120
    )
    resp.raise_for_status()
    raw = resp.json()["choices"][0]["message"]["content"].strip()
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)
    return json.loads(raw)


@app.route('/review_contract', methods=['POST'])
def review_contract():
    """合同评审接口：pymupdf 提取文字 → deepseek-v4-flash 审核"""
    try:
        if 'file' not in request.files:
            return jsonify({"success": False, "error": "未上传文件"})
        file = request.files['file']
        if not file or file.filename == '':
            return jsonify({"success": False, "error": "文件名不能为空"})

        file_ext = file.filename.rsplit('.', 1)[-1].lower()
        file_bytes = file.read()
        contract_name = file.filename

        if file_ext == 'pdf':
            contract_text = _extract_pdf_text(file_bytes)
            if not contract_text:
                return jsonify({"success": False, "error": "文字提取失败，请检查文件是否损坏"})
        elif file_ext == 'docx':
            import docx as docx_lib
            from io import BytesIO
            doc = docx_lib.Document(BytesIO(file_bytes))
            contract_text = '\n'.join([p.text for p in doc.paragraphs]).strip()
            if not contract_text:
                return jsonify({"success": False, "error": "文档内容为空"})
        elif file_ext == 'txt':
            contract_text = file_bytes.decode('utf-8', errors='ignore').strip()
        else:
            return jsonify({"success": False, "error": "仅支持 PDF、DOCX、TXT 格式"})

        logger.info(f"合同评审：{contract_name}，文字长度：{len(contract_text)}，发送给 DeepSeek")
        result = _call_deepseek_text_review(contract_text, contract_name)
        logger.info(f"合同评审完成：{contract_name}，返回 {len(result)} 条审核项")

        project_id = request.form.get('project_id') or request.args.get('project_id')
        return jsonify({"success": True, "result": result, "contract_name": contract_name})

    except json.JSONDecodeError as e:
        logger.error(f"合同评审 JSON 解析失败：{e}")
        return jsonify({"success": False, "error": "AI 返回格式异常，请重试"})
    except Exception as e:
        logger.error(f"合同评审异常：{e}", exc_info=True)
        return jsonify({"success": False, "error": str(e)})


# ===================== 投标论证模块 =====================
@app.route('/tender_bid_record')
def tender_bid_record():
    """投标论证 - 创建与记录页面"""
    module_projects = [p for p in projects if p["type"] == "tender_bid"]
    return render_template('tender_bid_record.html', projects=module_projects)


@app.route('/create_tender_bid_project', methods=['POST'])
def create_tender_bid_project():
    """创建投标论证项目"""
    project_name = request.form.get('project_name')
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not project_name:
        flash("项目名称不能为空！", "error")
        return redirect(url_for('tender_bid_record'))
    project_id = save_project("tender_bid", project_name, "bid", created_at)
    return redirect(url_for('tender_bid', project_id=project_id))


@app.route('/tender_bid')
def tender_bid():
    """投标论证详情页"""
    project_id = request.args.get('project_id')
    if not project_id:
        return redirect(url_for('tender_bid_record'))
    return render_template('tender_bid.html', project_id=project_id)


@app.route('/analyze_bid_collusion', methods=['POST'])
def analyze_bid_collusion():
    """围标串标分析，直接调用 DeepSeek"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"success": False, "error": "未接收到数据"})

        contracts = data.get('contracts', [])
        project_id = data.get('project_id')

        if not isinstance(contracts, list) or len(contracts) < 2:
            return jsonify({"success": False, "error": "请至少上传两份投标文件"})

        combined_content = ""
        for idx, c in enumerate(contracts):
            combined_content += f"\n\n===== 投标文件{idx + 1}：{c.get('name', f'文件{idx+1}')} =====\n"
            combined_content += c.get('content', '')

        file_names = "、".join([c.get('name', f'文件{i+1}') for i, c in enumerate(contracts)])

        system_prompt = """你是一名专业的招投标合规审查专家，擅长识别围标串标行为。
请对提供的多份投标文件进行围标串标分析，重点检查：
1. 文件内容雷同（文字、格式、技术方案高度相似）
2. 报价规律异常（梯度报价、差距过小）
3. 企业关联性（联系方式、地址、人员重叠）
4. 投标行为异常（陪标迹象、让标行为）

【原文引用要求】similar_segments 中每个 pairs 条目必须直接引用文件原文，不得为空或写"无原文证据"，每段不少于40字。

请严格按照以下 JSON 格式返回，不要输出任何其他内容：
{
  "risk_level": "high|medium|low",
  "conclusion": "总体结论（2-3句话）",
  "violations": [
    {
      "type": "违规行为",
      "risk": "high|medium|low",
      "description": "详细描述",
      "suggestion": "处理建议"
    }
  ],
  "similar_segments": [
    {
      "label": "雷同内容标题",
      "risk": "high|medium|low",
      "pairs": [
        {"file": "投标文件1的实际文件名", "text": "从该文件直接截取的原文，不少于40字"},
        {"file": "投标文件2的实际文件名", "text": "从该文件直接截取的原文，不少于40字"}
      ]
    }
  ],
  "bidders": [
    {
      "name": "投标人名称",
      "summary": "该投标人分析摘要"
    }
  ]
}"""

        payload = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"请对以下{len(contracts)}份投标文件（{file_names}）进行围标串标分析：\n{combined_content}"}
            ],
            "temperature": 0.1,
            "stream": False
        }
        headers = {
            "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
            "Content-Type": "application/json"
        }

        logger.info(f"围标串标分析，直接调用 DeepSeek，项目ID：{project_id}，文件数：{len(contracts)}")
        resp = requests.post(DEEPSEEK_BASE_URL, headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        raw = resp.json()["choices"][0]["message"]["content"]

        # 去掉 <think> 标签和 markdown 代码块
        raw = re.sub(r'<think>.*?</think>', '', raw, flags=re.DOTALL).strip()
        raw = re.sub(r'^```(?:json)?\s*', '', raw).strip()
        raw = re.sub(r'\s*```$', '', raw).strip()

        try:
            analysis_data = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            analysis_data = {
                "risk_level": _extract_risk_level(raw),
                "bidders": [],
                "violations": [],
                "similar_segments": [],
                "conclusion": raw,
            }

        analysis_data.setdefault("risk_level", "low")
        analysis_data.setdefault("bidders", [])
        analysis_data.setdefault("violations", [])
        analysis_data.setdefault("similar_segments", [])
        analysis_data.setdefault("conclusion", "")

        # 兜底：如果 similar_segments 的 pairs 原文为空，从 contracts 内容里自动提取
        contract_map = {c.get("name", f"文件{i+1}"): c.get("content", "") for i, c in enumerate(contracts)}
        for seg in analysis_data["similar_segments"]:
            for pair in seg.get("pairs", []):
                if not pair.get("text") or pair["text"] in ("无原文证据", "（无原文证据）", ""):
                    fname = pair.get("file", "")
                    content = contract_map.get(fname, "")
                    if content and seg.get("label"):
                        # 用 label 关键词在原文里找最相关的段落
                        keyword = seg["label"][:8]
                        idx = content.find(keyword)
                        if idx == -1:
                            # 找不到关键词就取前200字
                            pair["text"] = content[:200].strip()
                        else:
                            pair["text"] = content[max(0, idx-20):idx+180].strip()

        if project_id:
            update_project_data(project_id, {
                "contracts": [{"name": c.get("name"), "content": c.get("content")} for c in contracts],
                "analysis_result": analysis_data
            })

        return jsonify({"success": True, "result": analysis_data})

    except requests.exceptions.Timeout:
        return jsonify({"success": False, "error": "分析超时，请稍后重试"})
    except requests.exceptions.RequestException as e:
        logger.error(f"DeepSeek 请求失败: {e}")
        return jsonify({"success": False, "error": f"AI 服务请求失败：{str(e)}"})
    except Exception as e:
        logger.error(f"围标串标分析异常: {e}")
        return jsonify({"success": False, "error": str(e)})


def _extract_risk_level(text: str) -> str:
    """从 Markdown 报告里提取风险等级"""
    high_keywords = ["高风险", "极高", "高度异常", "围标串标嫌疑", "串通投标", "严重违规", "高度雷同"]
    medium_keywords = ["中风险", "中等风险", "存在风险", "需关注", "疑似"]
    for kw in high_keywords:
        if kw in text:
            return "high"
    for kw in medium_keywords:
        if kw in text:
            return "medium"
    return "low"


# ===================== 全链条功能路由 =====================

# 全链条5个步骤定义（顺序固定）
FULL_CHAIN_STEPS = [
    {"key": "tender_generate", "label": "招标文件生成",   "url_tpl": "/tender_generate?project_id={id}"},
    {"key": "tender_compliance","label": "招标文件审核",   "url_tpl": "/compliance?project_id={id}"},
    {"key": "tender_bid",       "label": "围标串标审查",   "url_tpl": "/tender_bid?project_id={id}"},
    {"key": "ai_argument",      "label": "投标文件智能评审","url_tpl": "/ai_argument_detail?project_id={id}"},
    {"key": "contract_review",  "label": "合同评审",       "url_tpl": "/contract_review?project_id={id}"},
]

# 全链条项目单独存储
full_chain_projects = []


def _fc_progress(fc_project):
    """计算全链条项目已完成步骤数（linked_steps 中有 project_id 的算完成）"""
    linked = fc_project.get("linked_steps", {})
    return sum(1 for step in FULL_CHAIN_STEPS if linked.get(step["key"]))


@app.route('/full_chain_record')
def full_chain_record():
    """全链条项目列表页"""
    projects_with_progress = []
    for p in full_chain_projects:
        done = _fc_progress(p)
        projects_with_progress.append({**p, "done_steps": done, "total_steps": len(FULL_CHAIN_STEPS)})
    return render_template('full_chain_record.html', projects=projects_with_progress)


@app.route('/create_full_chain_project', methods=['POST'])
def create_full_chain_project():
    """创建全链条项目"""
    project_name = request.form.get('project_name', '').strip()
    if not project_name:
        return redirect(url_for('full_chain_record'))
    project_id = len(full_chain_projects) + 1
    full_chain_projects.append({
        "id": project_id,
        "name": project_name,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "linked_steps": {}   # key -> project_id（来自各模块的 projects 列表）
    })
    return redirect(url_for('full_chain_detail', project_id=project_id))


@app.route('/full_chain_detail')
def full_chain_detail():
    """全链条项目详情页"""
    project_id = request.args.get('project_id', type=int)
    fc_project = next((p for p in full_chain_projects if p['id'] == project_id), None)
    if not fc_project:
        return redirect(url_for('full_chain_record'))

    # 构建步骤状态
    steps_info = []
    for step in FULL_CHAIN_STEPS:
        linked_pid = fc_project["linked_steps"].get(step["key"])
        linked_project = None
        if linked_pid:
            linked_project = next((p for p in projects if str(p['id']) == str(linked_pid)), None)
        steps_info.append({
            "key": step["key"],
            "label": step["label"],
            "url_tpl": step["url_tpl"],
            "linked_pid": linked_pid,
            "linked_project": linked_project,
            # 精简版，供前端 tojson 初始化预览用（不含 data 字段）
            "linked_pid_str": str(linked_pid) if linked_pid else "",
        })

    # 各模块可供迁移的历史记录
    module_records = {}
    key_to_type = {
        "tender_generate":   ("tender_generate", "generate"),
        "tender_compliance": ("tender_compliance", "compliance"),
        "tender_bid":        ("tender_bid",       "bid"),
        "ai_argument":       ("ai_argument",      "argument"),
        "contract_review":   ("contract_review",  "review"),
    }
    for step in FULL_CHAIN_STEPS:
        ptype, pmode = key_to_type[step["key"]]
        records = [p for p in projects if p["type"] == ptype and p["mode"] == pmode]
        module_records[step["key"]] = records

    done = _fc_progress(fc_project)
    return render_template(
        'full_chain_detail.html',
        fc_project=fc_project,
        steps_info=steps_info,
        module_records=module_records,
        done_steps=done,
        total_steps=len(FULL_CHAIN_STEPS),
    )


@app.route('/full_chain_link', methods=['POST'])
def full_chain_link():
    """将某模块的历史记录关联到全链条项目的某个步骤"""
    fc_id = request.form.get('fc_id', type=int)
    step_key = request.form.get('step_key', '').strip()
    module_pid = request.form.get('module_pid', '').strip()

    fc_project = next((p for p in full_chain_projects if p['id'] == fc_id), None)
    if not fc_project or not step_key:
        return jsonify({"success": False, "error": "参数错误"})

    if module_pid:
        fc_project["linked_steps"][step_key] = module_pid
    else:
        fc_project["linked_steps"].pop(step_key, None)

    done = _fc_progress(fc_project)
    return jsonify({"success": True, "done_steps": done, "total_steps": len(FULL_CHAIN_STEPS)})


# ===================== 全链条预览接口 =====================
@app.route('/get_linked_preview', methods=['GET'])
def get_linked_preview():
    """获取某模块项目的 AI 结果，用于全链条详情页预览"""
    step_key = request.args.get('step_key', '')
    project_id = request.args.get('project_id', '')
    if not step_key or not project_id:
        return jsonify({"success": False, "error": "缺少参数"})

    data = get_project_data(project_id)
    if not data:
        return jsonify({"success": False, "error": "暂无保存数据"})

    preview = {}
    if step_key == 'tender_generate':
        preview['type'] = 'html'
        preview['content'] = data.get('generated_content') or data.get('content') or ''
    elif step_key == 'tender_compliance':
        preview['type'] = 'html'
        preview['content'] = data.get('review_result') or data.get('content') or ''
    elif step_key == 'tender_bid':
        raw = data.get('analysis_result') or {}
        if isinstance(raw, str):
            try:
                raw = json.loads(raw)
            except Exception:
                raw = {}
        preview['type'] = 'bid_collusion'
        preview['content'] = raw
    elif step_key == 'ai_argument':
        preview['type'] = 'ai_argument'
        preview['content'] = data.get('analysis') or {}
    elif step_key == 'contract_review':
        # 合同评审前端存的是 HTML 字符串
        preview['type'] = 'html'
        preview['content'] = data.get('review_result') or ''

    if not preview.get('content'):
        return jsonify({"success": False, "error": "该记录暂无 AI 结果，请先在对应模块完成分析并保存"})

    return jsonify({"success": True, "preview": preview})


# ===================== 启动服务（保留你原有代码） =====================
if __name__ == '__main__':
    # debug=True 仅开发环境使用，生产环境需关闭
    # host=0.0.0.0 允许外网访问，port=5001 端口可自定义
    # threaded=True 开启多线程，提升并发处理能力
    app.run(debug=True, host='0.0.0.0', port=5001, threaded=True)