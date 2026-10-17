import argparse
import json
import os
import sys
import urllib.request
from datetime import date
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(__file__).parent
DEFAULT_INPUT = BASE / "demo_feedback.txt"
OUT_DIR = BASE / "output"

CATEGORIES = ["商品质量", "服务态度", "价格", "卫生", "商品缺货", "其他"]

KEYWORDS = {
    "商品质量": ["变质", "过期", "发霉", "异味", "不新鲜", "口感", "质量", "破损", "漏", "坏"],
    "服务态度": ["态度", "服务", "热情", "冷漠", "不耐烦", "礼貌", "贴心", "导购", "收银"],
    "价格": ["贵", "便宜", "价格", "涨价", "性价比", "实惠", "打折"],
    "卫生": ["卫生间", "厕所", "脏", "卫生", "苍蝇", "干净", "整洁", "油腻"],
    "商品缺货": ["没货", "缺货", "买不到", "断货", "卖完", "下架", "抢不到"],
}

NEGATIVE_WORDS = ["差", "坏", "投诉", "生气", "失望", "贵", "脏", "缺货", "变质", "过期", "发霉", "不新鲜", "冷漠", "不耐烦", "糟糕", "烂", "异味", "破损", "不敢", "没买到"]
POSITIVE_WORDS = ["好", "赞", "满意", "热情", "便宜", "新鲜", "干净", "贴心", "感谢", "棒", "喜欢", "实惠", "舒服", "好吃", "回购"]


def suggest_action(category, sentiment):
    if sentiment == "负面":
        if category == "商品质量":
            return "立即联系门店核实，24 小时内回访，必要时启动商品下架或抽检"
        if category == "卫生":
            return "立即联系门店核实，当天安排卫生整改/消杀，24 小时内回访"
        if category == "服务态度":
            return "转门店主管，当天回访道歉并反馈改进"
        if category == "价格":
            return "转采购/定价部门核查，3 个工作日内回复说明"
        if category == "商品缺货":
            return "转补货/采购，反馈到货时间并登记需求"
        return "登记并转相关部门，48 小时内回访"
    if sentiment == "正面":
        return "记录为正面案例，用于内部激励与宣传"
    return "登记归档，作为改进建议参考"


def rule_classify(text):
    category, matched = "其他", []
    for cat, kws in KEYWORDS.items():
        hit = [k for k in kws if k in text]
        if len(hit) > len(matched):
            category, matched = cat, hit

    if any(w in text for w in NEGATIVE_WORDS):
        sentiment = "负面"
    elif any(w in text for w in POSITIVE_WORDS):
        sentiment = "正面"
    else:
        sentiment = "中性"

    priority = "低"
    if sentiment == "负面":
        priority = "高" if category in ("商品质量", "卫生") else "中"

    reason = f"命中关键词[{('、'.join(matched) if matched else '无')}]，情绪={sentiment}"
    return {
        "category": category,
        "sentiment": sentiment,
        "priority": priority,
        "suggested_action": suggest_action(category, sentiment),
        "reason": reason,
    }


def _parse_json(text):
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text)


def llm_classify(text, base_url, api_key, model):
    prompt = (
        "你是胖东来门店的顾客反馈分析助手。分析下面这条顾客反馈，"
        "严格只输出一个 JSON 对象，不要输出任何其他文字、注释或代码块标记。\n"
        "字段与取值约束：\n"
        "- category: 商品质量 / 服务态度 / 价格 / 卫生 / 商品缺货 / 其他（六选一）\n"
        "- sentiment: 正面 / 中性 / 负面（三选一）\n"
        "- priority: 高 / 中 / 低（三选一，负面且涉及商品质量或卫生时为高）\n"
        "- suggested_action: 一句话处理建议\n"
        "- reason: 一句话归类理由\n\n"
        f"顾客反馈：{text}"
    )
    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": "你是一名严谨的零售客服质检员，输出必须为合法 JSON。"},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    req = urllib.request.Request(
        base_url.rstrip("/") + "/chat/completions",
        data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return _parse_json(data["choices"][0]["message"]["content"])


def render_markdown(results, current_date):
    counts = {}
    for r in results:
        key = (r["category"], r["sentiment"])
        counts[key] = counts.get(key, 0) + 1
    lines = [
        f"# 顾客反馈分类与优先级清单（{current_date.isoformat()}）",
        "",
        f"- 反馈总数：{len(results)} 条",
        f"- 负面：{sum(v for (c, s), v in counts.items() if s == '负面')} 条",
        f"- 正面：{sum(v for (c, s), v in counts.items() if s == '正面')} 条",
        f"- 中性：{sum(v for (c, s), v in counts.items() if s == '中性')} 条",
        "",
        "| 优先级 | 分类 | 情绪 | 顾客反馈 | 建议动作 | 判定方式 |",
        "|---|---|---|---|---|---|",
    ]
    order = {"高": 0, "中": 1, "低": 2}
    for r in sorted(results, key=lambda x: (order[x["priority"]], order.get(x["sentiment"], 2))):
        lines.append(
            "| {priority} | {category} | {sentiment} | {text} | {suggested_action} | {method} |".format(**r)
        )
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="顾客反馈分类与优先级助手")
    parser.add_argument("--input", default=str(DEFAULT_INPUT), help="反馈文本文件（每行一条，# 开头忽略）")
    parser.add_argument("--api-key", default=os.environ.get("DEEPSEEK_API_KEY", ""), help="大模型 API key（可选）")
    parser.add_argument("--base-url", default="https://api.deepseek.com", help="OpenAI 兼容 API 地址")
    parser.add_argument("--model", default="deepseek-chat", help="模型名")
    args = parser.parse_args()

    texts = [
        line.strip()
        for line in open(args.input, encoding="utf-8-sig")
        if line.strip() and not line.lstrip().startswith("#")
    ]

    use_llm = bool(args.api_key)
    results = []
    for t in texts:
        if use_llm:
            try:
                r = llm_classify(t, args.base_url, args.api_key, args.model)
                r["method"] = "大模型"
            except Exception:
                r = rule_classify(t)
                r["method"] = "规则兜底"
        else:
            r = rule_classify(t)
            r["method"] = "规则兜底"
        r["text"] = t
        results.append(r)

    current_date = date.today()
    markdown = render_markdown(results, current_date)
    OUT_DIR.mkdir(exist_ok=True)
    md_path = OUT_DIR / f"反馈分类结果_{current_date.isoformat()}.md"
    json_path = OUT_DIR / f"反馈分类结果_{current_date.isoformat()}.json"
    md_path.write_text(markdown, encoding="utf-8")
    json_path.write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(markdown)
    print(f"\n[已生成] {md_path}")
    print(f"[已生成] {json_path}")


if __name__ == "__main__":
    main()
