from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd


SOURCE_XLSX = Path("2025_eps.xlsx")
MAIN_OUTPUT = Path("index.html")
RANKING_OUTPUT = Path("ranking.html")

QUARTERS = [202503, 202506, 202509, 202512, 202603]
QUARTER_LABELS = {
    202503: "2025/03",
    202506: "2025/06",
    202509: "2025/09",
    202512: "2025/12",
    202603: "2026/03",
}
STRICT_RECENT_QUARTERS = [202506, 202509, 202512, 202603]


def parse_company_label(label: str, fallback_code: object) -> tuple[str, str]:
    text = str(label).strip()
    match = re.match(r"^([0-9A-Za-z]+)\s*(.*)$", text)
    if match:
        code = match.group(1).strip()
        name = match.group(2).strip() or text
        return code, name
    if pd.notna(fallback_code):
        return str(int(float(fallback_code))), text
    return text, text


def to_number(value: object) -> float | None:
    if pd.isna(value):
        return None
    return round(float(value), 2)


def calc_sum(values: list[float | None]) -> float | None:
    if any(value is None for value in values):
        return None
    return round(sum(values), 2)


def build_dataset(source_path: Path) -> dict:
    df = pd.read_excel(source_path)
    raw_columns = list(df.columns)
    if len(raw_columns) < 5:
        raise ValueError("Excel 欄位不足，無法建立 EPS 網站。")

    rename_map = {
        raw_columns[0]: "stock_label",
        raw_columns[1]: "quarter",
        raw_columns[2]: "stock_id",
        raw_columns[3]: "industry",
        raw_columns[4]: "eps",
    }
    if len(raw_columns) > 5:
        rename_map[raw_columns[5]] = "profit_growth"

    df = df.rename(columns=rename_map)
    df["quarter"] = pd.to_numeric(df["quarter"], errors="coerce").astype("Int64")
    df["eps"] = pd.to_numeric(df["eps"], errors="coerce")
    df["industry"] = df["industry"].fillna("未分類").astype(str).str.strip()
    df["stock_label"] = df["stock_label"].astype(str).str.strip()
    df = df[df["quarter"].isin(QUARTERS)].copy()
    df = df[df["stock_id"].notna()].copy()

    industry_order: list[str] = []
    company_order: list[str] = []
    companies: dict[str, dict] = {}

    for row in df.itertuples(index=False):
        code, name = parse_company_label(row.stock_label, row.stock_id)
        quarter = int(row.quarter)
        industry = row.industry or "未分類"
        company_key = (
            f"id:{int(float(row.stock_id))}"
            if pd.notna(row.stock_id)
            else f"label:{code}"
        )

        if industry not in industry_order:
            industry_order.append(industry)

        if company_key not in companies:
            companies[company_key] = {
                "code": code,
                "name": name,
                "label": row.stock_label,
                "industry": industry,
                "eps": {str(item): None for item in QUARTERS},
            }
            company_order.append(company_key)

        companies[company_key]["industry"] = industry
        companies[company_key]["eps"][str(quarter)] = to_number(row.eps)

    company_list: list[dict] = []
    latest_count = 0

    for company_key in company_order:
        company = companies[company_key]
        eps_map = company["eps"]
        strict_recent = calc_sum([eps_map[str(item)] for item in STRICT_RECENT_QUARTERS])
        latest_eps = eps_map["202603"]
        fallback_last = latest_eps if latest_eps is not None else eps_map["202503"]
        ranking_recent = calc_sum(
            [eps_map["202506"], eps_map["202509"], eps_map["202512"], fallback_last]
        )
        ranking_quarter = 202603 if latest_eps is not None else 202503 if fallback_last is not None else None
        if latest_eps is not None:
            latest_count += 1

        company["strictRecentEps"] = strict_recent
        company["rankingRecentEps"] = ranking_recent
        company["rankingQuarter"] = QUARTER_LABELS[ranking_quarter] if ranking_quarter else None
        company["hasLatestQuarter"] = latest_eps is not None
        company["latestQuarterFallback"] = latest_eps is None and fallback_last is not None
        company["detailTable"] = [
            {"label": QUARTER_LABELS[item], "value": eps_map[str(item)]} for item in QUARTERS
        ]
        company_list.append(company)

    ranking_list = [company for company in company_list if company["rankingRecentEps"] is not None]
    ranking_list.sort(key=lambda item: (-item["rankingRecentEps"], item["code"]))
    for index, company in enumerate(ranking_list, start=1):
        company["rank"] = index

    grouped = []
    for industry in industry_order:
        industry_companies = [company for company in company_list if company["industry"] == industry]
        industry_companies.sort(
            key=lambda item: (
                item["rankingRecentEps"] is None,
                -(item["rankingRecentEps"] or -999999),
                item["code"],
            )
        )
        grouped.append({"name": industry, "count": len(industry_companies), "companies": industry_companies})

    return {
        "meta": {
            "sourceFile": source_path.name,
            "sourceTimestamp": datetime.fromtimestamp(source_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
            "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "companyCount": len(company_list),
            "industryCount": len(grouped),
            "latestQuarterCompanyCount": latest_count,
            "rankingEligibleCount": len(ranking_list),
            "strictRecentMissingCount": sum(company["strictRecentEps"] is None for company in company_list),
        },
        "groups": grouped,
        "ranking": ranking_list,
    }


def base_style(accent: str, accent_soft: str, accent2: str, accent2_soft: str, bg: str) -> str:
    return f"""
    :root {{
      --bg: {bg}; --panel: rgba(255,255,255,.86); --line: rgba(24,36,44,.12);
      --ink: #18242c; --muted: #5c6972; --accent: {accent}; --accent-soft: {accent_soft};
      --accent-2: {accent2}; --accent-2-soft: {accent2_soft}; --shadow: 0 20px 50px rgba(24,36,44,.08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0; color: var(--ink); line-height: 1.6;
      font-family: "Microsoft JhengHei", "Segoe UI", sans-serif;
      background:
        radial-gradient(circle at top left, var(--accent-soft), transparent 28%),
        radial-gradient(circle at top right, var(--accent-2-soft), transparent 26%),
        linear-gradient(180deg, #faf7f0 0%, var(--bg) 100%);
    }}
    h1, h2, h3 {{ margin: 0; font-family: Georgia, "Times New Roman", serif; }}
    a {{ color: inherit; }}
    .shell {{ width: min(1280px, calc(100% - 36px)); margin: 0 auto; padding: 24px 0 56px; }}
    .hero, .panel {{
      border: 1px solid var(--line); border-radius: 28px; background: var(--panel);
      box-shadow: var(--shadow); backdrop-filter: blur(12px);
    }}
    .hero {{ padding: 32px; }}
    .panel {{ padding: 22px; margin-top: 22px; }}
    .eyebrow {{
      display: inline-flex; padding: 7px 12px; border-radius: 999px; font-size: 12px;
      letter-spacing: .08em; text-transform: uppercase; color: var(--muted); background: rgba(24,36,44,.05);
    }}
    h1 {{ margin-top: 16px; font-size: clamp(34px, 6vw, 68px); line-height: 1; }}
    .lead {{ margin: 16px 0 0; color: var(--muted); max-width: 860px; }}
    .stats {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-top: 22px; }}
    .stat {{ padding: 16px; border-radius: 18px; border: 1px solid var(--line); background: rgba(255,255,255,.7); }}
    .stat span {{ display: block; font-size: 13px; color: var(--muted); margin-bottom: 6px; }}
    .stat strong {{ font-size: 28px; line-height: 1; font-family: Georgia, "Times New Roman", serif; }}
    .actions, .pills {{ display: flex; flex-wrap: wrap; gap: 10px; }}
    .actions {{ margin-top: 22px; }}
    .btn, .pill {{
      display: inline-flex; align-items: center; gap: 8px; text-decoration: none;
      padding: 11px 16px; border-radius: 999px; border: 1px solid var(--line);
    }}
    .btn.primary {{ color: #fffaf4; background: var(--ink); border-color: transparent; }}
    .btn.secondary, .pill {{ background: rgba(255,255,255,.74); }}
    .top {{ display: flex; justify-content: space-between; gap: 12px; align-items: end; margin-bottom: 16px; }}
    .subtle {{ color: var(--muted); font-size: 14px; }}
    .controls {{ display: grid; grid-template-columns: minmax(220px, 1.2fr) auto; gap: 12px; }}
    input, select {{
      width: 100%; padding: 14px 16px; border-radius: 16px; border: 1px solid var(--line);
      background: rgba(255,255,255,.82); color: var(--ink); font-size: 15px;
    }}
    label.toggle {{
      display: inline-flex; align-items: center; gap: 10px; padding: 12px 16px; border-radius: 16px;
      border: 1px solid var(--line); background: rgba(255,255,255,.82);
    }}
    label.toggle input {{ width: 18px; height: 18px; accent-color: var(--accent-2); }}
    .footer-note {{ margin-top: 14px; color: var(--muted); font-size: 13px; }}
    .empty {{ padding: 24px; text-align: center; color: var(--muted); background: rgba(24,36,44,.04); border-radius: 18px; }}
    table {{ width: 100%; border-collapse: collapse; background: rgba(255,255,255,.82); }}
    th, td {{ padding: 12px 10px; text-align: left; border-bottom: 1px solid rgba(24,36,44,.08); }}
    th {{ font-size: 12px; color: var(--muted); text-transform: uppercase; letter-spacing: .06em; }}
    @media (max-width: 960px) {{
      .stats, .controls {{ grid-template-columns: 1fr; }}
      .shell {{ width: min(100% - 22px, 1280px); padding-top: 14px; }}
      .hero, .panel {{ padding: 18px; }}
      .top {{ flex-direction: column; align-items: start; }}
    }}
    """


def main_template(dataset: dict) -> str:
    data_json = json.dumps(dataset, ensure_ascii=False).replace("</", "<\\/")
    style = base_style("#bc5c34", "rgba(188,92,52,.14)", "#1f6c6a", "rgba(31,108,106,.14)", "#f4efe5")
    return f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="UTF-8" /><meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>EPS 知識庫</title><style>{style}
.groups {{ display: grid; gap: 16px; margin-top: 16px; }}
.group {{ padding: 20px; border-radius: 24px; border: 1px solid var(--line); background: rgba(255,255,255,.74); }}
.cards {{ display: grid; grid-template-columns: repeat(3, minmax(0,1fr)); gap: 12px; }}
.card {{ border: 1px solid var(--line); border-radius: 20px; background: rgba(255,251,246,.92); overflow: hidden; }}
.card summary {{ list-style: none; cursor: pointer; padding: 18px; }}
.card summary::-webkit-details-marker {{ display: none; }}
.code {{ display: inline-flex; padding: 5px 10px; border-radius: 999px; background: rgba(24,36,44,.07); color: var(--muted); font-size: 12px; margin-bottom: 10px; }}
.name-line {{ display: flex; justify-content: space-between; gap: 10px; align-items: start; }}
.name-line h3 {{ font-size: 28px; line-height: 1.05; }}
.badge {{ padding: 6px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; white-space: nowrap; }}
.badge.latest {{ color: var(--accent-2); background: var(--accent-2-soft); }}
.badge.fallback {{ color: var(--accent); background: var(--accent-soft); }}
.industry {{ margin-top: 8px; color: var(--muted); font-size: 14px; }}
.metrics {{ display: grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap: 10px; margin-top: 14px; }}
.metric {{ padding: 12px; border-radius: 16px; background: rgba(24,36,44,.04); }}
.metric span {{ display: block; font-size: 12px; color: var(--muted); margin-bottom: 4px; }}
.metric strong {{ font-size: 22px; line-height: 1; font-family: Georgia, "Times New Roman", serif; }}
.detail {{ padding: 0 18px 18px; }}
.summary-row {{ background: rgba(31,108,106,.06); font-weight: 700; }}
.note {{ margin: 10px 2px 0; color: var(--muted); font-size: 13px; }}
@media (max-width: 1100px) {{ .cards {{ grid-template-columns: repeat(2, minmax(0,1fr)); }} }}
@media (max-width: 760px) {{ .cards, .metrics {{ grid-template-columns: 1fr; }} .name-line {{ flex-direction: column; align-items: start; }} .name-line h3 {{ font-size: 24px; }} }}
</style></head>
<body><div class="shell">
<section class="hero">
  <div class="eyebrow">EPS Knowledge Base</div>
  <h1>依產業整理的 EPS 知識庫</h1>
  <p class="lead">主頁依 TSE 產業分組整理，每家公司可展開查看 2025/03 到 2026/03 的單季 EPS。主頁近四季口徑為 2025/06 + 2025/09 + 2025/12 + 2026/03。</p>
  <div class="stats" id="hero-stats"></div>
  <div class="actions"><a class="btn primary" href="ranking.html">查看近四季 EPS 排名</a><a class="btn secondary" href="./{dataset['meta']['sourceFile']}">開啟原始 Excel</a></div>
</section>
<section class="panel"><div class="top"><div><h2>產業導航</h2><div class="subtle">先搜尋，再跳到指定產業。</div></div></div><div class="pills" id="industry-nav"></div></section>
<section class="panel">
  <div class="top"><div><h2>公司總覽</h2><div class="subtle" id="result-caption"></div></div></div>
  <div class="controls"><input id="search" type="search" placeholder="搜尋公司、代號、產業，例如：2330 / 台積電 / 半導體" /><label class="toggle"><input id="latest-only" type="checkbox" /><span>只看已公布 2026/03</span></label></div>
  <div class="groups" id="groups"></div>
  <div class="footer-note">資料來源：<a href="./{dataset['meta']['sourceFile']}">{dataset['meta']['sourceFile']}</a>，Excel 更新時間 {dataset['meta']['sourceTimestamp']}，網站生成時間 {dataset['meta']['generatedAt']}。</div>
</section></div>
<script id="app-data" type="application/json">{data_json}</script>
<script>
const appData = JSON.parse(document.getElementById("app-data").textContent);
const fmt = new Intl.NumberFormat("zh-TW", {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
const formatValue = (value) => value === null || value === undefined ? "—" : fmt.format(value);
const badgeHtml = (company) => company.hasLatestQuarter
  ? '<span class="badge latest">已含 2026/03</span>'
  : company.latestQuarterFallback
    ? '<span class="badge fallback">排名頁改採 2025/03</span>'
    : '<span class="badge fallback">近四季資料不足</span>';
document.getElementById("hero-stats").innerHTML = [
  ["公司家數", appData.meta.companyCount],
  ["產業群組", appData.meta.industryCount],
  ["已公布 2026/03", appData.meta.latestQuarterCompanyCount],
  ["可進排名頁", appData.meta.rankingEligibleCount]
].map(([label, value]) => `<div class="stat"><span>${{label}}</span><strong>${{value.toLocaleString("zh-TW")}}</strong></div>`).join("");
document.getElementById("industry-nav").innerHTML = appData.groups.map((group) => `<a class="pill" href="#group-${{group.name}}">${{group.name}}<span class="subtle">${{group.count}} 家</span></a>`).join("");
function buildCard(company) {{
  const rows = company.detailTable.map((row) => `<tr><td>${{row.label}}</td><td>${{formatValue(row.value)}}</td></tr>`).join("");
  return `<details class="card"><summary><div class="code">${{company.code}}</div><div class="name-line"><div><h3>${{company.name}}</h3></div>${{badgeHtml(company)}}</div><div class="industry">${{company.industry}}</div><div class="metrics"><div class="metric"><span>主頁近四季</span><strong>${{formatValue(company.strictRecentEps)}}</strong></div><div class="metric"><span>排名頁口徑</span><strong>${{formatValue(company.rankingRecentEps)}}</strong></div></div></summary><div class="detail"><table><thead><tr><th>季度</th><th>單季 EPS</th></tr></thead><tbody>${{rows}}<tr class="summary-row"><td>近四季 EPS（2025/06 + 2025/09 + 2025/12 + 2026/03）</td><td>${{formatValue(company.strictRecentEps)}}</td></tr><tr class="summary-row"><td>排名頁口徑（2026/03 缺值時改採 2025/03）</td><td>${{formatValue(company.rankingRecentEps)}}</td></tr></tbody></table><div class="note">排名頁最後採用季度：${{company.rankingQuarter || "資料不足"}}。</div></div></details>`;
}}
function renderGroups() {{
  const query = document.getElementById("search").value.trim().toLowerCase();
  const latestOnly = document.getElementById("latest-only").checked;
  let visibleGroups = 0;
  let visibleCompanies = 0;
  const html = appData.groups.map((group) => {{
    const companies = group.companies.filter((company) => {{
      const haystack = `${{company.code}} ${{company.name}} ${{company.industry}} ${{company.label}}`.toLowerCase();
      return (!query || haystack.includes(query)) && (!latestOnly || company.hasLatestQuarter);
    }});
    if (!companies.length) return "";
    visibleGroups += 1;
    visibleCompanies += companies.length;
    return `<section class="group" id="group-${{group.name}}"><div class="top"><div><h2>${{group.name}}</h2><div class="subtle">${{companies.length}} 家顯示中</div></div></div><div class="cards">${{companies.map(buildCard).join("")}}</div></section>`;
  }}).join("");
  document.getElementById("result-caption").textContent = `目前顯示 ${{visibleGroups}} 個產業、${{visibleCompanies}} 家公司。`;
  document.getElementById("groups").innerHTML = html || '<div class="empty">沒有符合目前搜尋或篩選條件的公司。</div>';
}}
document.getElementById("search").addEventListener("input", renderGroups);
document.getElementById("latest-only").addEventListener("change", renderGroups);
renderGroups();
</script></body></html>"""


def ranking_template(dataset: dict) -> str:
    data_json = json.dumps(dataset, ensure_ascii=False).replace("</", "<\\/")
    style = base_style("#a84b2c", "rgba(168,75,44,.14)", "#1d6f65", "rgba(29,111,101,.14)", "#edf3ef")
    return f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="UTF-8" /><meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>近四季 EPS 排名</title><style>{style}
.table-wrap {{ overflow: auto; border-radius: 22px; border: 1px solid var(--line); background: rgba(255,255,255,.82); }}
table {{ min-width: 980px; }}
thead th {{ position: sticky; top: 0; background: #f4faf6; z-index: 1; }}
tbody tr:hover {{ background: rgba(29,111,101,.04); }}
.rank {{ font-weight: 700; font-family: Georgia, "Times New Roman", serif; }}
.company strong {{ display: block; font-size: 16px; }}
.company span {{ color: var(--muted); font-size: 13px; }}
.tag {{ display: inline-flex; padding: 6px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; }}
.tag.latest {{ color: var(--accent-2); background: var(--accent-2-soft); }}
.tag.fallback {{ color: var(--accent); background: var(--accent-soft); }}
.controls {{ grid-template-columns: minmax(220px, 1.4fr) minmax(180px, .8fr) minmax(160px, .6fr); }}
@media (max-width: 960px) {{ .controls {{ grid-template-columns: 1fr; }} }}
</style></head>
<body><div class="shell">
<section class="hero">
  <div class="eyebrow">EPS Ranking</div>
  <h1>近四季 EPS 排名</h1>
  <p class="lead">排名邏輯：2025/06 + 2025/09 + 2025/12 + 最新一季 EPS。若公司尚未公布 2026/03，最後一季改採 2025/03。</p>
  <div class="actions"><a class="btn primary" href="index.html">回到主頁</a><a class="btn secondary" href="./{dataset['meta']['sourceFile']}">開啟原始 Excel</a></div>
  <div class="stats" id="stats"></div>
</section>
<section class="panel">
  <div class="controls"><input id="search" type="search" placeholder="搜尋公司、代號、產業" /><select id="industry-filter"><option value="">全部產業</option></select><select id="limit-filter"><option value="all">全部排名</option><option value="50">前 50 名</option><option value="100">前 100 名</option><option value="200">前 200 名</option></select></div>
  <div class="table-wrap"><table><thead><tr><th>排名</th><th>公司</th><th>產業</th><th>2025/06</th><th>2025/09</th><th>2025/12</th><th>最後採用季度</th><th>採用 EPS</th><th>近四季 EPS</th></tr></thead><tbody id="ranking-body"></tbody></table></div>
  <div class="footer-note">資料來源：<a href="./{dataset['meta']['sourceFile']}">{dataset['meta']['sourceFile']}</a>，Excel 更新時間 {dataset['meta']['sourceTimestamp']}，網站生成時間 {dataset['meta']['generatedAt']}。</div>
</section></div>
<script id="app-data" type="application/json">{data_json}</script>
<script>
const appData = JSON.parse(document.getElementById("app-data").textContent);
const fmt = new Intl.NumberFormat("zh-TW", {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
const formatValue = (value) => value === null || value === undefined ? "—" : fmt.format(value);
document.getElementById("stats").innerHTML = [
  ["可排名公司", appData.meta.rankingEligibleCount],
  ["已公布 2026/03", appData.meta.latestQuarterCompanyCount],
  ["主頁近四季缺值", appData.meta.strictRecentMissingCount],
  ["產業數", appData.meta.industryCount]
].map(([label, value]) => `<div class="stat"><span>${{label}}</span><strong>${{value.toLocaleString("zh-TW")}}</strong></div>`).join("");
const industries = [...new Set(appData.ranking.map((item) => item.industry))];
document.getElementById("industry-filter").innerHTML += industries.map((industry) => `<option value="${{industry}}">${{industry}}</option>`).join("");
function buildRow(item) {{
  const lastValue = item.rankingQuarter === "2026/03" ? item.eps["202603"] : item.eps["202503"];
  const tagClass = item.hasLatestQuarter ? "tag latest" : "tag fallback";
  const tagLabel = item.hasLatestQuarter ? "最新值" : "2025/03 回補";
  return `<tr><td class="rank">${{item.rank}}</td><td class="company"><strong>${{item.code}} ${{item.name}}</strong><span>${{item.label}}</span></td><td>${{item.industry}}</td><td>${{formatValue(item.eps["202506"])}}</td><td>${{formatValue(item.eps["202509"])}}</td><td>${{formatValue(item.eps["202512"])}}</td><td><span class="${{tagClass}}">${{item.rankingQuarter}} · ${{tagLabel}}</span></td><td>${{formatValue(lastValue)}}</td><td>${{formatValue(item.rankingRecentEps)}}</td></tr>`;
}}
function renderTable() {{
  const query = document.getElementById("search").value.trim().toLowerCase();
  const industry = document.getElementById("industry-filter").value;
  const limit = document.getElementById("limit-filter").value;
  let items = appData.ranking.filter((item) => {{
    const haystack = `${{item.code}} ${{item.name}} ${{item.label}} ${{item.industry}}`.toLowerCase();
    return (!query || haystack.includes(query)) && (!industry || item.industry === industry);
  }});
  if (limit !== "all") items = items.slice(0, Number(limit));
  document.getElementById("ranking-body").innerHTML = items.length
    ? items.map(buildRow).join("")
    : '<tr><td colspan="9" class="empty">沒有符合目前條件的排名資料。</td></tr>';
}}
document.getElementById("search").addEventListener("input", renderTable);
document.getElementById("industry-filter").addEventListener("change", renderTable);
document.getElementById("limit-filter").addEventListener("change", renderTable);
renderTable();
</script></body></html>"""


def main() -> None:
    if not SOURCE_XLSX.exists():
        raise FileNotFoundError(f"找不到來源檔案：{SOURCE_XLSX}")
    dataset = build_dataset(SOURCE_XLSX)
    MAIN_OUTPUT.write_text(main_template(dataset), encoding="utf-8")
    RANKING_OUTPUT.write_text(ranking_template(dataset), encoding="utf-8")
    print(f"Generated {MAIN_OUTPUT} and {RANKING_OUTPUT} from {SOURCE_XLSX}.")


if __name__ == "__main__":
    main()
