"""
销售合同业务检查数据核对脚本
对比销售合同业务检查接口与3个相关业务表的数据
用法：
  python compare_contracts.py "合同名1" "合同名2" ...   指定合同名
  python compare_contracts.py --date-range 20260610-20260612  按时间段查询
  python compare_contracts.py                           查询全部合同
"""

import sys
import re
import requests
from datetime import datetime

# ============================================================
# 配置区域 - 只需修改 TOKEN
# ============================================================

TOKEN = 'Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJwa0lkIjoiMTYwMzAyMDU4MTgzMTEzOTMyOSIsInVzZXJfbmFtZSI6IiZBRE1JTiZjZW5ncWlAaXRnLm5ldCIsIm9wZW5JZCI6IjYyMDBjNjI2ZTRiMDYwZjU5ZWMzN2NjMiIsIm9hQ29tcGFueU5hbWUiOiLljqbpl6jlm73otLjmlbDlrZfnp5HmioDmnInpmZDlhazlj7giLCJjbGllbnRfaWQiOiJpdGctZ2F0ZXdheSIsIndvcmtOdW0iOiI1MDQ4NzUiLCJzYXBEZXB0Q29kZSI6IlNLNiIsInNhcERlcHRJZCI6MjAsInNjb3BlIjpbImFsbCJdLCJjb250ZXh0VHlwZSI6ImFkbWluX2NvbnRleHQiLCJjb21wYW55Ijoi5Y6m6Zeo5Zu96LS45pWw5a2X56eR5oqA5pyJ6ZmQ5YWs5Y-4Iiwic2FwQ29tcGFueUNvZGUiOiJDT00wMDAwMjU1IiwiZXhwIjoxNzgxMzQyNzU3LCJqdGkiOiIzNzM2NTc4OC02YmM5LTQ1NzAtOWRjNi0zZjQ2MTU5NmYzMzgiLCJvYURlcHRJZCI6IjEzNDMiLCJjb21wYW55Q29kZSI6IkNPTTAwMDAyNTUiLCJzYXBDb21wYW55SWQiOjcxLCJ0ZWxlcGhvbmUiOiIxODcyMDkzNjA2MSIsImRlcHQiOiLmtYvor5Xnu4QiLCJvYURlcHROYW1lIjoi6ZKi6ZOB5Lia5Yqh5pWw5a2X5YyW6YOoIiwicGhvbmUiOiIxODcyMDkzNjA2MSIsIm5hbWUiOiLmm77nkKoiLCJkZXB0Q29kZSI6IjMwMzQ2MiIsIm9hSWQiOiIxNDcxNyIsInNhcENvbXBhbnlOYW1lIjoi5Y6m6Zeo5Zu96LS45pWw5a2X56eR5oqA5pyJ6ZmQ5YWs5Y-4Iiwic2FwRGVwdE5hbWUiOiLCoOWbvei0uOaVsOenkemDqOmHkeWxnuS4muWKoeaUr-aMgemDqCIsInVzZXJuYW1lIjoiY2VuZ3FpQGl0Zy5uZXQiLCJvYUNvbXBhbnlJZCI6IjIzMSJ9.ae2p8r2la-svMZJUcGwF-5mccB6s7tBOU0ewcE0IBkXyO04NuKmjll3OEKIU56M9NSbbN0JrpLjUQgp6LwzaauMBG0VxurXtkqjzlViRtOU2m6t_YzujDdpZr1-O1HnZ8tyGSn3zfUWBuI8cODNLhCIMFS73_lVkacS5q2QpyBwL9bE67Tb75e-q_V4yFDVDNWKtrOPV1javEZloPXqHbAvNrxwsUMlMn1E1NA36rDXmcxcx0uQ_zsRQm5j0fbMGGCETdvbBCSQBQlo_CaXgzIG46hZNvGARyCwu571j9ARm36v8ChxsscSYMH_iNcRPZkv7NiAXP1X6Z48gyFnLTA'

HEADERS = {
    "accept": "application/json",
    "authorization": TOKEN,
    "content-type": "application/json;charset=UTF-8",
}

WEIGHT_TOLERANCE = 0.001

# ============================================================
# 解析合同号
# ============================================================

def _is_token_expired(msg: str) -> bool:
    """检测接口返回的消息是否表示 TOKEN 过期"""
    keywords = ["token过期", "token已过期", "token expired", "登录过期", "登录已过期",
                "登录失效", "认证失败", "身份验证", "授权过期", "凭证过期", "凭证失效",
                "访问令牌", "access token", "未授权", "unauthorized", "非法令牌",
                "令牌过期", "令牌失效", "会话过期", "会话失效", "session expired"]
    msg_lower = msg.lower()
    return any(kw in msg_lower for kw in keywords)


def get_contract_list() -> list:
    """从命令行参数解析合同列表，支持三种模式：
    1. --date-range 20260610-20260612  按时间段查询合同
    2. "合同名1" "合同名2" ...         指定合同名
    3. 无参数                          查询全部合同
    """
    args = [a.strip() for a in sys.argv[1:] if a.strip()]

    # 模式1：按时间段查询
    if args and args[0] == "--date-range":
        if len(args) < 2:
            print("❌ 请提供时间段，格式：--date-range 20260610-20260612", file=sys.stderr)
            sys.exit(1)
        return fetch_contract_list_by_date(args[1])

    # 模式2：指定合同名
    if args:
        return args

    # 模式3：查询全部
    return []


# ============================================================
# 按时间段查询销售合同列表
# ============================================================

def parse_date_range(date_range_str: str):
    """解析时间段字符串，返回 (start_ts_ms, end_ts_ms)
    支持格式：20260610-20260612 或 20260610
    """
    match = re.match(r'^(\d{8})(?:-(\d{8}))?$', date_range_str)
    if not match:
        raise ValueError(f"时间段格式错误：{date_range_str}，正确格式：20260610-20260612")

    start_str = match.group(1)
    end_str = match.group(2) or start_str  # 如果只传一个日期，开始=结束

    start_dt = datetime.strptime(start_str, "%Y%m%d")
    end_dt = datetime.strptime(end_str, "%Y%m%d")

    # 结束日期取当天 23:59:59.999
    end_dt = end_dt.replace(hour=23, minute=59, second=59, microsecond=999000)

    start_ms = int(start_dt.timestamp() * 1000)
    end_ms = int(end_dt.timestamp() * 1000)

    return start_ms, end_ms


def fetch_contract_list_by_date(date_range_str: str) -> list:
    """根据时间段查询销售合同列表，返回 outContractNo 列表"""
    start_ms, end_ms = parse_date_range(date_range_str)

    url = "https://manage.itgmetals.com/api/itg-es-scroll-webapp/sell-contract/scroll"
    body = {
        "contractNoList": [],
        "outContractNoList": [],
        "sapCodeList": [],
        "orgIdList": [],
        "deptGroupIdList": [],
        "customerIdList": [],
        "budgetContractNoList": [],
        "saleManIdList": [],
        "merchandiserId": [],
        "createAdList": [],
        "sellContractTypeList": [],
        "queryStatus": 9,
        "contractDateStart": start_ms,
        "contractDateEnd": end_ms
    }

    all_records = []
    # 分页拉取所有数据
    page = 1
    while True:
        body["pageNo"] = page
        body["pageSize"] = 200
        resp = requests.post(url, headers=HEADERS, json=body, timeout=30)
        
        # 先检查 HTTP 状态码
        if resp.status_code in [401, 403]:
            print("__TOKEN_EXPIRED__", file=sys.stderr)
            sys.exit(99)
        
        resp.raise_for_status()
        result = resp.json()
        if result.get("code") != 200:
            msg = result.get("msg", "")
            if _is_token_expired(msg):
                print("__TOKEN_EXPIRED__", file=sys.stderr)
                sys.exit(99)
            raise ValueError(f"销售合同列表接口异常: {msg}")

        data = result.get("data") or []
        if not data:
            break
        all_records.extend(data)
        # 如果返回数据不足一页，说明没有更多了
        if len(data) < 200:
            break
        page += 1

    # 提取 outContractNo 并去重
    contract_names = list(dict.fromkeys(
        str(r.get("outContractNo", "")).strip()
        for r in all_records
        if str(r.get("outContractNo", "")).strip()
    ))

    start_display = datetime.fromtimestamp(start_ms / 1000).strftime("%Y-%m-%d")
    end_display = datetime.fromtimestamp(end_ms / 1000).strftime("%Y-%m-%d")
    print(f"📅 查询时间段：{start_display} ~ {end_display}")
    print(f"📋 查到 {len(contract_names)} 个销售合同")

    return contract_names


# ============================================================
# API 请求 - 新的销售合同业务检查接口
# ============================================================

def _make_request(url: str, body: dict, error_msg: str):
    """统一的 HTTP 请求处理函数，包含 TOKEN 过期检测"""
    resp = requests.post(url, headers=HEADERS, json=body, timeout=30)
    
    # 先检查 HTTP 状态码
    if resp.status_code in [401, 403]:
        print("__TOKEN_EXPIRED__", file=sys.stderr)
        sys.exit(99)
    
    resp.raise_for_status()
    result = resp.json()
    if result.get("code") != 200:
        msg = result.get("msg", "")
        if _is_token_expired(msg):
            print("__TOKEN_EXPIRED__", file=sys.stderr)
            sys.exit(99)
        raise ValueError(f"{error_msg}: {msg}")
    return result.get("data") or []


# ============================================================
# API 请求 - 新的销售合同业务检查接口
# ============================================================

def fetch_contract_check(contract_list: list) -> list:
    """新的销售合同业务检查接口"""
    url = "https://manage.itgmetals.com/api/itg-es-scroll-webapp/sell-contract/sell-contract-check-scroll"
    body = {
        "contractNoList": [],
        "outContractNoList": contract_list,
        "sapCodeList": [],
        "customerIdList": [],
        "brandIdList": [],
        "orgIdList": [],
        "deptGroupIdList": [],
        "saleManIdList": [],
        "merchandiserIdList": [],
        "createAdList": [],
        "isGoodsCompleteList": [],
        "isInvoiceCompleteList": [],
        "isMoneyCompleteList": [],
        "istBusinessTagsList": [],
        "istBusinessSegmentationList": [],
        "finalPriceStatusList": [],
        "isSettlePriceFinalApproveList": [],
        "dataScop": None,
        "status": None
    }
    return _make_request(url, body, "销售合同业务检查接口异常")


def fetch_pickup_check(contract_list: list) -> list:
    """接口1：销售提单业务检查汇总表"""
    url = "https://manage.itgmetals.com/api/itg-es-scroll-webapp/sell-pickup/scroll-check"
    body = {
        "dataScop": 0,
        "billNoList": [],
        "sapCodeList": [],
        "sellContractNoList": [],
        "sellOutContractNoList": contract_list,
        "sellContractSapNoList": [],
        "settleCustomerIdList": [],
        "orgIdList": [],
        "deptGroupIdList": [],
        "warehouseIdList": [],
        "pickupTypeList": [],
        "sourceBillNoList": [],
        "goodsCompleteStatusList": [],
        "invoiceCompleteStatusList": [],
        "profitModelIdList": [],
        "saleManIdList": [],
        "merchandiserIdList": [],
        "createAdList": [],
        "auditAdList": [],
        "productAddressSelfNoList": []
    }
    return _make_request(url, body, "销售提单业务检查接口异常")


def fetch_real_pickup_check(contract_list: list) -> list:
    """接口2：销售实提业务检查汇总表"""
    url = "https://manage.itgmetals.com/api/itg-es-scroll-webapp/sell-real-pickup/scroll-check"
    body = {
        "billNoList": [],
        "sapCodeList": [],
        "realPickupTypeList": [],
        "sellContractNoList": [],
        "sellOutContractNoList": contract_list,
        "sellContractSapNoList": [],
        "settleCustomerIdList": [],
        "warehouseIdList": [],
        "orgIdList": [],
        "deptGroupIdList": [],
        "saleManIdList": [],
        "merchandiserIdList": [],
        "createAdList": [],
        "dataScop": 0
    }
    return _make_request(url, body, "销售实提业务检查接口异常")


def fetch_sell_return(contract_list: list) -> list:
    """接口3：销售退货报表 - 需要分页获取所有数据"""
    url = "https://manage.itgmetals.com/api/itg-es-scroll-webapp/sell-return-product/page"
    body = {
        "billNoList": [],
        "sapCodeList": [],
        "sellContractNoList": [],
        "sellContractOutNoList": contract_list,
        "sellContractSapNoList": [],
        "settleCustomerIdList": [],
        "warehouseIdList": [],
        "orgIdList": [],
        "deptGroupIdList": [],
        "createAdList": [],
        "brandIdList": [],
        "placeIdList": [],
        "materialList": [],
        "specList": [],
        "status": "",
        "openLockTypeList": []
    }

    all_records = []
    # 分页拉取所有数据
    page = 1
    while True:
        body["pageNo"] = page
        body["pageSize"] = 200
        resp = requests.post(url, headers=HEADERS, json=body, timeout=30)
        
        # 先检查 HTTP 状态码
        if resp.status_code in [401, 403]:
            print("__TOKEN_EXPIRED__", file=sys.stderr)
            sys.exit(99)
        
        resp.raise_for_status()
        result = resp.json()
        if result.get("code") != 200:
            msg = result.get("msg", "")
            if _is_token_expired(msg):
                print("__TOKEN_EXPIRED__", file=sys.stderr)
                sys.exit(99)
            raise ValueError(f"销售退货报表接口异常: {msg}")

        data = result.get("data") or []
        if not data:
            break
        all_records.extend(data)
        # 如果返回数据不足一页，说明没有更多了
        if len(data) < 200:
            break
        page += 1

    return all_records


# ============================================================
# 数据处理 & 对比
# ============================================================

def build_contract_check_map(records):
    """构建销售合同业务检查数据的映射"""
    result = {}
    for r in records:
        key = str(r.get("outContractNo", "")).strip()
        if key:
            result[key] = {
                "newPickupWeight": float(r.get("newPickupWeight") or 0),
                "newPickupSumAmount": float(r.get("newPickupSumAmount") or 0),
                "newRealPickWeight": float(r.get("newRealPickWeight") or 0),
                "newSellReturnWeight": float(r.get("newSellReturnWeight") or 0)
            }
    return result


def build_pickup_check_map(records):
    """构建销售提单业务检查的 sumWeight 汇总"""
    result = {}
    for r in records:
        key = str(r.get("sellOutContractNo", "")).strip()
        if key:
            result[key] = result.get(key, 0) + float(r.get("sumWeight") or 0)
    return result


def build_real_pickup_check_map(records):
    """构建销售实提业务检查的 sumWeight 汇总"""
    result = {}
    for r in records:
        key = str(r.get("sellOutContractNo", "")).strip()
        if key:
            result[key] = result.get(key, 0) + float(r.get("sumWeight") or 0)
    return result


def build_sell_return_map(records):
    """构建销售退货的 weight 汇总"""
    result = {}
    for r in records:
        key = str(r.get("sellContractOutNo", "")).strip()
        if key:
            result[key] = result.get(key, 0) + float(r.get("weight") or 0)
    return result


def compare_data(contract_check_map, pickup_check_map, real_pickup_check_map, sell_return_map):
    """对比所有数据，返回差异列表"""
    diff_rows = []

    # 获取所有合同号
    all_contracts = set(contract_check_map.keys()) | set(pickup_check_map.keys()) | set(real_pickup_check_map.keys()) | set(sell_return_map.keys())

    for contract_no in sorted(all_contracts):
        row = {"合同名称": contract_no}

        # 从销售合同业务检查接口获取基准值
        check_data = contract_check_map.get(contract_no)
        if check_data:
            row["提单重量(合同检查)"] = round(check_data["newPickupWeight"], 3)
            row["提单金额(合同检查)"] = round(check_data["newPickupSumAmount"], 3)
            row["实提重量(合同检查)"] = round(check_data["newRealPickWeight"], 3)
            row["销售退货重量(合同检查)"] = round(check_data["newSellReturnWeight"], 3)
        else:
            row["提单重量(合同检查)"] = 0
            row["提单金额(合同检查)"] = 0
            row["实提重量(合同检查)"] = 0
            row["销售退货重量(合同检查)"] = 0

        # 从销售提单业务检查获取实际值
        row["提单重量(提单检查)"] = round(pickup_check_map.get(contract_no, 0), 3)

        # 从销售实提业务检查获取实际值
        row["实提重量(实提检查)"] = round(real_pickup_check_map.get(contract_no, 0), 3)

        # 从销售退货报表获取实际值
        row["销售退货重量(退货报表)"] = round(sell_return_map.get(contract_no, 0), 3)

        # 计算差异
        row["提单重量差异"] = round(row["提单重量(合同检查)"] - row["提单重量(提单检查)"], 3)
        row["实提重量差异"] = round(row["实提重量(合同检查)"] - row["实提重量(实提检查)"], 3)
        row["销售退货重量差异"] = round(row["销售退货重量(合同检查)"] - row["销售退货重量(退货报表)"], 3)

        # 判断是否有差异
        has_diff = (
            abs(row["提单重量差异"]) > WEIGHT_TOLERANCE or
            abs(row["实提重量差异"]) > WEIGHT_TOLERANCE or
            abs(row["销售退货重量差异"]) > WEIGHT_TOLERANCE
        )

        if has_diff:
            diff_rows.append(row)

    return diff_rows


def format_weight(val):
    """格式化重量值，去除末尾无意义的0"""
    if val == 0:
        return "0"
    s = f"{val:.3f}"
    return s.rstrip('0').rstrip('.')


# ============================================================
# 主流程
# ============================================================

def main():
    contract_list = get_contract_list()

    print("🔄 正在查询销售合同业务检查数据...")
    contract_check_records = fetch_contract_check(contract_list)
    contract_check_map = build_contract_check_map(contract_check_records)

    print("🔄 正在查询销售提单业务检查数据...")
    pickup_check_records = fetch_pickup_check(contract_list)
    pickup_check_map = build_pickup_check_map(pickup_check_records)

    print("🔄 正在查询销售实提业务检查数据...")
    real_pickup_check_records = fetch_real_pickup_check(contract_list)
    real_pickup_check_map = build_real_pickup_check_map(real_pickup_check_records)

    print("🔄 正在查询销售退货报表数据...")
    sell_return_records = fetch_sell_return(contract_list)
    sell_return_map = build_sell_return_map(sell_return_records)

    print("🔄 正在对比数据...\n")
    diff_rows = compare_data(contract_check_map, pickup_check_map, real_pickup_check_map, sell_return_map)

    # 所有合同
    all_contracts = sorted(set(contract_check_map.keys()) | set(pickup_check_map.keys()) | set(real_pickup_check_map.keys()) | set(sell_return_map.keys()))

    # 输出每个合同的详细数据
    for contract_no in all_contracts:
        check_data = contract_check_map.get(contract_no)
        pickup_weight = round(pickup_check_map.get(contract_no, 0), 3)
        real_pickup_weight = round(real_pickup_check_map.get(contract_no, 0), 3)
        return_weight = round(sell_return_map.get(contract_no, 0), 3)

        if check_data:
            pw_check = round(check_data["newPickupWeight"], 3)
            pa_check = round(check_data["newPickupSumAmount"], 3)
            rpw_check = round(check_data["newRealPickWeight"], 3)
            srw_check = round(check_data["newSellReturnWeight"], 3)
        else:
            pw_check = pa_check = rpw_check = srw_check = 0

        pw_diff = round(pw_check - pickup_weight, 3)
        rpw_diff = round(rpw_check - real_pickup_weight, 3)
        srw_diff = round(srw_check - return_weight, 3)

        has_diff = abs(pw_diff) > WEIGHT_TOLERANCE or abs(rpw_diff) > WEIGHT_TOLERANCE or abs(srw_diff) > WEIGHT_TOLERANCE
        status = "❌" if has_diff else "✅"

        print(f"{status} {contract_no}")
        print(f"   提单重量：合同检查 {format_weight(pw_check)} | 提单检查 {format_weight(pickup_weight)} | 差异 {format_weight(pw_diff)}")
        print(f"   提单金额：合同检查 {format_weight(pa_check)}")
        print(f"   实提重量：合同检查 {format_weight(rpw_check)} | 实提检查 {format_weight(real_pickup_weight)} | 差异 {format_weight(rpw_diff)}")
        print(f"   销售退货：合同检查 {format_weight(srw_check)} | 退货报表 {format_weight(return_weight)} | 差异 {format_weight(srw_diff)}")
        print()

    # 汇总
    print("=" * 50)
    print(f"📊 核对完成：共 {len(all_contracts)} 个合同")
    if diff_rows:
        print(f"❌ 发现 {len(diff_rows)} 条重量差异：")
        for r in diff_rows:
            print(f"   · {r['合同名称']}")
            if abs(r["提单重量差异"]) > WEIGHT_TOLERANCE:
                print(f"     提单重量差异：{format_weight(r['提单重量(合同检查)'])} - {format_weight(r['提单重量(提单检查)'])} = {format_weight(r['提单重量差异'])}")
            if abs(r["实提重量差异"]) > WEIGHT_TOLERANCE:
                print(f"     实提重量差异：{format_weight(r['实提重量(合同检查)'])} - {format_weight(r['实提重量(实提检查)'])} = {format_weight(r['实提重量差异'])}")
            if abs(r["销售退货重量差异"]) > WEIGHT_TOLERANCE:
                print(f"     销售退货重量差异：{format_weight(r['销售退货重量(合同检查)'])} - {format_weight(r['销售退货重量(退货报表)'])} = {format_weight(r['销售退货重量差异'])}")
    else:
        print("✅ 所有合同重量核对一致，无差异！")


if __name__ == "__main__":
    main()
