#!/usr/bin/env python3
"""
企业授信报告系统 — Python 后端代理
在服务端调用企查查 QCC MCP API（6 大 Server），前端通过本地 API 获取数据。

QCC MCP 6 大 Server:
  1. qcc-company   — 工商信息、财务数据、年报、股东、对外投资
  2. qcc-risk      — 风险扫描
  3. qcc-ipr       — 知识产权（软著、专利）
  4. qcc-operation — 经营信息（招投标、资质）
  5. qcc-executive — 高管/董监高信息
  6. qcc-history   — 历史沿革（工商变更等）
"""

import http.server
import json
import urllib.request
import urllib.error
import os
import sys
import ssl
from concurrent.futures import ThreadPoolExecutor, as_completed

# ============ QCC MCP 配置 ============
QCC_SERVERS = {
    "company": {
        "url": "https://agent.qcc.com/mcp/company/stream",
        "token": "Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ"
    },
    "risk": {
        "url": "https://agent.qcc.com/mcp/risk/stream",
        "token": "Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ"
    },
    "ipr": {
        "url": "https://agent.qcc.com/mcp/ipr/stream",
        "token": "Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ"
    },
    "operation": {
        "url": "https://agent.qcc.com/mcp/operation/stream",
        "token": "Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ"
    },
    "executive": {
        "url": "https://agent.qcc.com/mcp/executive/stream",
        "token": "Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ"
    },
    "history": {
        "url": "https://agent.qcc.com/mcp/history/stream",
        "token": "Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ"
    }
}


def call_qcc_mcp(server, tool_name, args, timeout=30):
    """调用 QCC MCP 单个工具"""
    config = QCC_SERVERS.get(server)
    if not config:
        raise Exception(f"未知的 QCC 服务: {server}")

    request_body = {
        "jsonrpc": "2.0",
        "id": int(os.urandom(4).hex(), 16),
        "method": "tools/call",
        "params": {
            "name": tool_name,
            "arguments": args
        }
    }

    print(f"[QCC] 调用 {server}/{tool_name}...")

    req = urllib.request.Request(
        config["url"],
        data=json.dumps(request_body).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": config["token"]
        },
        method="POST"
    )

    # 跳过 SSL 证书验证（某些环境可能需要）
    ctx = ssl.create_default_context()

    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            text = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        raise Exception(f"QCC {server} HTTP {e.code}: {e.reason}")
    except Exception as e:
        raise Exception(f"QCC {server} 网络错误: {str(e)}")

    # 解析 SSE 格式响应
    lines = text.split("\n")
    for line in lines:
        if line.startswith("data: "):
            try:
                sse_data = json.loads(line[6:])
                if "error" in sse_data:
                    raise Exception(f"QCC {server}/{tool_name}: {json.dumps(sse_data['error'])}")
                result = sse_data.get("result")
                if result and result.get("content") and len(result["content"]) > 0:
                    content_text = result["content"][0]["text"]
                    try:
                        return json.loads(content_text)
                    except (json.JSONDecodeError, ValueError):
                        return content_text
                return result
            except json.JSONDecodeError:
                continue
            except Exception as e:
                if str(e).startswith("QCC "):
                    raise
                continue
    return None


def query_company_data(company_name):
    """查询企业全部数据（12 个工具并行）"""
    print(f"\n{'='*50}")
    print(f"[API] 查询企业: {company_name}")
    print(f"{'='*50}")

    data = {
        "company": None,
        "risk": None,
        "financial": None,
        "annualReports": None,
        "shareholders": None,
        "investments": None,
        "ipr": None,
        "operation": None,
        "executives": None,
        "history": None,
        "errors": []
    }

    # 定义所有查询任务
    tasks = [
        # ===== company server =====
        ("company", "get_company_registration_info", {"searchKey": company_name}, "company", "工商信息"),
        ("company", "get_financial_data", {"searchKey": company_name}, "financial", "财务数据"),
        ("company", "get_annual_reports", {"searchKey": company_name}, "annualReports", "企业年报"),
        ("company", "get_shareholder_info", {"searchKey": company_name}, "shareholders", "股东信息"),
        ("company", "get_external_investments", {"searchKey": company_name}, "investments", "对外投资"),
        # ===== risk server =====
        ("risk", "get_company_risk_scan", {"searchKey": company_name}, "risk", "风险扫描"),
        # ===== ipr server =====
        ("ipr", "get_software_copyright_info", {"searchKey": company_name}, "ipr_software", "软件著作权"),
        ("ipr", "get_patent_info", {"searchKey": company_name}, "ipr_patents", "专利信息"),
        # ===== operation server =====
        ("operation", "get_bidding_info", {"searchKey": company_name}, "op_bidding", "招投标信息"),
        ("operation", "get_qualifications", {"searchKey": company_name}, "op_qualifications", "资质证书"),
        # ===== executive server =====
        ("executive", "get_executive_info", {"searchKey": company_name}, "executives", "高管信息"),
        # ===== history server =====
        ("history", "get_company_change_history", {"searchKey": company_name}, "history", "历史沿革"),
    ]

    def execute_task(server, tool, args, key, label):
        try:
            result = call_qcc_mcp(server, tool, args)
            has_data = result is not None and (not isinstance(result, dict) or len(result) > 0)
            status = "已获取" if has_data else "无记录"
            print(f"  ✅ {label}: {status}")
            return (key, result, None)
        except Exception as e:
            err_msg = str(e)
            print(f"  ❌ {label}: {err_msg}")
            return (key, None, f"{label}: {err_msg}")

    # 并行执行
    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = {
            executor.submit(execute_task, server, tool, args, key, label): key
            for server, tool, args, key, label in tasks
        }
        for future in as_completed(futures):
            key, result, error = future.result()
            if error:
                data["errors"].append(error)
            else:
                # 处理特殊聚合 key
                if key == "ipr_software":
                    if data["ipr"] is None:
                        data["ipr"] = {}
                    data["ipr"]["software"] = result
                elif key == "ipr_patents":
                    if data["ipr"] is None:
                        data["ipr"] = {}
                    data["ipr"]["patents"] = result
                elif key == "op_bidding":
                    if data["operation"] is None:
                        data["operation"] = {}
                    data["operation"]["bidding"] = result
                elif key == "op_qualifications":
                    if data["operation"] is None:
                        data["operation"] = {}
                    data["operation"]["qualifications"] = result
                else:
                    data[key] = result

    print(f"[API] 查询完成，{len(data['errors'])} 个错误\n")
    return data


class APIHandler(http.server.SimpleHTTPRequestHandler):
    """自定义 HTTP 处理器，支持 API 路由 + 静态文件"""

    def __init__(self, *args, **kwargs):
        # 静态文件根目录设为上级目录（index.html 所在位置）
        super().__init__(*args, directory=os.path.dirname(os.path.abspath(__file__)) + "/..", **kwargs)

    def do_POST(self):
        if self.path == "/api/query":
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                params = json.loads(body)
                company_name = params.get("companyName", "").strip()

                if not company_name:
                    self._send_json(400, {"success": False, "error": "缺少企业名称"})
                    return

                data = query_company_data(company_name)
                self._send_json(200, {"success": True, "data": data})

            except json.JSONDecodeError:
                self._send_json(400, {"success": False, "error": "请求体 JSON 解析失败"})
            except Exception as e:
                self._send_json(500, {"success": False, "error": str(e)})
        else:
            self._send_json(404, {"success": False, "error": "Not Found"})

    def do_GET(self):
        if self.path == "/api/health":
            self._send_json(200, {"status": "ok", "time": self.date_time_string()})
        else:
            # 静态文件服务
            super().do_GET()

    def _send_json(self, status_code, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def log_message(self, format, *args):
        # 简化日志
        if "/api/" in str(args[0]):
            print(f"[HTTP] {args[0]}")
        else:
            pass  # 忽略静态文件请求日志


def main():
    port = int(os.environ.get("PORT", 3000))
    server = http.server.HTTPServer(("0.0.0.0", port), APIHandler)

    print(f"\n{'='*50}")
    print(f"  企业授信报告系统 — Python 后端代理")
    print(f"{'='*50}")
    print(f"  前端页面: http://localhost:{port}/credit-report/index.html")
    print(f"  API 端点: http://localhost:{port}/api/query")
    print(f"  健康检查: http://localhost:{port}/api/health")
    print(f"")
    print(f"  QCC MCP 6 大 Server:")
    print(f"    company | risk | ipr | operation | executive | history")
    print(f"  共 12 个查询工具并行调用")
    print(f"{'='*50}\n")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务器已停止")
        server.shutdown()


if __name__ == "__main__":
    main()
