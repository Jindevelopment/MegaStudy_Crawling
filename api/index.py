import base64
import json
import sys
import traceback
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from flask import Flask, jsonify, request, send_from_directory
import crawler as core

PUBLIC_DIR = Path(__file__).parent.parent / "public"

app = Flask(__name__)


@app.route("/")
def index():
    return send_from_directory(PUBLIC_DIR, "index.html")


@app.route("/api/teachers")
def get_teachers():
    try:
        teachers = core.load_teachers()
        return jsonify({"teachers": teachers})
    except core.CrawlerNetworkError as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 502
    except ValueError as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"알 수 없는 오류: {e}"}), 500


@app.route("/api/crawl", methods=["POST"])
def crawl():
    try:
        data = json.loads(request.get_data().decode("utf-8"))
        start_date = datetime.strptime(data["start_date"], "%Y/%m/%d").date()
        end_date = datetime.strptime(data["end_date"], "%Y/%m/%d").date()
        teachers = data["teachers"]

        if not teachers:
            return jsonify({"error": "선생님을 한 명 이상 선택해 주세요."}), 400
        if start_date > end_date:
            return jsonify({"error": "시작 날짜가 종료 날짜보다 늦을 수 없습니다."}), 400

        items: list[dict[str, str]] = []
        for teacher in teachers:
            tcd = teacher["code"]
            name = teacher["name"]
            tcd_items = core.crawl_items(start_date, end_date, tcd)
            for item in tcd_items:
                item["선생님 코드"] = tcd
                item["선생님 이름"] = name
            items.extend(tcd_items)

        rows1 = core.analyze1(items)
        rows2 = core.analyze2(items)

        excel_b64 = base64.b64encode(
            core.make_full_excel_bytes(items, rows1, rows2)
        ).decode("ascii")

        display_items = [
            {
                "날짜": it["날짜"],
                "선생님": it["선생님 이름"],
                "구분": it["구분"],
                "강좌명": it["강좌명"],
            }
            for it in sorted(items, key=lambda x: x["날짜"])
        ]

        return jsonify({
            "count": len(items),
            "excel_b64": excel_b64,
            "items": display_items,
            "analysis1": rows1,
            "analysis2": rows2,
        })
    except core.CrawlerNetworkError as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 502
    except (KeyError, ValueError) as e:
        traceback.print_exc()
        return jsonify({"error": f"요청 형식 오류: {e}"}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"크롤링 중 오류: {e}"}), 500
