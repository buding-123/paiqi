# app.py 自动计算明天日期、生成次日排期 + 企业微信推送
from flask import Flask, request
import os
import subprocess
import requests
from datetime import datetime, timedelta  # 新增日期计算模块

app = Flask(__name__)

@app.route('/trigger', methods=['POST'])
def trigger_task():
    try:
        # 自动计算明天的日期：今日+1天，格式 YYYY-MM-DD
        tomorrow = datetime.now() + timedelta(days=1)
        tomorrow_str = tomorrow.strftime("%Y-%m-%d")
        print(f"即将生成【{tomorrow_str}】的节目排期")

        # 执行脚本，动态传入明天日期
        subprocess.run(
            ['python3', 'generate_schedule.py', '-d', tomorrow_str],
            check=True
        )

        # 读取最新生成图片
        img_dir = '排期图片输出'
        files = os.listdir(img_dir)
        if not files:
            return "没有生成图片", 500
        
        latest_img = max([os.path.join(img_dir, f) for f in files], key=os.path.getctime)
        img_name = os.path.basename(latest_img)

        # 正确企业微信机器人链接（已删除重复URL）
        webhook_url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=31b39930-e5ae-44f0-912c-12b8c4f82bb8"
        
        # 你的PythonAnywhere用户名
        user_name = "ZYL2026"
        img_url = f"https://{user_name}.pythonanywhere.com/{img_dir}/{img_name}"

        headers = {'Content-Type': 'application/json'}
        data = {
            "msgtype": "news",
            "news": {
                "articles": [
                    {
                        "title": f"📅 {tomorrow_str} 节目排期已生成",
                        "description": f"明日排期表，请查收",
                        "url": img_url,
                        "picurl": img_url
                    }
                ]
            }
        }
        # 发送消息到企业微信群
        requests.post(webhook_url, headers=headers, json=data)

        return f"执行成功！已生成{tomorrow_str}排期并推送群聊", 200
    except Exception as e:
        err_msg = str(e)
        print("任务执行失败：", err_msg)
        return f"执行失败：{err_msg}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)