import subprocess
import os


KEYPO_CLI_PATH = "../KEYPO_API/keypo_cli.py"


def fetch_keypo_data(keyword, start_date, end_date):
    os.makedirs("data", exist_ok=True)

    tasks = {
        "sentidist": "data/sentidist.json",
        "hotkw": "data/hotkw.json",
        "hotrank": "data/hotrank.json",
        "freqdist": "data/freqdist.json",
    }

    for function_name, output_path in tasks.items():
        print(f"抓取 KEYPO：{function_name}")

        command = [
            "py",
            KEYPO_CLI_PATH,
            function_name,
            "--q",
            keyword,
            "--min",
            start_date,
            "--max",
            end_date,
            "--output",
            output_path
        ]

        subprocess.run(command, check=True)

    print("✅ KEYPO 資料抓取完成")