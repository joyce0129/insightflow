import subprocess
import os


KEYPO_CLI_PATH = "../KEYPO_API/keypo_cli.py"

# 抽獎類貼文排除字串（KEYPO 布林語法）
LOTTERY_EXCLUSION = (
    "!(抽|抽獎|文內抽|文末抽|就抽|手由|ㄔㄐ|扌由|扌由好禮|文末扌由獎|買再抽|下單抽|簽到抽|"
    "抽出|抽中|抽獎中|抽好禮|(抽 好禮 ~3)|加碼再抽|有機會抽到|得獎名單|中獎名單|獲獎名單|"
    "留言|留言加碼|留言拿|留言送|留言抽|留言揀|留言得|留言獲|(留言&(加碼|拿|送|抽|揀|得|獲))|"
    "留言手由|留言ㄔㄐ|留言扌由|第一則留言|留言與分享貼文|(留言 機會 ~)|留言有機會|留言有好康|"
    "留言有驚喜|留言有彩蛋|留言有禮|留言處有好康|留言處有驚喜|留言處有彩蛋|留言處有禮|"
    "驚喜在留言處|驚喜在第一則留言|請看第一則留言處|"
    "文末抽好禮|文末有驚喜|文末有好康|文末有彩蛋|文末有好禮|文末驚喜|文末好康|文末彩蛋|"
    "文末好禮|文末有禮|文末送禮|文末寵粉|文末抽|文末ㄔㄐ|文末手由|文末扌由|"
    "內文有驚喜|內文有好禮|內文有好康|內文有彩蛋|內文有禮|"
    "一樓有驚喜|一樓有好禮|一樓有好康|一樓有彩蛋|一樓找驚喜|一樓找好禮|一樓找好康|一樓找彩蛋|"
    "一樓抽|一樓ㄔㄐ|一樓手由|一樓扌由|"
    "1樓有驚喜|1樓有好禮|1樓有好康|1樓有彩蛋|1樓找驚喜|1樓找好禮|1樓找好康|1樓找彩蛋|"
    "1樓抽|1樓ㄔㄐ|1樓手由|1樓扌由|"
    "1f有驚喜|1f有好禮|1f有好康|1f有彩蛋|1f找驚喜|1f找好禮|1f找好康|1f找彩蛋|"
    "1f抽|1fㄔㄐ|1f手由|1f扌由|"
    "1F有驚喜|1F有好禮|1F有好康|1F有彩蛋|1F找驚喜|1F找好禮|1F找好康|1F找彩蛋|"
    "1F抽|1Fㄔㄐ|1F手由|1F扌由)"
)


def _build_query(keyword, exclude_lottery=False, custom_exclusion=""):
    query = keyword
    if exclude_lottery:
        query += f" {LOTTERY_EXCLUSION}"
        print("已套用抽獎排除條件")
    if custom_exclusion:
        query += f" {custom_exclusion}"
        print(f"已套用自訂排除條件")
    return query


def fetch_endpoints(keyword, start_date, end_date, output_dir, endpoints,
                     exclude_lottery=False, custom_exclusion=""):
    """依指定的 KEYPO 端點清單抓取資料，輸出到 output_dir/{endpoint}.json"""
    os.makedirs(output_dir, exist_ok=True)

    query = _build_query(keyword, exclude_lottery, custom_exclusion)

    for function_name in endpoints:
        output_path = f"{output_dir}/{function_name}.json"
        print(f"抓取 KEYPO：{function_name}（{keyword}）")

        command = [
            "py",
            KEYPO_CLI_PATH,
            function_name,
            "--q",
            query,
            "--min",
            start_date,
            "--max",
            end_date,
            "--output",
            output_path
        ]

        subprocess.run(command, check=True)


def fetch_keypo_data(keyword, start_date, end_date,
                     exclude_lottery=False, custom_exclusion=""):
    fetch_endpoints(
        keyword, start_date, end_date, "data",
        ["sentidist", "hotkw", "hotrank", "freqdist", "sprdtrnd"],
        exclude_lottery=exclude_lottery,
        custom_exclusion=custom_exclusion
    )
    print("KEYPO 資料抓取完成")


def fetch_volume_only(keyword, start_date, end_date, output_path,
                       exclude_lottery=False, custom_exclusion=""):
    """僅抓取 freqdist（總聲量趨勢），供去年同期比較等輕量查詢使用，
    避免為了一個總數重複抓取全部 5 個端點。"""
    out_folder = os.path.dirname(output_path)
    if out_folder:
        os.makedirs(out_folder, exist_ok=True)

    query = _build_query(keyword, exclude_lottery, custom_exclusion)

    command = [
        "py",
        KEYPO_CLI_PATH,
        "freqdist",
        "--q",
        query,
        "--min",
        start_date,
        "--max",
        end_date,
        "--output",
        output_path
    ]

    subprocess.run(command, check=True)
