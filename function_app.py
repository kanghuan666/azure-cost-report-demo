import azure.functions as func
import datetime
import logging
import os
import io
import time

from azure.core.exceptions import HttpResponseError
from azure.identity import DefaultAzureCredential
from azure.mgmt.costmanagement import CostManagementClient
from azure.storage.blob import BlobServiceClient
import pandas as pd

app = func.FunctionApp()

@app.timer_trigger(schedule="0 0 9 1 * *",
                   arg_name="myTimer",
                   run_on_startup=False,
                   use_monitor=False)
def monthlyCostReport(myTimer: func.TimerRequest) -> None:
    logging.info('月次コストレポート生成を開始します')

    # ===== 1. 環境変数読み込み =====
    subscription_id = os.environ["SUBSCRIPTION_ID"]
    storage_account = os.environ["STORAGE_ACCOUNT_NAME"]
    container_name = os.environ.get("REPORT_CONTAINER_NAME", "reports")

    # ===== 2. 対象期間（先月）の計算 =====
    today = datetime.date.today()
    first_of_this_month = today.replace(day=1)
    last_month_end_date = first_of_this_month - datetime.timedelta(days=1)
    last_month_start_date = last_month_end_date.replace(day=1)
    period_label = last_month_start_date.strftime('%Y-%m')

    # Cost Management API は ISO 8601 datetime（時刻付き）を要求
    last_month_start = datetime.datetime.combine(last_month_start_date, datetime.time.min)
    last_month_end = datetime.datetime.combine(last_month_end_date, datetime.time.max)

    logging.info(f'対象期間: {last_month_start_date} ～ {last_month_end_date}')

    # ===== 3. 認証 =====
    credential = DefaultAzureCredential()

    # ===== 4. Cost Management API 呼び出し =====
    cost_client = CostManagementClient(credential)
    scope = f"/subscriptions/{subscription_id}"

    query = {
        "type": "ActualCost",
        "timeframe": "Custom",
        "timePeriod": {
            "from": last_month_start.isoformat(),
            "to": last_month_end.isoformat()
        },
        "dataset": {
            "granularity": "Daily",
            "aggregation": {
                "totalCost": {"name": "Cost", "function": "Sum"}
            },
            "grouping": [
                {"type": "Dimension", "name": "ServiceName"},
                {"type": "Dimension", "name": "ResourceGroup"}
            ]
        }
    }

    # ===== 4. Cost Management API 呼び出し =====
    cost_client = CostManagementClient(credential)
    scope = f"/subscriptions/{subscription_id}"

    query = {
        "type": "ActualCost",
        "timeframe": "Custom",
        "timePeriod": {
            "from": last_month_start.isoformat(),
            "to": last_month_end.isoformat()
        },
        "dataset": {
            "granularity": "Daily",
            "aggregation": {
                "totalCost": {"name": "Cost", "function": "Sum"}
            },
            "grouping": [
                {"type": "Dimension", "name": "ServiceName"},
                {"type": "Dimension", "name": "ResourceGroup"}
            ]
        }
    }

    # レートリミット対策：429 時は指数バックオフで再試行
    max_retries = 5
    result = None
    for attempt in range(max_retries):
        try:
            result = cost_client.query.usage(scope=scope, parameters=query)
            break
        except HttpResponseError as e:
            if e.status_code == 429 and attempt < max_retries - 1:
                wait_time = 30 * (attempt + 1)
                logging.warning(f'レートリミット (429) に達しました。{wait_time}秒待機して再試行します（{attempt + 1}/{max_retries}）')
                time.sleep(wait_time)
            else:
                raise

    rows = list(result.rows) if result.rows else []
    columns = [c.name for c in result.columns]
    logging.info(f'取得行数: {len(rows)}')

    df = pd.DataFrame(rows, columns=columns)

    # ===== 5. Excel 生成 =====
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if df.empty:
            pd.DataFrame({"メッセージ": ["対象期間にコストデータがありません"]}).to_excel(
                writer, sheet_name='Summary', index=False)
        else:
            df.to_excel(writer, sheet_name='Detail', index=False)
            if 'Cost' in df.columns and 'ServiceName' in df.columns:
                summary = df.groupby('ServiceName')['Cost'].sum().reset_index()
                summary = summary.sort_values('Cost', ascending=False)
                summary.to_excel(writer, sheet_name='Summary', index=False)

    output.seek(0)

    # ===== 6. Blob にアップロード =====
    blob_url = f"https://{storage_account}.blob.core.windows.net"
    blob_service = BlobServiceClient(account_url=blob_url, credential=credential)

    # コンテナが存在しない場合は作成
    container_client = blob_service.get_container_client(container_name)
    try:
        container_client.create_container()
        logging.info(f'コンテナを作成しました: {container_name}')
    except Exception as e:
        # 既に存在する場合は無視（ResourceExistsError）
        logging.info(f'コンテナは既に存在します: {container_name}')

    blob_name = f"cost-report-{period_label}.xlsx"
    blob_client = blob_service.get_blob_client(container=container_name, blob=blob_name)
    blob_client.upload_blob(output.getvalue(), overwrite=True)

    logging.info(f'レポート出力完了: {container_name}/{blob_name}')