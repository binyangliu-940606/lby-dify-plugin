import re
import json
import pandas as pd

from collections.abc import Generator
from typing import List, Dict, Tuple, Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from .get_constraint_spec_draft_json import finalize_spec
from .get_constraint_spec_json import build_frame
from .get_dataSet import generate_dataset
from .check_and_loop import validate


class LbyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        #获取参数
        payload_json = json.loads(tool_parameters["payload_json"]) 
        constraint_spec_draft_json = payload_json["constraint_spec_draft_json"]
        

        constraint_spec_json = finalize_spec(constraint_spec_draft_json)
        log = []
        dataset = []
        count = 5
        while True:
            #获取规范后的内容
            frame_json,log = build_frame(constraint_spec_json)
            #生成数据矩阵数据内容
            dataset = generate_dataset(constraint_spec_json, frame_json)
            #校验
            report = validate(constraint_spec_json, dataset,log)

            log.extend(constraint_spec_json.get("warnings",""))
            # log.append(report)

            passed = bool(report["hard"]["passed"] and report["soft"]["passed"])

            count -= 1
            if count==0 or passed:
                break
        
        df = pd.DataFrame(dataset)
            
        # csv_blob 就是可返回的 blob（二进制）
        csv_blob = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        yield self.create_blob_message(
            blob=csv_blob,          # 直接传入 bytes，不需要 base64 编码
            meta={
                # MIME 类型
                "mime_type": "text/csv; charset=utf-8",  
                "filename": "数据矩阵.csv",
            },
        )

        log_str = "建议人工复核（5–10分钟）\r\n"
        for i, item in enumerate(log, start=1):  # start=1 -> 序号从1开始
            log_str += str(i) + ". " + str(item)+"\r\n"

        # 生成 log 的 bytes（建议 utf-8-sig，Windows/Excel/记事本更友好；纯 utf-8 也行）
        data = log_str.encode("utf-8-sig")
        yield self.create_blob_message(
            blob=data,          # 直接传入 bytes，不需要 base64 编码
            meta={
                # MIME 类型
                "mime_type": "text/plain; charset=utf-8",  
                "filename": "数据矩阵说明.txt",
            },
        )



