import logging
import os
import tempfile
from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from task_manager import TaskManager

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

app: FastAPI = FastAPI()
task_manager: TaskManager = TaskManager()


@app.post("/upload")
async def upload_file(file: UploadFile) -> JSONResponse:
    """接收 Word 文件，创建处理任务"""
    logger.info(f"Received file: {file.filename}")

    # 验证文件类型
    if not file.filename.endswith(".docx"):
        logger.warning(f"Invalid file type: {file.filename}")
        raise HTTPException(status_code=400, detail="仅支持 .docx 文件")

    # 保存到临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        task_id = task_manager.create_task()
        logger.info(f"Task {task_id} created for file {file.filename}")
        return JSONResponse({
            "task_id": task_id,
            "status": "pending",
            "message": "任务已创建"
        })
    finally:
        # 处理完成后删除临时文件
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
            logger.info(f"Temp file {tmp_path} deleted")


@app.get("/task/{task_id}")
def get_task_status(task_id: str) -> dict:
    """查询任务状态"""
    result = task_manager.get_status(task_id)
    if result is None:
        logger.warning(f"Task {task_id} not found")
        raise HTTPException(status_code=404, detail="任务不存在")
    return result


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
