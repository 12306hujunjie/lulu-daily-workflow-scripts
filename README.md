# Lulu 的日常工作流程脚本项目

用于处理年休假统计等日常流程的自动化脚本集合。

## 本地运行

```bash
poetry install
poetry run python main.py
```

## GitHub Actions 生成 Windows EXE

仓库已配置工作流：

- `.github/workflows/build-windows-exe.yml`

触发后会在 Artifacts 中产出：

- `annual_leave_tool_windows_exe`

其中包含 `annual_leave_tool.exe`。
