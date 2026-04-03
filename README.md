# muling_import.exe

用于桌面端批量导入睦邻人员数据的 `pywebview` 应用。

## 功能简介

- 登录后进入主界面
- 支持“人员管理”批量导入 Excel 数据
- 支持下载导入模板
- 上传过程中显示进度、成功/失败/跳过统计
- 失败日志支持复制

## 项目结构

```text
.
├── api.py                # 接口调用与 Excel 导入逻辑
├── main.py               # pywebview 桌面入口
├── build_windows.sh      # Windows Git Bash 打包脚本
├── requirements.txt      # Python 依赖
├── file/
│   └── addPerson.xlsx    # 导入模板
└── web/
    ├── app.js            # 前端交互逻辑
    ├── login.html        # 登录页
    ├── main.html         # 主页面
    └── style.css         # 样式
```

## Excel 导入说明

- 第 1 行是表头
- 第 2 行是模板示例行
- 第 3 行开始才会被当成真实数据导入
- 当前读取字段：
  - D 列：姓名
  - E 列：手机号
  - F 列：性别
  - G 列：身份证号
  - H 列：登录密码

## 本地运行

```bash
python -m venv venv
source venv/bin/activate
python -m pip install -r requirements.txt
python main.py
```

## Windows 打包

请在 Windows 的 Git Bash 中执行：

```bash
bash build_windows.sh
```

打包完成后会生成：

```text
dist/医护人员批量添加.exe
```

## Git 说明

- 当前仓库建议只提交运行代码、模板文件和打包脚本
- `tests/`、`venv/`、`dist/`、`build/`、`__pycache__/` 等内容已加入 `.gitignore`
