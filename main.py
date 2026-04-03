import json
import os
import shutil
import sys

import webview

from api import login, get_org_id, batch_add_from_excel, SSR_ID_MAP


def resource_path(relative_path):
    """获取资源文件路径，兼容 PyInstaller 打包后的路径"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


class Api:
    """暴露给前端 JS 调用的 Python 方法"""

    def __init__(self):
        self.token = None
        self.client_id = None
        self.org_id = None
        self.org_name = ""
        self.account = ""
        self.window = None

    def set_window(self, window):
        self.window = window

    def do_login(self, account, password, img_code):
        """登录并获取机构信息"""
        try:
            token, client_id = login(account, password, img_code)
            self.token = token
            self.client_id = client_id
            self.account = account

            org_id, org_name = get_org_id(token, client_id)
            self.org_id = org_id
            self.org_name = org_name

            # 切换到主功能页
            self.window.load_url(self._url('web/main.html'))
            return {"success": True}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def get_user_info(self):
        """返回当前登录用户信息"""
        return {
            "account": self.account,
            "org_name": self.org_name,
        }

    def select_excel_file(self):
        """弹出文件选择对话框，返回文件路径"""
        file_types = ('Excel 文件 (*.xlsx;*.xls)', )
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=file_types
        )
        if result and len(result) > 0:
            return result[0]
        return None

    def start_upload(self, role_name, file_path):
        """开始批量上传"""
        ssr_id = SSR_ID_MAP.get(role_name)
        if not ssr_id:
            return {"success": False, "message": f"未知角色: {role_name}"}

        def progress_callback(current, total, name, status, message):
            # 通过 JS 回调更新前端进度
            self.window.evaluate_js(
                "updateProgress("
                f"{current}, {total}, "
                f"{json.dumps(name or '', ensure_ascii=False)}, "
                f"{json.dumps(status or '', ensure_ascii=False)}, "
                f"{json.dumps(message or '', ensure_ascii=False)}"
                ")"
            )

        try:
            stats = batch_add_from_excel(
                token=self.token,
                client_id=self.client_id,
                org_id=self.org_id,
                ssr_id=ssr_id,
                file_path=file_path,
                progress_callback=progress_callback,
            )
            return {"success": True, "stats": stats}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def download_template(self):
        """弹出保存对话框并保存人员导入模板"""
        template_path = resource_path('file/addPerson.xlsx')
        if not os.path.exists(template_path):
            return {"success": False, "message": "模板文件不存在"}

        dialog_type = webview.FileDialog.SAVE if hasattr(webview, 'FileDialog') else webview.SAVE_DIALOG
        file_types = ('Excel 文件 (*.xlsx)', )
        result = self.window.create_file_dialog(
            dialog_type,
            save_filename='addPerson.xlsx',
            file_types=file_types,
        )
        if not result:
            return {"success": False, "cancelled": True}

        save_path = result[0] if isinstance(result, (list, tuple)) else result
        shutil.copyfile(template_path, save_path)
        return {"success": True, "path": save_path}

    def do_logout(self):
        """退出登录，返回登录页"""
        self.token = None
        self.client_id = None
        self.org_id = None
        self.org_name = ""
        self.account = ""
        self.window.load_url(self._url('web/login.html'))

    def _url(self, relative):
        """构造本地文件 URL"""
        path = resource_path(relative)
        # Windows 路径需要转换
        if sys.platform == 'win32':
            return 'file:///' + path.replace('\\', '/')
        return 'file://' + path


def main():
    api = Api()
    login_url = resource_path('web/login.html')

    window = webview.create_window(
        '导入',
        url=login_url,
        js_api=api,
        width=800,
        height=520,
        resizable=True,
        min_size=(600, 400),
    )
    api.set_window(window)
    webview.start(debug=False)


if __name__ == '__main__':
    main()
