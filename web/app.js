// ========== 登录页逻辑 ==========
async function doLogin() {
    const account = document.getElementById('account');
    const password = document.getElementById('password');
    const imgCode = document.getElementById('imgCode');
    const btn = document.getElementById('loginBtn');
    const errorMsg = document.getElementById('errorMsg');

    if (!account || !password) return;

    const acc = account.value.trim();
    const pwd = password.value.trim();
    const code = imgCode ? imgCode.value.trim() : '8888';

    if (!acc || !pwd) {
        errorMsg.textContent = '请输入账号和密码';
        return;
    }

    btn.disabled = true;
    btn.textContent = '登录中...';
    errorMsg.textContent = '';

    try {
        const result = await pywebview.api.do_login(acc, pwd, code);
        if (result.success) {
            // 登录成功，pywebview 后端会切换页面
        } else {
            errorMsg.textContent = result.message || '登录失败';
        }
    } catch (e) {
        errorMsg.textContent = '登录异常: ' + e;
    } finally {
        btn.disabled = false;
        btn.textContent = '登 录';
    }
}

// ========== 主功能页逻辑 ==========
let uploading = false;

// 菜单切换
function switchTab(tab) {
    if (uploading) return;
    document.querySelectorAll('.tab-page').forEach(el => el.style.display = 'none');
    document.querySelectorAll('.menu-item').forEach(el => el.classList.remove('active'));
    const page = document.getElementById('tab-' + tab);
    const menu = document.getElementById('menu-' + tab);
    if (page) page.style.display = 'block';
    if (menu) menu.classList.add('active');
}

// 页面加载时显示用户信息
window.addEventListener('pywebviewready', async function() {
    const userInfoEl = document.getElementById('userInfo');
    if (userInfoEl) {
        try {
            const info = await pywebview.api.get_user_info();
            if (info) {
                userInfoEl.textContent = '账号: ' + info.account + '  |  机构: ' + info.org_name;
            }
        } catch (e) {
            // 忽略，可能在登录页
        }
    }
});

async function selectRole(roleName) {
    if (uploading) return;

    try {
        const filePath = await pywebview.api.select_excel_file();
        if (!filePath) return; // 用户取消了选择

        uploading = true;
        setButtonsDisabled(true);
        showProgress();
        resetProgress();

        const result = await pywebview.api.start_upload(roleName, filePath);
        if (result.success) {
            showSummary(result.stats);
        } else {
            appendLog('fail', '上传失败: ' + result.message);
        }
    } catch (e) {
        appendLog('fail', '异常: ' + e);
    } finally {
        uploading = false;
        setButtonsDisabled(false);
    }
}

async function downloadTemplate() {
    try {
        const result = await pywebview.api.download_template();
        if (!result || result.cancelled) return;
        if (result.success) {
            alert('模板已保存到: ' + result.path);
        } else {
            alert(result.message || '模板下载失败');
        }
    } catch (e) {
        alert('模板下载异常: ' + e);
    }
}

async function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
        return true;
    }

    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.setAttribute('readonly', '');
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);
    textarea.select();

    try {
        return document.execCommand('copy');
    } finally {
        document.body.removeChild(textarea);
    }
}

function setButtonsDisabled(disabled) {
    const btns = document.querySelectorAll('.btn-role');
    btns.forEach(btn => {
        const alwaysDisabled = btn.dataset.alwaysDisabled === 'true';
        btn.disabled = disabled || alwaysDisabled;
    });
}

function showProgress() {
    const section = document.getElementById('progressSection');
    if (section) section.style.display = 'block';
}

function resetProgress() {
    const bar = document.getElementById('progressBar');
    const text = document.getElementById('progressText');
    const percent = document.getElementById('progressPercent');
    const log = document.getElementById('progressLog');
    const summary = document.getElementById('summary');
    if (bar) bar.style.width = '0%';
    if (text) text.textContent = '准备中...';
    if (percent) percent.textContent = '0%';
    if (log) log.innerHTML = '';
    if (summary) summary.style.display = 'none';
}

// 由 Python 后端调用，更新进度
function updateProgress(current, total, name, status, message) {
    const bar = document.getElementById('progressBar');
    const text = document.getElementById('progressText');
    const percent = document.getElementById('progressPercent');

    const pct = Math.round((current / total) * 100);
    if (bar) bar.style.width = pct + '%';
    if (text) text.textContent = '第 ' + current + ' / ' + total + ' 条: ' + name;
    if (percent) percent.textContent = pct + '%';

    appendLog(status, '[' + name + '] ' + message);
}

function appendLog(status, text) {
    const log = document.getElementById('progressLog');
    if (!log) return;

    const div = document.createElement('div');
    div.className = 'log-entry log-' + status;

    const content = document.createElement('span');
    content.className = 'log-text';
    content.textContent = text;
    div.appendChild(content);

    if (status === 'fail') {
        const copyBtn = document.createElement('button');
        copyBtn.type = 'button';
        copyBtn.className = 'log-copy-btn';
        copyBtn.textContent = '复制';
        copyBtn.onclick = async function() {
            const copied = await copyText(text);
            copyBtn.textContent = copied ? '已复制' : '复制失败';
            setTimeout(function() {
                copyBtn.textContent = '复制';
            }, 1200);
        };
        div.appendChild(copyBtn);
    }

    log.appendChild(div);
    log.scrollTop = log.scrollHeight;
}

function showSummary(stats) {
    const summary = document.getElementById('summary');
    if (!summary) return;
    summary.style.display = 'block';
    summary.textContent = '上传完成！共 ' + stats.total + ' 条，成功 ' + stats.success +
        ' 条，失败 ' + stats.fail + ' 条，跳过 ' + stats.skip + ' 条';
}

function doLogout() {
    if (uploading) {
        alert('正在上传中，请等待完成');
        return;
    }
    pywebview.api.do_logout();
}
