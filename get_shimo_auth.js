/**
 * 石墨文档认证信息一键导出脚本
 *
 * 使用方法：
 * 1. 在浏览器中打开石墨文档并确保已登录
 *    https://teamwork.getech.cn/shimo-h5/shimo-edit/e1898e9f4b794a4786fcdfead749736c
 * 2. 按 F12 打开开发者工具 → 切换到 Console 标签
 * 3. 复制粘贴下面整段代码，按回车运行
 * 4. 屏幕上会出现一个文本框，里面就是 SHIMO_AUTH 的值
 * 5. 点击「复制」按钮，粘贴到 GitHub Secrets 或 Streamlit Cloud Secrets 中
 */

(function () {
  "use strict";

  // ── 收集 cookies ──────────────────────────────────────────────
  const cookies = document.cookie.split(";").map((c) => {
    const [name, ...rest] = c.trim().split("=");
    return {
      name: name,
      value: rest.join("="),
      domain: location.hostname,
      path: "/",
    };
  });

  // ── 收集 localStorage ────────────────────────────────────────
  const lsData = {};
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    lsData[key] = localStorage.getItem(key);
  }

  // ── 打包并 Base64 编码 ────────────────────────────────────────
  const bundle = {
    cookies: cookies,
    localStorage: lsData,
    capturedAt: new Date().toISOString(),
    url: location.href,
  };

  const jsonStr = JSON.stringify(bundle);
  const base64 = btoa(unescape(encodeURIComponent(jsonStr)));

  // ── 在页面上显示结果 ──────────────────────────────────────────
  const overlay = document.createElement("div");
  overlay.style.cssText = `
    position: fixed; top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(0,0,0,0.6); z-index: 99999;
    display: flex; align-items: center; justify-content: center;
    font-family: 'Microsoft YaHei', monospace;
  `;

  const box = document.createElement("div");
  box.style.cssText = `
    background: #fff; border-radius: 12px; padding: 24px;
    max-width: 680px; width: 90%; max-height: 80vh;
    overflow-y: auto; box-shadow: 0 8px 32px rgba(0,0,0,0.3);
  `;

  box.innerHTML = `
    <h3 style="margin:0 0 8px;color:#333;">✅ 石墨文档认证信息已导出</h3>
    <p style="margin:0 0 12px;color:#666;font-size:13px;">
      Cookies: <b>${cookies.length}</b> 个 &nbsp;|&nbsp;
      localStorage: <b>${Object.keys(lsData).length}</b> 个 &nbsp;|&nbsp;
      长度: <b>${base64.length}</b> 字符
    </p>

    <p style="margin:16px 0 4px;font-size:13px;color:#333;">
      <b>📋 请复制以下内容，设为 GitHub Secrets 中的 <code>SHIMO_AUTH</code>：</b>
    </p>
    <textarea id="shimo-auth-output" readonly
      style="width:100%;height:80px;font-size:12px;font-family:monospace;
             padding:8px;border:2px solid #1890FF;border-radius:6px;
             resize:vertical;word-break:break-all;"
    >${base64}</textarea>

    <div style="margin-top:12px;display:flex;gap:8px;flex-wrap:wrap;">
      <button id="shimo-copy-btn" style="
        background:#1890FF;color:#fff;border:none;padding:8px 20px;
        border-radius:6px;cursor:pointer;font-size:14px;font-weight:500;
      ">📋 一键复制</button>

      <button id="shimo-close-btn" style="
        background:#eee;color:#333;border:none;padding:8px 20px;
        border-radius:6px;cursor:pointer;font-size:14px;
      ">关闭</button>
    </div>

    <p id="shimo-hint" style="margin:8px 0 0;font-size:12px;color:#27AE60;display:none;">
      ✅ 已复制到剪贴板！
    </p>

    <details style="margin-top:16px;border-top:1px solid #eee;padding-top:12px;">
      <summary style="cursor:pointer;font-size:12px;color:#999;">🔍 查看原始 JSON 数据</summary>
      <pre style="font-size:10px;color:#666;max-height:200px;overflow:auto;
                  background:#f5f5f5;padding:8px;border-radius:4px;margin-top:8px;
                  white-space:pre-wrap;word-break:break-all;"
      >${jsonStr.replace(/</g, "&lt;")}</pre>
    </details>
  `;

  overlay.appendChild(box);
  document.body.appendChild(overlay);

  // ── 按钮事件 ──────────────────────────────────────────────────
  document.getElementById("shimo-copy-btn").onclick = function () {
    const ta = document.getElementById("shimo-auth-output");
    ta.select();
    ta.setSelectionRange(0, 99999);
    navigator.clipboard.writeText(ta.value).then(() => {
      const hint = document.getElementById("shimo-hint");
      hint.style.display = "block";
      setTimeout(() => (hint.style.display = "none"), 3000);
    });
  };

  document.getElementById("shimo-close-btn").onclick = function () {
    overlay.remove();
  };

  overlay.onclick = function (e) {
    if (e.target === overlay) overlay.remove();
  };

  // ── 同时输出到控制台 ──────────────────────────────────────────
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("SHIMO_AUTH (Base64, " + base64.length + " chars):");
  console.log(base64);
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("Cookies: " + cookies.length + " | localStorage: " + Object.keys(lsData).length);
  console.log("复制上面 base64 字符串，粘贴到 GitHub Secrets 的 SHIMO_AUTH 中");
})();