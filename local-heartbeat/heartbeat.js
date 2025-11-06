// heartbeat.js
require("dotenv").config();
const msal = require("@azure/msal-node");
const fetch = require("node-fetch");

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const TARGET_UPN = process.env.TARGET_UPN; // 目标用户 OneDrive（你的 E5 账号）

if (!TENANT_ID || !CLIENT_ID || !TARGET_UPN) {
  console.error("请在 .env 中配置 TENANT_ID、CLIENT_ID、TARGET_UPN");
  process.exit(1);
}

// MSAL 配置：设备码登录（用户在浏览器输入验证码登录）
const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
};

const deviceCodeRequest = {
  scopes: ["https://graph.microsoft.com/.default"], // 使用应用已授予的委托权限
  deviceCodeCallback: (response) => {
    console.log("=================================================");
    console.log("请按提示在浏览器中登录你的 E5 账户：");
    console.log(response.message);
    console.log("=================================================");
  },
};

async function getToken() {
  const pca = new msal.PublicClientApplication(msalConfig);
  const authResponse = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
  return authResponse.accessToken;
}

async function callGraph(accessToken) {
  const ts = new Date().toISOString().replace(/[:.]/g, "-");

  // 1. 读取 OneDrive 配额信息
  console.log("读取 OneDrive 配额信息...");
  const driveResp = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      TARGET_UPN
    )}/drive`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!driveResp.ok) {
    const errText = await driveResp.text();
    throw new Error(
      `读取 drive 失败: HTTP ${driveResp.status}\n${errText.slice(0, 500)}`
    );
  }

  const driveJson = await driveResp.json();
  console.log("driveType:", driveJson.driveType);
  console.log("quota used/total:", driveJson.quota?.used, "/", driveJson.quota?.total);

  // 2. 写入 Dev-Heartbeat 下的心跳文件
  console.log("写入 Dev-Heartbeat 心跳文件...");
  const content = [
    "VS Code local Graph heartbeat",
    `timestamp=${ts}`,
    `target_upn=${TARGET_UPN}`,
    `drive_type=${driveJson.driveType}`,
    `quota_used=${driveJson.quota?.used ?? "n/a"}`,
    `quota_total=${driveJson.quota?.total ?? "n/a"}`,
  ].join("\n");

  const uploadUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
    TARGET_UPN
  )}/drive/root:/Dev-Heartbeat/local-heartbeat-${ts}.txt:/content`;

  const putResp = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "text/plain",
    },
    body: content,
  });

  const putBody = await putResp.text();
  console.log("上传 HTTP 状态:", putResp.status);
  if (putResp.ok) {
    console.log("上传成功，返回片段：");
    console.log(putBody.slice(0, 300));
  } else {
    throw new Error(`上传失败: HTTP ${putResp.status}\n${putBody.slice(0, 500)}`);
  }
}

(async () => {
  try {
    console.log("获取 Graph 访问令牌（设备码登录）...");
    const token = await getToken();
    console.log("获取 token 成功，调用 Graph...");
    await callGraph(token);
    console.log("✅ 本次 VS Code 本地心跳完成。");
  } catch (err) {
    console.error("❌ 发生错误：", err.message);
    process.exit(1);
  }
})();
