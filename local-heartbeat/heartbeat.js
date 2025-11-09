// heartbeat.js
require("dotenv").config();
const msal = require("@azure/msal-node");
const fetch = require("node-fetch");

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const TARGET_UPN = process.env.TARGET_UPN;
const TEAMS_WEBHOOK_URL = process.env.TEAMS_WEBHOOK_URL || "";

if (!TENANT_ID || !CLIENT_ID || !TARGET_UPN) {
  console.error("❌ 请在 .env 中配置 TENANT_ID、CLIENT_ID、TARGET_UPN");
  process.exit(1);
}

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
};

const deviceCodeRequest = {
  scopes: [
    "User.Read",
    "Files.ReadWrite.All",
    "Mail.Send",
    "Calendars.ReadWrite",
    "offline_access",
  ],
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

async function graphGet(url, token) {
  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const text = await resp.text();
  if (!resp.ok) {
    throw new Error(`GET ${url} 失败: HTTP ${resp.status}\n${text.slice(0, 500)}`);
  }
  return JSON.parse(text);
}

async function graphPost(url, token, body) {
  const resp = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  const text = await resp.text();
  if (!resp.ok) {
    throw new Error(`POST ${url} 失败: HTTP ${resp.status}\n${text.slice(0, 500)}`);
  }
  return text ? JSON.parse(text) : {};
}

async function graphPutText(url, token, content, contentType = "text/plain") {
  const resp = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": contentType,
    },
    body: content,
  });
  const text = await resp.text();
  if (!resp.ok) {
    throw new Error(`PUT ${url} 失败: HTTP ${resp.status}\n${text.slice(0, 500)}`);
  }
  return { status: resp.status, body: text };
}

async function sendTeamsMessage(text) {
  if (!TEAMS_WEBHOOK_URL) {
    console.log("⚠ 未配置 TEAMS_WEBHOOK_URL，跳过 Teams 通知");
    return;
  }
  const payload = {
    "@type": "MessageCard",
    "@context": "https://schema.org/extensions",
    summary: "Local Graph heartbeat",
    themeColor: "0076D7",
    title: "Local Graph heartbeat (VS Code)",
    text,
  };

  const resp = await fetch(TEAMS_WEBHOOK_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!resp.ok) {
    const t = await resp.text();
    console.warn(`⚠ 发送 Teams 消息失败: HTTP ${resp.status}\n${t.slice(0, 300)}`);
  } else {
    console.log("✅ Teams 通知已发送 ，请查收!");
  }
}

async function callGraph(accessToken) {
  const now = new Date();
  const ts = now.toISOString().replace(/[:.]/g, "-");

  // 1. 读取 OneDrive 配额信息
  console.log("读取 OneDrive 配额信息...");
  const driveJson = await graphGet(
    "https://graph.microsoft.com/v1.0/me/drive",
    accessToken
  );
  console.log("driveType:", driveJson.driveType);
  console.log("quota used/total:", driveJson.quota?.used, "/", driveJson.quota?.total);

  // 2. 写入 Dev-Heartbeat 心跳文件
  console.log("写入 Dev-Heartbeat 心跳文件...");
  const heartbeatContent = [
    "VS Code local Graph heartbeat",
    `timestamp=${ts}`,
    `upn=${TARGET_UPN}`,
    `drive_type=${driveJson.driveType}`,
    `quota_used=${driveJson.quota?.used ?? "n/a"}`,
    `quota_total=${driveJson.quota?.total ?? "n/a"}`,
  ].join("\n");

  const heartbeatUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/Dev-Heartbeat/local-heartbeat-${ts}.txt:/content`;
  const putResult = await graphPutText(heartbeatUrl, accessToken, heartbeatContent);
  console.log("心跳文件上传 HTTP 状态:", putResult.status);

  // 3. 发送测试邮件
  console.log("发送测试邮件...");
  const mailBody = {
    message: {
      subject: "[E5 Dev] VS Code 本地 Graph 心跳测试",
      body: {
        contentType: "Text",
        content:
          `这是一封来自 VS Code 本地脚本的测试邮件。\n\n` +
          `时间：${now.toISOString()}\n` +
          `OneDrive used/total: ${driveJson.quota?.used} / ${driveJson.quota?.total}\n`,
      },
      toRecipients: [
        {
          emailAddress: {
            address: TARGET_UPN,
          },
        },
      ],
    },
    saveToSentItems: true,
  };

  await graphPost("https://graph.microsoft.com/v1.0/me/sendMail", accessToken, mailBody);
  console.log("✅ 测试邮件已发送到:", TARGET_UPN);

  // 4. 创建日历事件（当前时间 +10 分钟开始，持续 30 分钟）
  console.log("创建测试日历事件...");
  const start = new Date(Date.now() + 10 * 60 * 1000); // 10 分钟后
  const end = new Date(start.getTime() + 30 * 60 * 1000); // 30 分钟后结束

  const fmt = (d) => d.toISOString().replace(/\.\d{3}Z$/, ""); // 去掉毫秒，Graph 也能接受

  const eventBody = {
    subject: "[E5 Dev] VS Code 本地心跳事件",
    body: {
      contentType: "HTML",
      content:
        `<p>这是 VS Code 本地脚本创建的测试事件。</p>` +
        `<p>时间：${start.toISOString()} ~ ${end.toISOString()}</p>`,
    },
    start: {
      dateTime: fmt(start),
      timeZone: "UTC",
    },
    end: {
      dateTime: fmt(end),
      timeZone: "UTC",
    },
    location: {
      displayName: "VS Code Heartbeat",
    },
    attendees: [
      {
        type: "required",
        emailAddress: {
          address: TARGET_UPN,
        },
      },
    ],
  };

  const eventResp = await graphPost(
    "https://graph.microsoft.com/v1.0/me/events",
    accessToken,
    eventBody
  );
  console.log("✅ 日历事件已创建，id:", eventResp.id);

  // 5. 发送 Teams 消息（通过 Incoming Webhook）
  const teamsText =
    `VS Code 本地 Graph 心跳完成：\n` +
    `- Heartbeat 文件：local-heartbeat-${ts}.txt\n` +
    `- 邮件已发送给：${TARGET_UPN}\n` +
    `- 日历事件：${eventResp.subject}\n` +
    `- OneDrive used/total: ${driveJson.quota?.used} / ${driveJson.quota?.total}\n` +
    `时间：${now.toISOString()}`;

  await sendTeamsMessage(teamsText);
}

(async () => {
  try {
    console.log("获取 Graph 访问令牌（设备码登录）...");
    const token = await getToken();
    console.log("获取 token 成功，调用 Graph...");
    await callGraph(token);
    console.log("✅ 本次 VS Code 本地心跳（OneDrive + 邮件 + 日历 + Teams）完成。");
  } catch (err) {
    console.error("❌ 发生错误发生错误！！！：", err.message);
    process.exit(1);
  }
})();
