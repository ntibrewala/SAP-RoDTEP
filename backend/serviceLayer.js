const https = require("https");
const http = require("http");
const { URL } = require("url");
const { loadSAPConfig } = require("./sapConfig");

let b1Session = "";
let routeId = "";
let isLoggedIn = false;
let activeCredentials = null;

function rawRequest(method, fullUrl, bodyObj = null, cookies = "", extraHeaders = {}) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(fullUrl);
    const isHttps = parsedUrl.protocol === "https:";
    const transport = isHttps ? https : http;

    const bodyStr = bodyObj !== null ? JSON.stringify(bodyObj) : null;

    const headers = {
      "User-Agent": "Mozilla/5.0",
      "Accept": "application/json",
      ...extraHeaders
    };

    if (bodyStr !== null) {
      headers["Content-Type"] = "application/json";
      headers["Content-Length"] = Buffer.byteLength(bodyStr, "utf8");
    }

    if (cookies) {
      headers["Cookie"] = cookies;
    }

    const options = {
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || (isHttps ? 50000 : 80),
      path: parsedUrl.pathname + parsedUrl.search,
      method,
      headers,
      rejectUnauthorized: false
    };

    const req = transport.request(options, (res) => {
      const chunks = [];
      res.on("data", (chunk) => chunks.push(chunk));
      res.on("end", () => {
        const rawBody = Buffer.concat(chunks).toString("utf8");
        const setCookieHeaders = res.headers["set-cookie"] || [];
        let parsedBody = null;
        if (rawBody) {
          try {
            parsedBody = JSON.parse(rawBody);
          } catch (_) {
            parsedBody = rawBody;
          }
        }
        resolve({
          statusCode: res.statusCode,
          headers: res.headers,
          setCookies: setCookieHeaders,
          body: parsedBody
        });
      });
    });

    req.on("error", (err) => reject(err));
    if (bodyStr !== null) {
      req.write(bodyStr);
    }
    req.end();
  });
}

function applySessionCookies(res) {
  if (res.body && res.body.SessionId) {
    b1Session = res.body.SessionId;
  }

  for (const cookieStr of res.setCookies || []) {
    const b1Match = cookieStr.match(/B1SESSION=([^;]+)/);
    if (b1Match) b1Session = b1Match[1];
    const routeMatch = cookieStr.match(/ROUTEID=([^;]+)/);
    if (routeMatch) routeId = routeMatch[1];
  }

  isLoggedIn = true;
}

async function login(force = false, customCredentials = null) {
  if (customCredentials && customCredentials.userName && customCredentials.password) {
    activeCredentials = customCredentials;
    force = true;
  }

  if (isLoggedIn && b1Session && !force) {
    return;
  }

  const cfg = loadSAPConfig(true);
  const loginUrl = `${cfg.ServiceLayer.url}/Login`;
  const primaryPayload = {
    CompanyDB: cfg.companyDB,
    UserName: activeCredentials?.userName || cfg.SAP.userName,
    Password: activeCredentials?.password || cfg.SAP.password
  };

  let res = await rawRequest("POST", loginUrl, primaryPayload);

  if ((res.statusCode < 200 || res.statusCode >= 300) && activeCredentials && cfg.SAP.userName) {
    const fallbackPayload = {
      CompanyDB: cfg.companyDB,
      UserName: cfg.SAP.userName,
      Password: cfg.SAP.password
    };
    res = await rawRequest("POST", loginUrl, fallbackPayload);
  }

  if (res.statusCode < 200 || res.statusCode >= 300) {
    const errMsg = res.body?.error?.message?.value || JSON.stringify(res.body);
    throw new Error(`SAP Service Layer Login failed (${res.statusCode}): ${errMsg}`);
  }

  applySessionCookies(res);
}

async function request(method, endpoint, payload = null, retry = true, extraHeaders = {}, customCredentials = null) {
  await login(false, customCredentials);
  const cfg = loadSAPConfig();
  const cleanEndpoint = endpoint.startsWith("/") ? endpoint : `/${endpoint}`;
  const fullUrl = `${cfg.ServiceLayer.url}${cleanEndpoint}`;
  const cookieHeader = `B1SESSION=${b1Session}${routeId ? `; ROUTEID=${routeId}` : ""}`;

  const res = await rawRequest(method, fullUrl, payload, cookieHeader, extraHeaders);

  if (res.statusCode === 401 && retry) {
    isLoggedIn = false;
    await login(true, customCredentials);
    return request(method, endpoint, payload, false, extraHeaders, null);
  }

  if (res.statusCode < 200 || res.statusCode >= 300) {
    const errMsg = res.body?.error?.message?.value || JSON.stringify(res.body);
    throw new Error(`SAP Service Layer ${method} ${endpoint} error (${res.statusCode}): ${errMsg}`);
  }

  return res.body;
}

module.exports = {
  login,
  get: (endpoint, extraHeaders) => request("GET", endpoint, null, true, extraHeaders),
  post: (endpoint, payload, extraHeaders, customCredentials) => request("POST", endpoint, payload, true, extraHeaders, customCredentials),
  patch: (endpoint, payload, extraHeaders) => request("PATCH", endpoint, payload, true, extraHeaders),
  put: (endpoint, payload, extraHeaders) => request("PUT", endpoint, payload, true, extraHeaders),
  delete: (endpoint, extraHeaders) => request("DELETE", endpoint, null, true, extraHeaders)
};
