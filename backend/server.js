require("dotenv").config();
const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const { loadSAPConfig } = require("./sapConfig");

const app = express();
app.set("etag", false);
app.use(cors());
app.use(express.json({ limit: "10mb" }));

app.use(["/api", "/rod/api"], (req, res, next) => {
  res.set("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.set("Pragma", "no-cache");
  res.set("Expires", "0");
  next();
});

const rodtepRoutes = require("./routes/rodtep");
app.use("/api/rodtep", rodtepRoutes);
app.use("/rod/api/rodtep", rodtepRoutes);

app.get(["/api/sap-status", "/rod/api/sap-status"], (req, res) => {
  const cfg = loadSAPConfig();
  res.json({
    status: "ok",
    configPath: cfg.configPath,
    schema: cfg.schema,
    hanaServer: cfg.HANA.serverNode,
    serviceLayerUrl: cfg.ServiceLayer.url
  });
});

const candidateBuiltDirs = [
  path.join(__dirname, "../frontend"),
  path.join(__dirname, "../frontendbuilt"),
  path.join(__dirname, "frontendbuilt"),
  path.join(__dirname, "../frontend/dist")
];
const builtDir = candidateBuiltDirs.find((d) => fs.existsSync(path.join(d, "index.html"))) || candidateBuiltDirs[0];

if (fs.existsSync(builtDir)) {
  app.use("/rod", express.static(builtDir));
  app.use(express.static(builtDir));
  app.use((req, res, next) => {
    if (req.path.startsWith("/api/") || req.path.startsWith("/rod/api/")) {
      return next();
    }
    res.sendFile("index.html", { root: builtDir }, (err) => {
      if (err) next(err);
    });
  });
}

const PORT = process.env.PORT || 5065;
const HOST = process.env.HOST || "0.0.0.0";
app.listen(PORT, HOST, (err) => {
  if (err) {
    console.error(`Failed to start RoDTEP Server on ${HOST}:${PORT}:`, err.message);
    process.exit(1);
  }
  const cfg = loadSAPConfig();
  console.log(`RoDTEP Scrip Server running on http://${HOST}:${PORT} and LAN http://10.10.0.15:${PORT}/rod (Schema: ${cfg.schema}, Config: ${cfg.configPath})`);
});
