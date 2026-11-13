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

app.use("/api", (req, res, next) => {
  res.set("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.set("Pragma", "no-cache");
  res.set("Expires", "0");
  next();
});

const rodtepRoutes = require("./routes/rodtep");
app.use("/api/rodtep", rodtepRoutes);

app.get("/api/sap-status", (req, res) => {
  const cfg = loadSAPConfig();
  res.json({
    status: "ok",
    configPath: cfg.configPath,
    schema: cfg.schema,
    hanaServer: cfg.HANA.serverNode,
    serviceLayerUrl: cfg.ServiceLayer.url
  });
});

const builtDir = fs.existsSync(path.join(__dirname, "../frontendbuilt/index.html"))
  ? path.join(__dirname, "../frontendbuilt")
  : path.join(__dirname, "../frontend/dist");

if (fs.existsSync(builtDir)) {
  app.use(express.static(builtDir));
  app.use((req, res, next) => {
    res.sendFile("index.html", { root: builtDir }, (err) => {
      if (err) next(err);
    });
  });
}

const PORT = process.env.PORT || 5060;
app.listen(PORT, (err) => {
  if (err) {
    console.error(`Failed to start RoDTEP Server on port ${PORT}:`, err.message);
    process.exit(1);
  }
  const cfg = loadSAPConfig();
  console.log(`RoDTEP Scrip Server running on port ${PORT} (Schema: ${cfg.schema}, Config: ${cfg.configPath})`);
});
