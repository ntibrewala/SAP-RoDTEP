const hana = require("@sap/hana-client");
const { loadSAPConfig } = require("./sapConfig");

function getCandidateNodes(primaryNode) {
  const nodes = [
    primaryNode,
    "10.10.0.12:39015",
    "10.10.0.12:30015",
    "10.10.0.112:30015",
    "10.10.0.113:30015"
  ].filter(Boolean);
  return Array.from(new Set(nodes));
}

function tryConnectAndExec(nodes, idx, uid, pwd, schema, sql, params, resolve, reject) {
  if (idx >= nodes.length) {
    return reject(new Error("SAP HANA connection failed across all candidate server nodes."));
  }

  const connOptions = {
    serverNode: nodes[idx],
    uid,
    pwd
  };

  const connection = hana.createConnection();
  connection.connect(connOptions, (connErr) => {
    if (connErr) {
      if (idx + 1 < nodes.length) {
        return tryConnectAndExec(nodes, idx + 1, uid, pwd, schema, sql, params, resolve, reject);
      }
      return reject(new Error(`SAP HANA connection error (${nodes[idx]}): ` + connErr.message));
    }

    connection.exec(`SET SCHEMA "${schema}"`, (schemaErr) => {
      if (schemaErr) {
        connection.disconnect(() => {});
        return reject(new Error("SAP HANA SET SCHEMA error: " + schemaErr.message));
      }

      const cb = (execErr, rows) => {
        connection.disconnect(() => {});
        if (execErr) {
          return reject(new Error("SAP HANA query error: " + execErr.message));
        }
        resolve([rows || []]);
      };

      if (params && params.length > 0) {
        connection.exec(sql, params, cb);
      } else {
        connection.exec(sql, cb);
      }
    });
  });
}

function execHana(sql, params = [], customSchema = null, useReadWrite = false) {
  const cfg = loadSAPConfig();
  const schema = customSchema || cfg.schema;
  const creds = useReadWrite ? cfg.HANA.readWrite : cfg.HANA.readOnly;
  const nodes = getCandidateNodes(cfg.HANA.serverNode);

  return new Promise((resolve, reject) => {
    tryConnectAndExec(nodes, 0, creds.uid, creds.pwd, schema, sql, params, resolve, reject);
  });
}

module.exports = {
  query: (sql, params, schema) => execHana(sql, params, schema, false),
  execute: (sql, params, schema) => execHana(sql, params, schema, true)
};
