const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const root = path.resolve(__dirname, "..");
const requiredFiles = [
  "manifest.xml",
  "send-guard.xml",
  "src/config.js",
  "src/config.json",
  "src/shared.js",
  "src/runtime.js",
  "src/runtime-classic.js",
  "src/taskpane.js",
  "src/taskpane.html",
  "src/taskpane.css"
];

for (const relativePath of requiredFiles) {
  const absolutePath = path.join(root, relativePath);
  if (!fs.existsSync(absolutePath)) {
    throw new Error(`Missing required file: ${relativePath}`);
  }
}

for (const relativePath of ["src/config.js", "src/shared.js", "src/runtime.js", "src/runtime-classic.js", "src/taskpane.js"]) {
  execFileSync(process.execPath, ["--check", path.join(root, relativePath)], { stdio: "inherit" });
}

JSON.parse(fs.readFileSync(path.join(root, "src/config.json"), "utf8"));

const manifest = fs.readFileSync(path.join(root, "manifest.xml"), "utf8");
const expectedManifestSnippets = [
  '<Set Name="Mailbox" MinVersion="1.12"/>',
  '<Version>1.0.2.0</Version>',
  '<LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="Block"/>',
  '<Permissions>ReadWriteItem</Permissions>',
  '<Runtime resid="Runtime.Url">',
  '<Override type="javascript" resid="RuntimeClassicJs.Url"/>',
  '<Control xsi:type="Button" id="OpenPane.Button">'
];

for (const snippet of expectedManifestSnippets) {
  if (!manifest.includes(snippet)) {
    throw new Error(`manifest.xml is missing: ${snippet}`);
  }
}

const productionManifest = fs.readFileSync(path.join(root, "send-guard.xml"), "utf8");
if (!productionManifest.includes("https://send-guard.gavin.cloud")) {
  throw new Error("send-guard.xml should contain Send Guard deployment URLs.");
}

if (productionManifest.includes("APP_BASE_URL")) {
  throw new Error("send-guard.xml should not contain APP_BASE_URL placeholders.");
}

if (!productionManifest.includes("?v=1.0.2")) {
  throw new Error("send-guard.xml should include cache-busting resource URLs.");
}

const runtimeJs = fs.readFileSync(path.join(root, "src/runtime.js"), "utf8");
if (!runtimeJs.includes("Office.onReady") || !runtimeJs.includes("Office.actions.associate")) {
  throw new Error("src/runtime.js should initialise Office before associating event handlers.");
}

const commandsHtml = fs.readFileSync(path.join(root, "src/commands.html"), "utf8");
if (!commandsHtml.includes("Office.onReady")) {
  throw new Error("src/commands.html should call Office.onReady.");
}

console.log("Add-in project checks passed.");
