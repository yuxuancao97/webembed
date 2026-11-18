// Installs (or removes) the add-in manifest for PowerPoint on Mac.
//
//   node scripts/sideload.mjs            install, pointing at https://localhost:3000
//   ADDIN_URL=https://me.github.io/webembed node scripts/sideload.mjs
//                                        install, pointing at your hosted copy
//   node scripts/sideload.mjs --remove   uninstall
//   node scripts/sideload.mjs --print    write the (rewritten) manifest to stdout
//
// PowerPoint reads developer manifests from its "wef" folder. Quit and
// reopen PowerPoint after running this.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src = path.join(here, "..", "manifest.xml");
const wef = path.join(os.homedir(), "Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef");
const dest = path.join(wef, "webembed.manifest.xml");

if (process.argv.includes("--remove")) {
  fs.rmSync(dest, { force: true });
  console.log("Removed", dest);
  process.exit(0);
}

let xml = fs.readFileSync(src, "utf8");
const base = (process.env.ADDIN_URL || "").replace(/\/+$/, "");
if (base) {
  if (!base.startsWith("https://")) {
    console.error("ADDIN_URL must start with https://");
    process.exit(1);
  }
  // AppDomain takes an origin only; everything else takes the full base URL.
  xml = xml.replace(/<AppDomain>https:\/\/localhost:3000<\/AppDomain>/, `<AppDomain>${new URL(base).origin}</AppDomain>`);
  xml = xml.replaceAll("https://localhost:3000", base);
}

if (process.argv.includes("--print")) {
  process.stdout.write(xml);
  process.exit(0);
}

fs.mkdirSync(wef, { recursive: true });
fs.writeFileSync(dest, xml);
console.log("Installed", dest);
console.log("Add-in pages load from", base || "https://localhost:3000");
console.log("Quit and reopen PowerPoint, then: Home > Add-ins > More Add-ins > My Add-ins (or Developer Add-ins) > WebEmbed.");
