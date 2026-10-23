#!/usr/bin/env node
// Denticon report robot: runs the group Executive Summary for one day (From = Thru) in a Chrome that keeps its own
// profile (data/denticon-browser, where the user logged into Denticon themselves), exports it as Excel and saves it
// into data/inbox, where the app imports it into the day-by-day history.
//
//   node tools/denticon_robot.js pull [--date YYYY-MM-DD | --from A --to B] [--report exec|referral|both] [--headed]
//   node tools/denticon_robot.js login                 open the window for you to log in (close it when done)
//
// Denticon ends every login after at most 120 minutes, so a pull usually starts with a login: with --headed the
// robot shows the Denticon login page in its window and waits (up to 15 minutes) for the user to log in, then
// carries on. The script never knows or types the password. Files hold office totals only (no patient data) but
// stay in data/, which is git-ignored. Status for the app: data/denticon-robot.json.
"use strict";
const fs = require("fs");
const path = require("path");
const os = require("os");

const ROOT = path.resolve(__dirname, "..");
const PROFILE = path.join(ROOT, "data", "denticon-browser");
const INBOX = path.join(ROOT, "data", "inbox");
const STATUS = path.join(ROOT, "data", "denticon-robot.json");
const REPORT_PAGE = "https://a1.denticon.com/ASPX/Reports/GroupRptsDlyRpts.aspx";
const LOGIN_PAGE = "https://www.denticon.com/login";
const REPORT_NAME = "GrpExecSumm2";   // the group Executive Summary; its Excel export is what the app parses

function playwright() {
  const home = os.homedir();
  const candidates = [
    path.join(ROOT, "node_modules", "playwright-core"),
    path.join(home, ".npm-global", "lib", "node_modules", "@playwright", "cli", "node_modules", "playwright-core"),
    path.join(home, ".npm-global", "lib", "node_modules", "playwright-core"),
    "/usr/lib/node_modules/playwright-core",
  ];
  for (const c of candidates) if (fs.existsSync(c)) return require(c);
  throw new Error("playwright-core not found; install @playwright/cli or playwright-core");
}

const args = process.argv.slice(2);
const cmd = args[0] || "pull";
const opt = name => { const i = args.indexOf("--" + name); return i >= 0 ? args[i + 1] : null; };
const flag = name => args.includes("--" + name);

const pad = n => String(n).padStart(2, "0");
const iso = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const us = isoDay => { const [y, m, d] = isoDay.split("-"); return `${m}/${d}/${y}`; };
const yesterday = () => { const d = new Date(); d.setDate(d.getDate() - 1); return iso(d); };
const daysBetween = (a, b) => { const out = []; for (let d = new Date(a + "T12:00:00"); iso(d) <= b; d.setDate(d.getDate() + 1)) out.push(iso(d)); return out; };
const log = (...a) => console.log(new Date().toISOString().slice(11, 19), ...a);

function writeStatus(patch) {
  let cur = {};
  try { cur = JSON.parse(fs.readFileSync(STATUS, "utf8")); } catch {}
  fs.writeFileSync(STATUS, JSON.stringify({ ...cur, ...patch, updated_at: new Date().toISOString() }, null, 2));
}

async function launch(headed) {
  const { chromium } = playwright();
  fs.mkdirSync(PROFILE, { recursive: true });
  return chromium.launchPersistentContext(PROFILE, { channel: "chrome", headless: !headed, viewport: null, acceptDownloads: true,
    args: ["--disable-blink-features=AutomationControlled"] });
}

// The pull steps live in denticon_pull.js and are reloaded for every command, so a fix there needs no new login.
function pullModule() {
  const file = path.join(__dirname, "denticon_pull.js");
  delete require.cache[require.resolve(file)];
  return require(file);
}

(async () => {
  if (cmd === "login") {
    const context = await launch(true);
    const page = context.pages()[0] || await context.newPage();
    await page.goto(LOGIN_PAGE);
    log("log into Denticon in the window, then close the window");
    await new Promise(r => context.on("close", r));
    writeStatus({ needs_login: false, logged_in_at: new Date().toISOString() });
    return;
  }
  if (cmd === "pull") {
    const days = opt("from") && opt("to") ? daysBetween(opt("from"), opt("to")) : [opt("date") || yesterday()];
    if (opt("report")) fs.writeFileSync(path.join(ROOT, "data", "denticon-robot.report"), opt("report"));   // exec | referral | both
    const context = await launch(flag("headed"));
    const page = context.pages()[0] || await context.newPage();
    const ok = await pullModule().pullDays(context, page, days, flag("headed"));
    await context.close();
    process.exit(ok ? 0 : 1);
  }
  if (cmd === "serve") {
    // A standing robot window: stays open (so one login serves many tries), runs the pulls written to
    // data/denticon-robot.cmd ("pull 2026-10-01 2026-10-02" or "quit"), and touches Denticon every 15 minutes.
    const CMD = path.join(ROOT, "data", "denticon-robot.cmd");
    const context = await launch(true);
    let page = context.pages()[0] || await context.newPage();
    await page.goto(REPORT_PAGE, { waitUntil: "domcontentloaded" }).catch(() => {});
    let login = await pullModule().needsLogin(page);
    writeStatus({ serving: true, state: login ? "waiting_login" : "idle", needs_login: login, message: login ? "Log into Denticon in the robot window." : "Robot window open, logged in. Waiting for a pull." });
    let lastTouch = Date.now();
    log("serving; write a command to", CMD);
    while (true) {
      if (fs.existsSync(CMD)) {
        const line = fs.readFileSync(CMD, "utf8").trim(); fs.unlinkSync(CMD);
        const [verb, a, b] = line.split(/\s+/);
        log("command:", line);
        if (verb === "quit") break;
        if (verb === "pull" && a) {
          await pullModule().pullDays(context, page, daysBetween(a, b || a), true);
          page = context.pages()[0] || await context.newPage();
          writeStatus({ serving: true });
          login = await pullModule().needsLogin(page).catch(() => false);
        }
        lastTouch = Date.now();
      }
      if (login && Date.now() - lastTouch > 10 * 1000) {
        // Waiting for a login: look at the open tabs every 10 s (without moving them), so the status flips as soon as someone logs in.
        lastTouch = Date.now();
        const p = context.pages().find(x => /a1\.denticon\.com/i.test(x.url())) || page;
        const title = await p.title().catch(() => "");
        if (/a1\.denticon\.com/i.test(p.url()) && /Welcome,\s*\S/.test(title) && !(await pullModule().needsLogin(p))) {
          login = false; page = p;
          writeStatus({ serving: true, state: "idle", needs_login: false, logged_in_at: new Date().toISOString(), message: "Robot window open, logged in. Waiting for a pull." });
        }
      } else if (Date.now() - lastTouch > 15 * 60 * 1000) {
        lastTouch = Date.now();
        await page.goto(REPORT_PAGE, { waitUntil: "domcontentloaded" }).catch(() => {});
        login = await pullModule().needsLogin(page);
        writeStatus({ serving: true, state: login ? "waiting_login" : "idle", needs_login: login, message: login ? "Log into Denticon in the robot window." : "Robot window open, logged in. Waiting for a pull." });
      }
      await new Promise(r => setTimeout(r, 2000));
    }
    await context.close();
    writeStatus({ serving: false, state: "idle", message: "Robot window closed." });
    return;
  }
  console.error("usage: denticon_robot.js pull [--date D | --from A --to B] [--headed] | serve | login");
  process.exit(2);
})().catch(e => { console.error(e.message); writeStatus({ state: "error", ok: false, message: e.message.split("\n")[0], last_run: new Date().toISOString() }); process.exit(1); });
