// The pull steps for the Denticon robot, kept in their own file so the standing robot window (denticon_robot.js
// serve) can reload them without restarting the browser. One day = one Executive Summary (From = Thru), exported
// to Excel by clicking Denticon's own icon, saved into data/inbox.
"use strict";
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const INBOX = path.join(ROOT, "data", "inbox");
const PULLS = path.join(ROOT, "data", "pulls");        // month-to-date (range) reports pulled for comparison, not imported
const RANGE_MARK = path.join(ROOT, "data", "denticon-robot.range");   // present = the next pull is one range report (From..Thru)
const EXPLORE_MARK = path.join(ROOT, "data", "denticon-robot.explore");   // present = look at a page instead of pulling (mapping new reports)
const REPORT_MARK = path.join(ROOT, "data", "denticon-robot.report");     // present = which report to pull ("referral"); default = Executive Summary
const REFERRAL_PAGE = "https://a1.denticon.com/ASPX/NewReports.aspx?RptName=Referral Reports";
const DAILY_PAGE = "https://a1.denticon.com/ASPX/NewReports.aspx?RptName=Daily Reports";
const TXPLAN_PAGE = "https://a1.denticon.com/ASPX/NewReports.aspx?RptName=Treatment Plan Reports";
const ORTHO_REFERRAL_CODE = "ZREFORTH";   // the "Ortho Referral" procedure code under Other

// Open the report viewer the "Print / Preview" popup points at in a plain tab, wait for the report to draw, click an
// export icon in its header and save the download. Shared by the newer report screens.
async function exportFromPopup(context, page, fr, printSelector, iconId, file, label) {
  for (const extra of context.pages()) if (extra !== page) await extra.close().catch(() => {});
  const [popup] = await Promise.all([context.waitForEvent("page", { timeout: 120000 }), clickIn(fr, printSelector)]);
  await popup.waitForLoadState("domcontentloaded").catch(() => {});
  for (let i = 0; i < 15 && !/r2\.denticon\.com/i.test(popup.url()); i++) await popup.waitForTimeout(1000);
  const viewerUrl = popup.url();
  log(`${label}: viewer ${viewerUrl.slice(0, 140)}`);
  await popup.close().catch(() => {});
  if (!/ReportViewer/i.test(viewerUrl)) throw new Error(`${label}: the report did not open the report viewer: ${viewerUrl.slice(0, 100)}`);
  const viewer = await context.newPage();
  await viewer.goto(viewerUrl, { waitUntil: "domcontentloaded", timeout: 120000 });
  await viewer.bringToFront().catch(() => {});
  const start = Date.now();
  while (!(await reportDrawn(viewer))) {
    if (Date.now() - start > 300000) { await viewer.screenshot({ path: path.join(ROOT, "data", "denticon-robot-viewer-fail.png") }).catch(() => {}); throw new Error(`${label}: the report never finished drawing within 5 minutes`); }
    await viewer.waitForTimeout(2000);
  }
  log(`${label}: drawn after ${((Date.now() - start) / 1000).toFixed(0)} s`);
  await viewer.waitForTimeout(2000);
  const header = viewer.frames().find(f => /ReportHeader/i.test(f.url()));
  if (!header) throw new Error("the viewer's header (with the Export icons) was not found");
  const [download] = await Promise.all([viewer.waitForEvent("download", { timeout: 180000 }), header.click("#" + iconId)]);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  await download.saveAs(file + ".part");
  await viewer.close().catch(() => {});
  const body = fs.readFileSync(file + ".part");
  log(`${label}: ${body.length} bytes`);
  if (body.length < 2000) { fs.unlinkSync(file + ".part"); throw new Error(`${label}: the export is too small (${body.length} bytes)`); }
  fs.renameSync(file + ".part", file);
  return file;
}

// Pick all offices in a newer report screen's office dialog (the visible "Select" next to the office box).
async function pickAllOffices(fr) {
  await clickIn(fr, "xpath=//input[@id='txtselOffices']/following::button[1]");
  await fr.waitForSelector("#chkAllOfc", { timeout: 30000 });
  if (!(await fr.isChecked("#chkAllOfc"))) await clickIn(fr, "#chkAllOfc");
  await fr.waitForTimeout(600);
  const left = await fr.$$eval("#availabelist:visible option", o => o.length).catch(() => 0);
  const already = await fr.$$eval("#selectedlist:visible option", o => o.length).catch(() => 0);
  if (left > 0 && already < 10) { await clickIn(fr, "#btnRight:visible"); await fr.waitForTimeout(600); }
  await clickIn(fr, "button.btn-primary:has-text('apply') >> visible=true");
  await fr.waitForTimeout(800);
  const offices = await fr.$eval("#txtselOffices", el => el.value).catch(() => "");
  if (!offices || offices.split(",").length < 2 && !/all/i.test(offices)) throw new Error(`only one office selected (${offices})`);
  return offices;
}

// The Treatment Plan Status (# of ortho referrals): Treatment Plan Reports › Treatment Plan Status, all offices, the
// day, every status, procedure code Other › Ortho Referral (ZREFORTH) only, exported as Excel from the viewer.
async function pullTxPlanStatus(context, page, frm, thru) {
  await page.goto(TXPLAN_PAGE, { waitUntil: "domcontentloaded", timeout: 90000 });
  if (await needsLogin(page)) throw new Error("not logged in");
  let fr = null;
  for (let i = 0; i < 30 && !fr; i++) { fr = page.frames().find(f => /TreatmentPlanReports/i.test(f.url())); if (!fr) await page.waitForTimeout(1000); }
  if (!fr) throw new Error("the Treatment Plan Reports screen did not load");
  await fr.waitForSelector("#rbTreatPlanSts", { timeout: 60000 });
  await clickIn(fr, "#rbTreatPlanSts");
  await fr.waitForSelector("#txtselCode", { timeout: 30000 });
  await clickIn(fr, "#rbReportFormat1");   // Detail
  await clickIn(fr, "#rbPrintFormatPDF");  // the viewer; Excel is exported from there
  await clickIn(fr, "#rbOffice");
  const offices = await pickAllOffices(fr);
  for (const [sel, val] of [["#dtStartDate", us(frm)], ["#dtEndDate", us(thru)]]) {
    await fr.fill(sel, ""); await fr.fill(sel, val); await fr.$eval(sel, el => { el.dispatchEvent(new Event("change", { bubbles: true })); el.blur(); });
  }
  await clickIn(fr, "#rbAll");             // every treatment plan status
  // Procedure code: Other › Ortho Referral only.
  await clickIn(fr, "xpath=//input[@id='txtselCode']/following::button[1]");
  await fr.waitForSelector("#chkAllCodes", { timeout: 30000 });
  if (await fr.isChecked("#chkAllCodes")) await clickIn(fr, "#chkAllCodes");
  await fr.waitForTimeout(600);
  await clickIn(fr, "#cat-other");
  await fr.waitForTimeout(1000);
  // The code dialog's lists share their ids with the office dialog's, so work on the visible ones only.
  const value = await fr.$$eval("#availabelist:visible option", (os, code) => (os.find(o => o.textContent.toUpperCase().includes(code)) || {}).value, ORTHO_REFERRAL_CODE).catch(() => null);
  if (!value) throw new Error(`${ORTHO_REFERRAL_CODE} (Ortho Referral) was not found under Other in the code list`);
  await fr.$eval("#availabelist:visible", (el, v) => { for (const o of el.options) o.selected = o.value === v; el.dispatchEvent(new Event("change", { bubbles: true })); }, value);
  await clickIn(fr, "#btnRight:visible");
  await fr.waitForTimeout(600);
  const chosen = await fr.$$eval("#selectedlist:visible option", os => os.map(o => o.textContent.trim().slice(0, 40)));
  await clickIn(fr, "button.btn-primary:has-text('apply') >> visible=true");
  await fr.waitForTimeout(800);
  const codes = await fr.$eval("#txtselCode", el => el.value).catch(() => "");
  log(`tx plan status ${frm}..${thru}: offices "${offices.slice(0, 30)}", codes "${codes.slice(0, 40)}" (dialog: ${chosen.join("; ").slice(0, 60)})`);
  if (!codes.toUpperCase().includes(ORTHO_REFERRAL_CODE) || codes.split(",").length !== 1) throw new Error(`the code filter is "${codes}", not just ${ORTHO_REFERRAL_CODE}`);
  const dir = frm === thru ? INBOX : PULLS;
  const file = path.join(dir, frm === thru ? `Denticon_TreatPlanStatus ${frm}.xlsx` : `Denticon_TreatPlanStatus ${frm} to ${thru} pulled ${new Date().toISOString().slice(0, 16).replace("T", " ").replace(":", "")}.xlsx`);
  return exportFromPopup(context, page, fr, "button.btn-primary:has-text('Print / Preview')", "ExcelIconImg", file, "tx plan status");
}

// The ortho Daily Journal (ortho collection): Daily Reports › Daily Journal, Detail, all offices, the day, Patient
// Type Ortho, Payments and Adjustments only (no Services), exported as Word, which the app's parser reads.
async function pullOrthoJournal(context, page, frm, thru) {
  await page.goto(DAILY_PAGE, { waitUntil: "domcontentloaded", timeout: 90000 });
  if (await needsLogin(page)) throw new Error("not logged in");
  let fr = null;
  for (let i = 0; i < 30 && !fr; i++) { fr = page.frames().find(f => /DailyReports/i.test(f.url())); if (!fr) await page.waitForTimeout(1000); }
  if (!fr) throw new Error("the Daily Reports screen did not load");
  await fr.waitForSelector("#rbDlyJournal", { timeout: 60000 });
  await clickIn(fr, "#rbDlyJournal");
  await fr.waitForSelector("#orthoPatientType", { timeout: 30000 });
  await clickIn(fr, "#test5");          // Detail
  await clickIn(fr, "#rptPrntPdf");     // the PDF choice opens the viewer; Word is exported from there
  await clickIn(fr, "#test2");          // by Office
  await clickIn(fr, "button.office-btns >> visible=true");
  await fr.waitForSelector("#chkAllOfc", { timeout: 30000 });
  if (!(await fr.isChecked("#chkAllOfc"))) await clickIn(fr, "#chkAllOfc");
  await fr.waitForTimeout(600);
  const left = await fr.$$eval("#availabelist option", o => o.length).catch(() => 0);
  const already = await fr.$$eval("#selectedlist option", o => o.length).catch(() => 0);
  if (left > 0 && already < 10) { await clickIn(fr, "#btnRight"); await fr.waitForTimeout(600); }
  await clickIn(fr, "button.btn-primary:has-text('apply') >> visible=true");
  await fr.waitForTimeout(800);
  await clickIn(fr, "#test3");          // by transaction date
  for (const [sel, val] of [["#dtStartDate", us(frm)], ["#dtEndDate", us(thru)]]) {
    await fr.fill(sel, ""); await fr.fill(sel, val); await fr.$eval(sel, el => { el.dispatchEvent(new Event("change", { bubbles: true })); el.blur(); });
  }
  await clickIn(fr, "#orthoPatientType");
  if (await fr.isChecked("#check1")) await clickIn(fr, "#check1");      // Services off
  if (!(await fr.isChecked("#check2"))) await clickIn(fr, "#check2");   // Payments on
  if (!(await fr.isChecked("#check3"))) await clickIn(fr, "#check3");   // Adjustments on
  const offices = await fr.$eval("#txtselOffices", el => el.value).catch(() => "");
  const ortho = await fr.isChecked("#orthoPatientType");
  log(`ortho journal ${frm}..${thru}: offices "${offices.slice(0, 40)}", patient type ${ortho ? "Ortho" : "NOT ORTHO"}, services off`);
  if (!ortho) throw new Error("the Patient Type did not switch to Ortho");
  if (!offices || offices.split(",").length < 2 && !/all/i.test(offices)) throw new Error(`only one office selected (${offices})`);
  const dir = frm === thru ? INBOX : PULLS;
  const file = path.join(dir, frm === thru ? `Denticon_OrthoDailyJournal ${frm}.doc` : `Denticon_OrthoDailyJournal ${frm} to ${thru} pulled ${new Date().toISOString().slice(0, 16).replace("T", " ").replace(":", "")}.doc`);
  return exportFromPopup(context, page, fr, "button.btn-primary:has-text('Print / Preview')", "WordIconImg", file, "ortho journal");
}

// Click something that may sit inside a frame and may be scrolled out of its box (a script click still runs the handler).
async function clickIn(fr, sel) {
  try { await fr.click(sel, { timeout: 4000 }); } catch { await fr.$eval(sel, el => el.click()); }
}

// The Referral Production Listing (NP collection; grouped by ZIP code for the map), all offices, From..Thru, as Excel.
// Reports › Referral Reports › Referral Production List, in the newer report screen (an app inside a frame).
async function pullReferral(context, page, frm, thru) {
  await page.goto(REFERRAL_PAGE, { waitUntil: "domcontentloaded", timeout: 90000 });
  if (await needsLogin(page)) throw new Error("not logged in");
  let fr = null;
  for (let i = 0; i < 30 && !fr; i++) { fr = page.frames().find(f => /ReferralReports/i.test(f.url())); if (!fr) await page.waitForTimeout(1000); }
  if (!fr) throw new Error("the Referral Reports screen did not load");
  await fr.waitForSelector("#rbReferralProductionList", { timeout: 60000 });
  await clickIn(fr, "#rbReferralProductionList");
  await fr.waitForSelector("#cbGrpByZip", { timeout: 30000 });
  await clickIn(fr, "#rbPrintFormatPDF");   // the PDF choice opens the report viewer, whose Export to Excel gives the formatted report
  await clickIn(fr, "#rbOffice");
  // All offices: the office Select dialog (tick "All", move everything across, apply), the same helper as the others.
  await pickAllOffices(fr);
  // Dates: the procedure dates are the period; the First Visit dates make it a NEW-patient listing (patients whose
  // first visit falls in the window), as the manual monthly report is run (First Visit = From..Thru). For a one-day
  // report the first-visit window is the whole month, so a deposit paid before the first visit still counts and the
  // days of a month add up to the month's final figure.
  const monthEnd = d => { const [y, m] = d.split("-").map(Number); return `${y}-${String(m).padStart(2, "0")}-${String(new Date(y, m, 0).getDate()).padStart(2, "0")}`; };
  const fvFrom = frm === thru ? frm.slice(0, 8) + "01" : frm, fvThru = frm === thru ? monthEnd(frm) : thru;
  for (const [sel, val] of [["#dtProcedureStartDate", us(frm)], ["#dtProcedureEndDate", us(thru)], ["#dtFirstVisitStartDate", us(fvFrom)], ["#dtFirstVisitEndDate", us(fvThru)]]) {
    await fr.fill(sel, ""); await fr.fill(sel, val); await fr.$eval(sel, el => { el.dispatchEvent(new Event("change", { bubbles: true })); el.blur(); });
  }
  if (!(await fr.isChecked("#cbGrpByZip"))) await clickIn(fr, "#cbGrpByZip");
  const offices = await fr.$eval("#txtselOffices", el => el.value).catch(() => "");
  log(`referral ${frm}..${thru}: offices "${offices.slice(0, 60)}", first visit ${us(fvFrom)}..${us(fvThru)}, group by zip`);
  if (!offices || offices.split(",").length < 2 && !/all/i.test(offices)) throw new Error(`only one office selected (${offices}); the all-offices step did not take`);
  // Print / Preview opens the report viewer in a popup; as with the Executive Summary, take its address, open it in a
  // plain tab, wait for the PDF to draw, then click the viewer's Export to Excel and catch the download.
  for (const extra of context.pages()) if (extra !== page) await extra.close().catch(() => {});
  const [popup] = await Promise.all([context.waitForEvent("page", { timeout: 120000 }), clickIn(fr, "button.btn-primary:has-text('Print / Preview')")]);
  await popup.waitForLoadState("domcontentloaded").catch(() => {});
  for (let i = 0; i < 15 && !/r2\.denticon\.com/i.test(popup.url()); i++) await popup.waitForTimeout(1000);
  const viewerUrl = popup.url();
  log("referral: viewer", viewerUrl.slice(0, 140));
  await popup.close().catch(() => {});
  if (!/ReportViewer/i.test(viewerUrl)) throw new Error("the Referral report did not open the report viewer: " + viewerUrl.slice(0, 100));
  const viewer = await context.newPage();
  await viewer.goto(viewerUrl, { waitUntil: "domcontentloaded", timeout: 120000 });
  await viewer.bringToFront().catch(() => {});
  const start = Date.now();
  while (!(await reportDrawn(viewer))) {
    if (Date.now() - start > 300000) { await viewer.screenshot({ path: path.join(ROOT, "data", "denticon-referral-tab.png") }).catch(() => {}); throw new Error("the Referral report never finished drawing within 5 minutes"); }
    await viewer.waitForTimeout(2000);
  }
  log(`referral ${frm}..${thru}: drawn after ${((Date.now() - start) / 1000).toFixed(0)} s`);
  await viewer.waitForTimeout(2000);
  const header = viewer.frames().find(f => /ReportHeader/i.test(f.url()));
  if (!header) throw new Error("the viewer's header (with the Export icons) was not found");
  const [download] = await Promise.all([viewer.waitForEvent("download", { timeout: 180000 }), header.click("#ExcelIconImg")]);
  const dir = frm === thru ? INBOX : PULLS;
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, frm === thru ? `Denticon_ReferralRefProdList ${frm}.xlsx` : `Denticon_ReferralRefProdList ${frm} to ${thru} pulled ${new Date().toISOString().slice(0, 16).replace("T", " ").replace(":", "")}.xlsx`);
  await download.saveAs(file + ".part");
  await viewer.close().catch(() => {});
  const body = fs.readFileSync(file + ".part");
  const ok = body.length > 2000 && body[0] === 0x50 && body[1] === 0x4b;
  log(`referral ${frm}..${thru}: ${body.length} bytes${ok ? "" : " (not an Excel file)"}`);
  if (!ok) { fs.unlinkSync(file + ".part"); throw new Error(`the download is not an Excel file (${body.length} bytes)`); }
  fs.renameSync(file + ".part", file);
  return file;
}

// Mapping aid: open a page, click what the marker says, then log the visible form controls and take a picture.
async function explore(context, page, spec) {
  await page.goto(spec.url, { waitUntil: "domcontentloaded", timeout: 90000 });
  if (await needsLogin(page)) throw new Error("not logged in");
  for (const sel of spec.clicks || []) {   // the control may live inside a frame
    let hit = false;
    for (const fr of page.frames()) {
      if (!(await fr.$(sel).catch(() => null))) continue;
      try { await fr.click(sel, { timeout: 5000 }); }
      catch { await fr.$eval(sel, el => el.click()); }   // the frame scrolls inside a small box; a script click still fires the page's handler
      hit = true; break;
    }
    if (!hit) log("explore: nothing matches", sel);
    await page.waitForTimeout(3000);
  }
  const dump = () => [...document.querySelectorAll("input,select,button,label,a,textarea")]
    .filter(e => e.offsetParent && e.tagName !== "LABEL" && !/^chkUser|^MenuBar_|^PoWMenu_|partialOfficeSearch|txtSessionTimer/.test(e.id || e.name || "") && !(e.tagName === "A" && !e.textContent.trim()))
    .map(e => [e.tagName, e.type || "", e.id || e.name || e.className.split(" ")[0] || "", e.tagName === "SELECT" ? [...e.options].slice(0, 14).map(o => o.text.trim().slice(0, 30)).join("|") : (e.value || e.textContent || e.title || "").trim().slice(0, 50), e.checked === true ? "checked" : "",
      /radio|checkbox/.test(e.type || "") ? (() => { const lab = e.labels && e.labels[0] && e.labels[0].textContent; if (lab && lab.trim()) return lab.trim().slice(0, 60);
        let n = e.nextSibling; while (n && !(n.textContent || "").trim()) n = n.nextSibling; if (n && (n.textContent || "").trim()) return n.textContent.trim().replace(/\s+/g, " ").slice(0, 60);
        const td = e.closest("td"); const nx = td && td.nextElementSibling; if (nx && nx.innerText.trim()) return nx.innerText.trim().replace(/\s+/g, " ").slice(0, 60);
        return ((e.closest("tr") || e.parentElement || {}).innerText || "").trim().replace(/\s+/g, " ").slice(0, 60); })() : ""])
    .slice(0, 260);
  log(`explore ${page.url().slice(0, 100)}: ${page.frames().length} frames`);
  for (const fr of page.frames()) {
    const controls = await fr.evaluate(dump).catch(() => null);
    log(` frame ${fr.name() || "-"} ${fr.url().slice(0, 110)}: ${controls ? controls.length + " controls" : "no access"}`);
    for (const c of controls || []) log("  " + JSON.stringify(c));
  }
  await page.screenshot({ path: path.join(ROOT, "data", "denticon-explore.png"), fullPage: true }).catch(() => {});
}
const STATUS = path.join(ROOT, "data", "denticon-robot.json");
const REPORT_PAGE = "https://a1.denticon.com/ASPX/Reports/GroupRptsDlyRpts.aspx";
const LOGIN_PAGE = "https://www.denticon.com/login";
const REPORT_NAME = "GrpExecSumm2";
const PDF_VIEWER = "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai";   // Chrome's own PDF viewer: the report is drawn as a PDF

const pad = n => String(n).padStart(2, "0");
const us = isoDay => { const [y, m, d] = isoDay.split("-"); return `${m}/${d}/${y}`; };
const log = (...a) => console.log(new Date().toISOString().slice(11, 19), ...a);

function writeStatus(patch) {
  let cur = {};
  try { cur = JSON.parse(fs.readFileSync(STATUS, "utf8")); } catch {}
  fs.writeFileSync(STATUS, JSON.stringify({ ...cur, ...patch, updated_at: new Date().toISOString() }, null, 2));
}

// True when Denticon shows a login screen, a "Session Timeout" page or a page shell with no user.
async function needsLogin(page) {
  const url = page.url().toLowerCase();
  if (url.includes("login")) return true;
  if ((await page.locator("input[type=password]").count()) > 0) return true;
  const title = await page.title().catch(() => "");
  if (/Welcome,\s*::/.test(title)) return true;
  const text = await page.locator("body").innerText().catch(() => "");
  return /Session Timeout|login session has expired/i.test(text);
}

// Show the login page and wait for the user to log in (the window is theirs; no password passes through here).
async function waitForLogin(context, page, minutes = 15) {
  writeStatus({ state: "waiting_login", needs_login: true, message: "Log into Denticon in the robot window." });
  await page.goto(LOGIN_PAGE, { waitUntil: "domcontentloaded" }).catch(() => {});
  await page.bringToFront().catch(() => {});
  const until = Date.now() + minutes * 60 * 1000;
  while (Date.now() < until) {
    await new Promise(r => setTimeout(r, 4000));
    const p = context.pages().find(x => /a1\.denticon\.com/i.test(x.url())) || page;
    const title = await p.title().catch(() => "");
    if (/a1\.denticon\.com/i.test(p.url()) && /Welcome,\s*\S/.test(title)) {
      writeStatus({ state: "running", needs_login: false, logged_in_at: new Date().toISOString(), message: "Logged in. Pulling the reports." });
      return p;
    }
  }
  throw new Error("not logged in: nobody logged into Denticon in the robot window within " + minutes + " minutes");
}

// The report is drawn when the viewer shows it as a PDF (Chrome's PDF viewer frame or an embedded PDF).
async function reportDrawn(viewer) {
  if (viewer.frames().some(f => f.url().startsWith(PDF_VIEWER))) return true;
  for (const fr of viewer.frames()) {
    const hit = await fr.evaluate(() => {
      const emb = document.querySelector("embed[type='application/pdf'], object[type='application/pdf']");
      const text = (document.body && document.body.innerText) || "";
      return !!emb || (/New Patients by First Visit Date/i.test(text) && /Production/i.test(text));
    }).catch(() => false);
    if (hit) return true;
  }
  return false;
}

// One Executive Summary for From..Thru. A one-day report (From = Thru) goes to the inbox; a longer range goes to
// data/pulls (a month-to-date snapshot to compare against the day-by-day history).
async function pullReport(context, page, frm, thru) {
  const day = frm;
  await page.goto(REPORT_PAGE, { waitUntil: "domcontentloaded", timeout: 90000 });
  if (await needsLogin(page)) throw new Error("not logged in");
  await page.click("#rbExecSumm");
  await page.waitForSelector("#chkOfficeOfficeCheckList_OfficeCheckList0", { timeout: 60000 });
  if (!(await page.isChecked("#chkOfficeOfficeCheckList_OfficeCheckList0"))) await page.check("#chkOfficeOfficeCheckList_OfficeCheckList0");
  await page.fill("#txtStartDt", us(frm));
  await page.fill("#txtEndDt", us(thru));
  // Print / Preview opens the viewer as a popup that never draws here; take its address and open it in a plain tab.
  for (const extra of context.pages()) if (extra !== page) await extra.close().catch(() => {});
  const [popup] = await Promise.all([context.waitForEvent("page", { timeout: 60000 }), page.click("#btnPrint")]);
  await popup.waitForLoadState("domcontentloaded").catch(() => {});
  for (let i = 0; i < 12 && !/r2\.denticon\.com/i.test(popup.url()); i++) await popup.waitForTimeout(1000);
  const viewerUrl = popup.url();
  await popup.close().catch(() => {});
  const frameUrl = new URL(viewerUrl);
  const rptPath = frameUrl.searchParams.get("rptPath") || "", rptArgs = frameUrl.searchParams.get("args") || "";
  if (!rptPath.endsWith(REPORT_NAME) || !rptArgs.includes(us(frm)) || !rptArgs.includes(us(thru))) throw new Error(`unexpected viewer address: ${viewerUrl.slice(0, 120)}`);
  const viewer = await context.newPage();
  await viewer.goto(viewerUrl, { waitUntil: "domcontentloaded", timeout: 120000 });
  await viewer.bringToFront().catch(() => {});
  const start = Date.now();
  while (!(await reportDrawn(viewer))) {
    if (Date.now() - start > 300000) { await viewer.screenshot({ path: path.join(ROOT, "data", "denticon-robot-viewer-fail.png") }).catch(() => {}); throw new Error("the report never finished drawing within 5 minutes"); }
    await viewer.waitForTimeout(2000);
  }
  const waited = ((Date.now() - start) / 1000).toFixed(0);
  await viewer.waitForTimeout(2000);
  // Click "Export to Excel" in the viewer's header and catch the download.
  const header = viewer.frames().find(f => /ReportHeader/i.test(f.url()));
  if (!header) throw new Error("the viewer's header (with the Export icons) was not found");
  const dir = frm === thru ? INBOX : PULLS;
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, frm === thru ? `Denticon_${REPORT_NAME} ${day}.xlsx` : `Denticon_${REPORT_NAME} ${frm} to ${thru} pulled ${new Date().toISOString().slice(0, 16).replace("T", " ").replace(":", "")}.xlsx`);
  const [download] = await Promise.all([viewer.waitForEvent("download", { timeout: 120000 }), header.click("#ExcelIconImg")]);
  await download.saveAs(file + ".part");
  const body = fs.readFileSync(file + ".part");
  const complete = body.length > 20000 && body.includes("xl/worksheets/sheet2.xml");
  log(`report for ${frm === thru ? day : frm + " to " + thru}: drawn after ${waited} s, export ${body.length} bytes${complete ? "" : " (incomplete)"}`);
  if (!complete) { fs.unlinkSync(file + ".part"); throw new Error(`the export is not the full report (${body.length} bytes)`); }
  fs.renameSync(file + ".part", file);
  await viewer.close().catch(() => {});
  return file;
}

const pullDay = (context, page, day) => pullReport(context, page, day, day);

async function pullDays(context, page, days, headed) {
  const done = [], failed = [];
  if (fs.existsSync(EXPLORE_MARK)) {
    const spec = JSON.parse(fs.readFileSync(EXPLORE_MARK, "utf8")); fs.unlinkSync(EXPLORE_MARK);
    try { await explore(context, page, spec); writeStatus({ state: "done", message: "Explored " + spec.url, days: [], failed: [] }); }
    catch (e) { log("explore failed", e.message); writeStatus({ state: "error", message: "Explore failed: " + e.message, days: [], failed: [] }); }
    return true;
  }
  // A range marker turns the first job into one month-to-date report From..Thru (then the days follow if asked).
  let jobs = days.map(d => [d, d]);
  if (fs.existsSync(RANGE_MARK)) {
    const how = fs.readFileSync(RANGE_MARK, "utf8").trim(); fs.unlinkSync(RANGE_MARK);
    jobs = [[days[0], days[days.length - 1]], ...(how === "range+days" ? jobs : [])];
  }
  let which = "exec";
  if (fs.existsSync(REPORT_MARK)) { which = fs.readFileSync(REPORT_MARK, "utf8").trim(); fs.unlinkSync(REPORT_MARK); }
  const all = { exec: ["Executive Summary", pullReport], referral: ["Referral Production", pullReferral], ortho: ["Ortho Daily Journal", pullOrthoJournal], txplan: ["Treatment Plan Status", pullTxPlanStatus] };
  const runners = (which === "both" ? ["exec", "referral"] : which === "all" ? ["exec", "referral", "ortho", "txplan"] : [which in all ? which : "exec"]).map(k => all[k]);
  jobs = jobs.flatMap(j => runners.map(r => [...j, r]));
  const names = runners.map(r => r[0]).join(" + ");
  writeStatus({ state: "running", message: `Pulling ${jobs.length} report${jobs.length === 1 ? "" : "s"} (${names}).`, days: [], failed: [], started_at: new Date().toISOString() });
  const STOP = path.join(ROOT, "data", "denticon-robot.stop");   // present = stop after the report in progress
  if (fs.existsSync(STOP)) fs.unlinkSync(STOP);
  for (const [frm, thru, [rname, run]] of jobs) {
    if (fs.existsSync(STOP)) { fs.unlinkSync(STOP); log("stopped on request"); failed.push({ day: "stopped", error: "stopped on request" }); break; }
    const day = (frm === thru ? frm : `${frm} to ${thru}`) + (runners.length > 1 ? ` (${rname})` : "");
    try {
      let f;
      try { f = await run(context, page, frm, thru); }
      catch (e) {
        if (!/not logged in/.test(e.message) || !headed) throw e;
        page = await waitForLogin(context, page);
        f = await run(context, page, frm, thru);
      }
      done.push(day); log("saved", path.basename(f));
      writeStatus({ days: done, message: `Saved ${done.length} of ${jobs.length}.` });
    } catch (e) {
      let where = "";
      try { where = ` [at ${page.url().slice(0, 100)}]`; await page.screenshot({ path: path.join(ROOT, "data", "denticon-robot-last.png") }); } catch {}
      const error = e.message.split("\n")[0] + where;
      failed.push({ day, error }); log("failed", day, error);
      if (/not logged in/.test(error)) break;
    }
  }
  const needsLoginNow = failed.some(f => /not logged in/.test(f.error));
  writeStatus({ state: failed.length ? "error" : "done", last_run: new Date().toISOString(), days: done, failed, ok: failed.length === 0, needs_login: needsLoginNow,
    message: failed.length ? (needsLoginNow ? "Not logged in. Click Pull again and log in when the window opens." : `${failed.length} day${failed.length === 1 ? "" : "s"} failed: ${failed[0].error}`)
                           : `Saved ${done.length} file${done.length === 1 ? "" : "s"} to the inbox.` });
  return failed.length === 0;
}

module.exports = { pullDays, needsLogin, waitForLogin, REPORT_PAGE, LOGIN_PAGE };
