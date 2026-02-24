import { log } from "console";
import { exec } from "child_process";
import { Browser, chromium, Page } from "playwright";
import * as net from "net";
import * as fs from "fs";
import * as readline from "readline";
// @ts-ignore
import thaiIdCard from "thai-id-card";
import { thFaker } from "./thFaker";

function promptInput(question: string): Promise<string> {
   return new Promise((resolve) => {
      const rl = readline.createInterface({
         input: process.stdin,
         output: process.stdout,
      });
      rl.question(question, (answer) => {
         rl.close();
         resolve(answer.trim());
      });
   });
}

function waitForPort(
   port: number,
   host: string,
   timeout = 15000,
): Promise<void> {
   return new Promise((resolve, reject) => {
      const start = Date.now();
      const tryConnect = () => {
         const socket = new net.Socket();
         socket.once("connect", () => {
            socket.destroy();
            resolve();
         });
         socket.once("error", () => {
            socket.destroy();
            if (Date.now() - start > timeout) {
               reject(new Error(`Timed out waiting for port ${port}`));
            } else {
               setTimeout(tryConnect, 500);
            }
         });
         socket.connect(port, host);
      };
      tryConnect();
   });
}

function isPortOpen(port: number, host: string): Promise<boolean> {
   return new Promise((resolve) => {
      const socket = new net.Socket();
      socket.once("connect", () => {
         socket.destroy();
         resolve(true);
      });
      socket.once("error", () => {
         socket.destroy();
         resolve(false);
      });
      socket.connect(port, host);
   });
}

interface OrderData {
   thaiId: string;
   firstName: string;
   lastName: string;
   fullName: string;
   tel: string;
}

function generateOrderData(): OrderData {
   const thaiId = thaiIdCard.generate();
   const firstName = thFaker.person.firstName();
   const lastName = thFaker.person.lastName();
   return {
      thaiId,
      firstName,
      lastName,
      fullName: `${firstName} ${lastName}`,
      tel: "0661128806",
   };
}

async function createOrder(page: Page, data: OrderData) {
   if (page.isClosed()) {
      page = await page.context().newPage();
   }
   try {
      await page.goto("https://prmorder.uatsiamsmile.com/premiumnoticepages");
   } catch (err) {
      // if frame detached, recreate page and retry
      const e: any = err;
      if (e && e.message && e.message.includes("Frame has been detached")) {
         page = await page.context().newPage();
         await page.goto(
            "https://prmorder.uatsiamsmile.com/premiumnoticepages",
         );
      } else {
         throw err;
      }
   }
   await page.waitForLoadState("networkidle");
   log("Navigated to order page, waiting for content to load...");

   await page.fill('input[name="cardDetail"]', data.thaiId);
   await page.waitForTimeout(800);
   await page.click('div[role="button"][aria-haspopup="listbox"]');
   await page.waitForTimeout(800);
   await page.fill('input[name="firstName"]', data.firstName);
   await page.waitForTimeout(800);
   await page.fill('input[name="lastName"]', data.lastName);
   await page.waitForTimeout(800);
   await page.fill('input[name="phoneNumber"]', data.tel);
   await page
      .locator('button[type="submit"]', { hasText: "ประกันสุขภาพ" })
      .click();
   await page.waitForTimeout(800);
   await page
      .locator('div[role="button"]', { hasText: "เลือกแผนประกัน" })
      .click();
   await page.waitForTimeout(800);
   await page.click('li[data-value="56"]');
   await page.waitForTimeout(800);
   await page.locator('div[role="button"]', { hasText: "เลือกประเภท" }).click();
   await page.waitForTimeout(800);
   await page.click('li[data-value="2"]');
   await page.waitForTimeout(800);
   await page.fill('textarea[name="customerName"]', "ทดสอบ ระบบ");
   await page.locator('button[type="submit"]', { hasText: "ยืนยัน" }).click();
   await page.waitForTimeout(800);
   await page.locator('button[type="submit"]', { hasText: "ชำระเงิน" }).click();
   await page.waitForTimeout(800);
   await page.locator('button[type="button"]', { hasText: "ยืนยัน" }).click();
   await page.waitForTimeout(800);
   await page.click('input[name="rdoPayment"][value="1"]');
   await page.waitForTimeout(800);
   await page.locator('button[type="button"]', { hasText: "ยืนยัน" }).click();
   await page.waitForTimeout(800);
   await page.locator("button.swal2-confirm", { hasText: "ยืนยัน" }).click();
   await page.waitForTimeout(800);
   await page.locator("button.swal2-confirm", { hasText: "ยืนยัน" }).click();
   await page.waitForTimeout(800);
   log("Order process completed!");
}

async function processPayment(page: Page, fullName: string) {
   await page.goto("https://prmorder.uatsiamsmile.com/premiummanagementpages");
   await page.waitForTimeout(800);
   // search for tr with td that has text fullName
   const orderRow = page.locator("tr", {
      has: page.locator("td", { hasText: fullName }),
   });
   // click actions button in the same row
   await orderRow.locator('button[type="button"]', { hasText: "..." }).click();
   await page.waitForTimeout(800);
   await page
      .getByRole("button", { name: "จ่ายเงิน (สำหรับ SIT เท่านั้น)" })
      .click();
   await page.locator("button.swal2-confirm", { hasText: "ยืนยัน" }).click();
   await page.waitForTimeout(2000);
   log("Payment confirmed for order with customer:", fullName);
}

async function openSSSPage(page: Page, fullName: string) {
   // open sss
   // redirect to https://prmorder.uatsiamsmile.com/premiummanagementpages if not already there
   if (!page.url().includes("premiummanagementpages")) {
      await page.goto(
         "https://prmorder.uatsiamsmile.com/premiummanagementpages",
      );
      await page.waitForLoadState("networkidle");
   }
   await page.waitForTimeout(800);
   await page.getByRole("tab", { name: "รับชำระแล้ว" }).click();
   // search for tr with td that has text fullName
   const paidOrderRow = page.locator("tr", {
      has: page.locator("td", { hasText: fullName }),
   });
   log("Found paid order row for customer:", fullName);
   // log first td in the row
   await paidOrderRow.locator("td").nth(0).click();
   await page.waitForTimeout(800);
   // The nested "table" is a <div role="table">, not a <table> element
   // It's inside a <td colspan="12"> that only appears when expanded
   const nestedTable = page.locator('td[colspan="12"] div[role="table"]');
   // click "..." button in the nested table's data row
   const nestedRow = nestedTable.locator("tbody tr").first();
   await nestedRow.locator('button[type="button"]', { hasText: "..." }).click();
   await page.getByRole("button", { name: "เปิดหน้าแอพ" }).click();
   await page.waitForTimeout(800);
   log("Opened SSS app for customer:", fullName);
}

async function generateAppId(sssPage: Page): Promise<string> {
   const appIdInput =
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$txtAppID"]';
   const checkAppBtn =
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$btnCheckDuplicate"]';

   while (true) {
      const appId = String(Math.floor(1000000 + Math.random() * 9000000)); // 7-digit random
      log("Trying App ID:", appId);
      await sssPage.fill(appIdInput, appId);
      await sssPage.click(checkAppBtn);
      await sssPage.waitForTimeout(1000);

      const bodyText = await sssPage.locator("body").innerText();
      if (bodyText.includes("สามารถใช้เลข App นี้ได้")) {
         log("App ID accepted:", appId);
         return appId;
      }
      if (bodyText.includes("เลข App ซ้ำ")) {
         log("App ID duplicated, retrying...");
         continue;
      }
      // fallback: if neither message found, retry
      log("Unexpected response, retrying...");
   }
}

async function processSSSApp(
   browser: Browser,
   data: OrderData,
   appId: string,
   dcrMonth: string,
   dcrYear: string,
   payerRelation: string,
) {
   // get sss page
   const pages = browser.contexts()[0].pages();
   const sssPage = pages[pages.length - 1];
   await sssPage.waitForLoadState("networkidle");
   log("SSS app page loaded for customer:", data.fullName);
   log("SSS app URL:", sssPage.url());

   // Tab 1: ข้อมูล Application
   // fill เลขที่ Application
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$txtAppID"]',
      appId,
   );

   // select เดือน DCR (Month Period) - ค่าที่ผู้ใช้เลือก
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$ucMonthYearPeriod$ddlMonth"]',
      dcrMonth,
   );

   // select ปี DCR (Year Period) - ค่าที่ผู้ใช้เลือก
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$ucMonthYearPeriod$ddlYear"]',
      dcrYear,
   );
   await sssPage.waitForTimeout(800);

   // select คำนำหน้า (Title) - "คุณ" = value "1001"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$ddlTitle"]',
      "1001",
   );

   // fill ชื่อ (First Name)
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$txtFirstName"]',
      data.firstName,
   );

   // fill สกุล (Last Name)
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$txtLastName"]',
      data.lastName,
   );

   // fill วันเกิด (Birth Date) - format dd/mm/yyyy (Buddhist era)
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$ucBirthdate$txtDate"]',
      "01/01/2546",
   );

   await sssPage.click("button.ui-datepicker-close");
   await sssPage.waitForTimeout(800);

   // select รถ Zebra Car - "000 (สำนักงาน -)"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$ddlZebraCar"]',
      "ZB6603000003",
   );

   // select ช่วงเวลา - "ช่วงเช้า (9:00 - 12:00)"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$ucApplicationDetail1$ucUDWDay1$ddlTimePeriod"]',
      "6901",
   );

   await sssPage.waitForTimeout(800);

   sssPage.waitForTimeout(800);

   // then click : <input type="submit" name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$btnNextAppDetail" value="ถัดไป &gt;" onclick=" this.disabled = true; __doPostBack('ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$btnNextAppDetail','');" id="ContentPlaceHolder1_TabContainer1_tabApplicationDetail_btnNextAppDetail" style="background-color:#99FF99;height:50px;width:150px;">
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabApplicationDetail$btnNextAppDetail"]',
   );
   await sssPage.waitForTimeout(800);
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$ucConfirmDialog$btnOK"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // Tab 2: ข้อมูลผู้เอาประกัน
   // fill เลขบัตรประชาชน (Thai ID) : <input name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$ucZCardID1$txtCardID" type="text" value="9736520" maxlength="13" id="ContentPlaceHolder1_TabContainer1_tabCustomerDetail_ucCustomerDetail1_ucZCardID1_txtCardID">
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$ucZCardID1$txtCardID"]',
      data.thaiId,
   );

   // select อาชีพ - "เกษตรกร"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$ucOccupation1$ddlOccupation"]',
      "1000",
   );

   // select สถานภาพ - "โสด"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$ddlMaritalStatus"]',
      "5001",
   );

   // select กรุ๊ปเลือด - "A"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$ddlBloodType"]',
      "4001",
   );

   // select เพศ - "ชาย"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$ddlSex"]',
      "1001",
   );

   // fill น้ำหนัก
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$txtWeight"]',
      "60",
   );

   // fill ส่วนสูง
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$txtHeight"]',
      "168",
   );

   // fill โทรศัพท์บ้าน
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$txtHomePhone"]',
      "-",
   );

   // fill โทรศัพท์ที่ทำงาน
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$txtWorkPhone"]',
      "-",
   );

   // fill โทรศัพท์มือถือ
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$txtMobilePhone"]',
      data.tel,
   );

   // fill Email
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$txtEmail"]',
      "-",
   );

   // ที่อยู่ตามบัตรประชาชน
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$txtNo"]',
      "-",
   );
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$txtVillageName"]',
      "-",
   );
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$txtMoo"]',
      "-",
   );
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$txtFloor"]',
      "-",
   );
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$txtSoi"]',
      "-",
   );
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$txtRoad"]',
      "-",
   );

   // จังหวัด → อำเภอ → ตำบล (cascade dropdowns with postback)
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$ucProvinceAmphoeTumbol$ddlProvince"]',
      "50", // เชียงใหม่
   );
   await sssPage.waitForLoadState("networkidle");

   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$ucProvinceAmphoeTumbol$ddlAmphoe"]',
      "523", // เมืองเชียงใหม่
   );
   await sssPage.waitForLoadState("networkidle");

   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$ucProvinceAmphoeTumbol$ddlTumbol"]',
      "ช้างเผือก", // ช้างเผือก
   );
   await sssPage.waitForLoadState("networkidle");

   // เบอร์โทรศัพท์ (1)
   await sssPage.fill(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel1$ucHomeAddress$txtPhoneNo1"]',
      data.tel,
   );

   // click tab "ที่อยู่ที่ทำงานของผู้เอาประกัน"
   await sssPage.click(
      "#__tab_ContentPlaceHolder1_TabContainer1_tabCustomerDetail_ucCustomerDetail1_tabContainMain_TabPanel2",
   );
   await sssPage.waitForLoadState("networkidle");

   // select radio "ที่อยู่ที่ทำงานเหมือนที่อยู่บ้าน"
   await sssPage.check(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel2$rdbWorkAddress"][value="rdbWorkAddressSameHome"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // click tab "ที่อยู่ที่สามารถติดต่อได้ (ที่อยู่ที่ส่งเอกสาร)ของผู้เอาประกัน"
   await sssPage.click(
      "#__tab_ContentPlaceHolder1_TabContainer1_tabCustomerDetail_ucCustomerDetail1_tabContainMain_TabPanel3",
   );
   await sssPage.waitForLoadState("networkidle");

   // select radio "ที่อยู่ที่ติดต่อได้เหมือนที่อยู่บ้าน"
   await sssPage.check(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$ucCustomerDetail1$tabContainMain$TabPanel3$rdbContactAddress"][value="rdbContactAddressSameHome"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // click "ถัดไป >" button
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabCustomerDetail$btnNextCustDetail"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // select ความสัมพันธ์ผู้ชำระเงิน ตามที่ผู้ใช้ป้อน
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ddlPayerRelationShip"]',
      payerRelation || "3000",
   );
   await sssPage.waitForLoadState("networkidle");

   // หากไม่ใช่ตัวเอง ให้กรอกข้อมูลสุ่มเพื่อเลี่ยงช่องว่าง
   if ((payerRelation || "3000") !== "3000") {
      await sssPage.check(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$rdbCard"][value="rdbZCardID"]',
      );

      await sssPage.fill(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$ucZCardIDPayer$txtCardID"]',
         thaiIdCard.generate(),
      );

      await sssPage.selectOption(
         'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$ddlTitle"]',
         "1001",
      );
      await sssPage.fill(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$txtFirstName"]',
         thFaker.person.firstName(),
      );
      await sssPage.fill(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$txtLastName"]',
         thFaker.person.lastName(),
      );
      await sssPage.selectOption(
         'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$ddlOccupation"]',
         "1000", // เกษตรกร
      );

      await sssPage.selectOption(
         'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$ddlOccupationLevel"]',
         "1001", // เกษตรกร
      );

      await sssPage.fill(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$txtEmail"]',
         "-",
      );
      await sssPage.fill(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$txtPhoneNumber"]',
         data.tel,
      );

      await sssPage.click(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$btnCopyContactAddress"]',
      );

      await sssPage.click(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$ucPayer1$btnCopyWorkAddress"]',
      );

      await sssPage.waitForLoadState("networkidle");
   }

   // click "ถัดไป >" button on payer detail tab
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayerDetail$btnNextPayer"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // select วิธีการชำระเงิน - "เงินสด"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayPremiumPayer$ucPayerPayPremium$ddlPayMethod"]',
      "9001",
   );
   await sssPage.waitForLoadState("networkidle");

   // click "ถัดไป >" button on pay premium payer tab
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabPayPremiumPayer$btnNextPayer"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // click "เพิ่ม" button on heir detail tab
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabHeirDetail$ucHeir1$btnAdd"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // select ความสัมพันธ์ผู้รับผลประโยชน์ - "เป็นบุคคลเดียวกันกับผู้เอาประกัน"
   await sssPage.selectOption(
      'select[name="ctl00$ContentPlaceHolder1$TabContainer1$tabHeirDetail$ucHeir1$ddlRelation"]',
      "3000",
   );
   await sssPage.waitForLoadState("networkidle");

   // click "บันทึก" button on heir detail tab
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabHeirDetail$ucHeir1$btnSave"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // click "ถัดไป >" button on heir detail tab
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabHeirDetail$btnNextHeir"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // click "ลูกค้าสุขภาพดี" button on underwrite tab
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabUnderwrite$btnGoodHealthAgent"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // select radio "ผ่าน (rdbPass)" on underwrite tab
   await sssPage.check(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabUnderwrite$ucPHUnderwriteFromAgent$Result"][value="rdbPass"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // click "ถัดไป >" button on underwrite tab
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabUnderwrite$btnNextUnderwrite"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // select radio "ยินยอมเปิดเผยข้อมูลโรค (rdoConsentDiseClose)" on memo tab
   await sssPage.check(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabMemo$ucPHConsentDiseClose$Consent"][value="rdoConsentDiseClose"]',
   );
   await sssPage.waitForLoadState("networkidle");

   // click "ข้อมูล PDPA" button - opens new tab
   const [pdpaPage] = await Promise.all([
      browser.contexts()[0].waitForEvent("page"),
      sssPage.click(
         'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabMemo$ucPDPAData$btnPDPAData"]',
      ),
   ]);
   await pdpaPage.waitForLoadState("networkidle");

   // check all checkboxes on PDPA page
   const checkboxes = await pdpaPage.locator('input[type="checkbox"]').all();
   for (const checkbox of checkboxes) {
      if (!(await checkbox.isChecked())) {
         await checkbox.check();
      }
   }

   // click "บันทึก" button on PDPA page
   await pdpaPage.click("button#btn_save");
   await pdpaPage.waitForTimeout(800);
   await pdpaPage.click("button.confirm");
   await pdpaPage.waitForTimeout(800);
   await pdpaPage.close();
   await sssPage.waitForLoadState("networkidle");

   // click "บันทึกและส่ง MO ทันที" button on memo tab
   await sssPage.click(
      'input[name="ctl00$ContentPlaceHolder1$TabContainer1$tabMemo$btnFinish_SendMO"]',
   );
   await sssPage.waitForLoadState("networkidle");

   log("SSS form filled for customer:", data.fullName);
}

function appendSuccessLog(data: OrderData, appId: string) {
   const timestamp = new Date().toLocaleString("th-TH", {
      timeZone: "Asia/Bangkok",
   });
   const logLine = `[${timestamp}] สำเร็จ | ชื่อ: ${data.fullName} | เลขบัตรประชาชน: ${data.thaiId} | App ID: ${appId}\n`;
   fs.appendFileSync("success-log.txt", logLine, "utf-8");
   log("บันทึก log สำเร็จลงไฟล์ success-log.txt");
}

async function main() {
   // prompt user for DCR and run parameters before opening browser
   const currentMonth = String(new Date().getMonth() + 1);
   const monthNames: Record<string, string> = {
      "1": "ม.ค.",
      "2": "ก.พ.",
      "3": "มี.ค.",
      "4": "เม.ย.",
      "5": "พ.ค.",
      "6": "มิ.ย.",
      "7": "ก.ค.",
      "8": "ส.ค.",
      "9": "ก.ย.",
      "10": "ต.ค.",
      "11": "พ.ย.",
      "12": "ธ.ค.",
   };
   log("\n=== เลือกเดือน DCR ===");
   log("1=ม.ค.  2=ก.พ.  3=มี.ค.  4=เม.ย.  5=พ.ค.  6=มิ.ย.");
   log("7=ก.ค.  8=ส.ค.  9=ก.ย.  10=ต.ค.  11=พ.ย. 12=ธ.ค.");
   let dcrMonth = currentMonth;
   while (true) {
      const monthInput = await promptInput(
         `กรุณากรอกเลขเดือน (1-12, Enter=${currentMonth}): `,
      );
      if (monthInput === "") {
         break;
      }
      if (monthNames[monthInput]) {
         dcrMonth = monthInput;
         break;
      }
      log(`ค่าไม่ถูกต้อง: "${monthInput}" กรุณากรอกใหม่`);
   }
   log(`เลือกเดือน: ${monthNames[dcrMonth]} (${dcrMonth})\n`);

   const defaultYear = String(new Date().getFullYear() + 543);
   let dcrYear = "";
   while (true) {
      const dcrYearInput = await promptInput(
         `กรุณากรอกปี พ.ศ. DCR (Enter เพื่อใช้ค่าปัจจุบัน ${defaultYear}): `,
      );
      if (dcrYearInput === "") {
         dcrYear = defaultYear;
         break;
      } else if (/^\d{4}$/.test(dcrYearInput)) {
         dcrYear = dcrYearInput;
         break;
      } else {
         log(
            `ค่าไม่ถูกต้อง: "${dcrYearInput}" กรุณากรอกเป็นตัวเลข พ.ศ. เช่น ${defaultYear}`,
         );
      }
   }
   log(`เลือกปี: ${dcrYear}\n`);

   let runTimes = 1;
   const timesInput = await promptInput(
      "กรุณากรอกจำนวนครั้งที่ต้องการทำ (Enter=1): ",
   );
   if (/^\d+$/.test(timesInput) && parseInt(timesInput) >= 1) {
      runTimes = parseInt(timesInput);
   } else if (timesInput !== "") {
      log(`ค่าไม่ถูกต้อง: "${timesInput}" ใช้ค่าเริ่มต้น 1`);
   }
   log(`จะทำจำนวน ${runTimes} ครั้ง\n`);

   let payerRelation = "3000";
   let payerRelationLabel = "เป็นบุคคลเดียวกันกับผู้เอาประกัน";
   const payerInput = await promptInput(
      "ผู้ชำระเบี้ยเป็นบุคคลเดียวกันกับผู้เอาประกัน? (y=ใช่, n=ไม่ใช่, Enter=ใช่): ",
   );
   if (payerInput.toLowerCase() === "n") {
      payerRelation = "3042";
      payerRelationLabel = "เพื่อน";
   }
   log(`เลือกความสัมพันธ์: ${payerRelationLabel}\n`);

   const port = 9222;
   const host = "127.0.0.1";

   // Check if Edge is already running with debugging port
   const alreadyRunning = await isPortOpen(port, host);

   if (alreadyRunning) {
      log("Edge already running on port 9222, connecting...");
   } else {
      log("Edge not running on port 9222, launching...");
      const command = `"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe" --remote-debugging-port=${port}`;
      exec(command, (error, stdout, stderr) => {
         if (error) {
            console.error(`Error executing command: ${error}`);
            return;
         }
         console.log(`stdout: ${stdout}`);
      });

      log("Waiting for Edge to start on port 9222...");
      await waitForPort(port, host);
      log("Edge is ready!");
   }

   const browser = await chromium.connectOverCDP("http://127.0.0.1:9222");

   const context = browser.contexts()[0];
   const page = context.pages()[0] || (await context.newPage());

   await page.bringToFront();

   // Navigate to the desired URL
   // await page.goto("https://prmorder.uatsiamsmile.com");
   // await page.waitForLoadState("networkidle");
   // // if โดน redirect ไปหน้า login รอให้ login เสร็จแล้วค่อยทำงานต่อ
   // if (page.url().includes("authlogin")) {
   //    log("Please log in to the website...");
   //    await page.waitForURL("https://prmorder.uatsiamsmile.com/*");
   //    log("Login successful, continuing...");
   // }
   // log("Current URL:", page.url());

   for (let i = 1; i <= runTimes; i++) {
      log(`\n===== ครั้งที่ ${i}/${runTimes} =====`);
      // ensure we have a live page
      let page = context.pages()[0] || (await context.newPage());
      if (page.isClosed()) {
         page = await context.newPage();
      }
      await page.bringToFront();
      const data = generateOrderData();
      await createOrder(page, data);
      await processPayment(page, data.fullName);

      // await openSSSPage(browser, page, "ทดสอบ ระบบ");
      // page.goto(
      //    "http://uat.siamsmile.co.th:9157/Modules/PH/frmPHNewApp1.aspx?IGCode=SUdOVzY5MDIwMDA1Njg=",
      // );
      await openSSSPage(page, data.fullName);
      const sssPages = browser.contexts()[0].pages();
      const sssPage = sssPages[sssPages.length - 1];
      await sssPage.waitForLoadState("networkidle");
      // check is sssPage url includes "frmPHNewApp1.aspx"
      if (!sssPage.url().includes("frmPHNewApp1.aspx")) {
         log("Error: SSS app page did not open correctly.");
         return;
      }
      const appId = await generateAppId(sssPage);
      await processSSSApp(
         browser,
         data,
         appId,
         dcrMonth,
         dcrYear,
         payerRelation,
      );

      appendSuccessLog(data, appId);
      log(`===== สำเร็จครั้งที่ ${i}/${runTimes} =====\n`);
   }

   log(`\nเสร็จสิ้นทั้งหมด ${runTimes} ครั้ง`);
}
main();
