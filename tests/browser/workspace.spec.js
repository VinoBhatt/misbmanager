const { test, expect } = require('@playwright/test');

test('every fund management view renders without browser errors', async ({page}) => {
  const errors=[];
  page.on('pageerror',error=>errors.push(error.message));
  await page.goto('/');
  await expect(page.locator('#content .kpis').first()).toBeVisible();
  for(const view of ['reconciliation','matching','exposure','receivables','repayment-alerts','profit','portfolio','cash','projection','monitor','reports','statement','allocation','allocation-email','simcopy','data','audit','dashboard']) {
    await page.locator(`.nav[data-view="${view}"]`).click();
    await expect(page.locator('#content')).not.toBeEmpty();
  }
  expect(errors).toEqual([]);
  await page.screenshot({path:'.local/dashboard.png',fullPage:true});
});

test('new note allocation is saved for Simulation Copy',async({page})=>{
  await page.goto('/');
  await expect(page.locator('#content .kpis').first()).toBeVisible();
  await page.locator('.nav[data-view="allocation"]').click();
  await page.locator('[name="loan_code"]').fill('IIF-9999');
  await page.locator('[name="note_name"]').fill('Precision Tools Manufacturer 11');
  await page.locator('[name="issuer_name"]').fill('Example Manufacturing Sdn Bhd');
  await page.locator('[name="company_id"]').fill('6001');
  await page.locator('[name="rating"]').fill('CR6');
  await page.locator('[name="loan_note_size"]').fill('386000');
  await page.locator('[name="investment_amount"]').fill('63729.67');
  await page.locator('[name="gross_pa"]').fill('15.6');
  await page.locator('[name="term"]').fill('90');
  await page.locator('[name="allocation_date"]').fill('2026-09-17');
  await page.locator('[name="disbursal_date"]').fill('2026-09-25');
  await page.locator('[name="campaign_start"]').fill('2026-09-17');
  await page.locator('[name="campaign_end"]').fill('2026-09-24');
  await page.locator('[name="business_description"]').fill('Precision component manufacturing and invoice financing.');
  await page.getByRole('button',{name:'Save allocation'}).click();
  await expect(page.locator('#allocationMessage')).toContainText('IIF-9999 is ready');
  await expect(page.locator('#allocationSaved')).toContainText('Precision Tools Manufacturer 11');
  let workflow=page.locator('.update-approval[data-code="IIF-9999"]');
  await workflow.locator('xpath=..').getByLabel('Next approval status').selectOption('Ready for Approval');
  await workflow.click();
  await expect(page.locator('#allocationMessage')).toContainText('Ready for Approval');
  workflow=page.locator('.update-approval[data-code="IIF-9999"]');
  await workflow.locator('xpath=..').getByLabel('Next approval status').selectOption('Approved');
  await workflow.click();
  await expect(page.locator('#allocationMessage')).toContainText('Approved');
  await page.screenshot({path:'.local/allocation-generator.png',fullPage:true});
  await page.locator('.nav[data-view="allocation-email"]').click();
  await expect(page.locator('#allocationEmailPreview')).toContainText('Dear Muamalat Invest Operations Team');
  await expect(page.locator('#allocationEmailPreview')).toContainText('IIF9999-17092026');
  await expect(page.locator('.fund-position')).toContainText('Total expected available fund as at');
  await page.screenshot({path:'.local/allocation-email.png',fullPage:true});
  await page.locator('.nav[data-view="allocation"]').click();
  await page.locator('.delete-allocation[data-code="IIF-9999"]').click();
  await expect(page.locator('#allocationSaved')).not.toContainText('Precision Tools Manufacturer 11');
});

test('statement upload produces a downloadable PDF',async({page})=>{
  test.skip(!process.env.MISB_STATEMENT_TEST_LEDGER,'Set MISB_STATEMENT_TEST_LEDGER to test a real workbook upload');
  await page.goto('/');
  await expect(page.locator('#content .kpis').first()).toBeVisible();
  await page.locator('.nav[data-view="statement"]').click();
  await page.locator('#pdfLedgerFile').setInputFiles(process.env.MISB_STATEMENT_TEST_LEDGER);
  await page.locator('#pdfCutoff').fill('2026-09-02');
  await page.getByRole('button',{name:'Prepare statement',exact:true}).click();
  await expect(page.locator('#pdfStatementPreview')).toContainText('679 transactions', {timeout:30000});
  await expect(page.locator('#pdfStatementPreview')).toContainText('2,038,351.00');
  const downloading=page.waitForEvent('download');
  await page.getByRole('button',{name:'Download account statement PDF',exact:true}).click();
  const download=await downloading;
  expect(download.suggestedFilename()).toBe('MISB_Account_Statement_2026-09-02.pdf');
  await download.saveAs('.local/browser-statement.pdf');
  await page.screenshot({path:'.local/statement-section.png',fullPage:true});
});

test('mobile navigation and login fit the screen',async({page})=>{
  await page.setViewportSize({width:390,height:844});
  await page.goto('/');
  await expect(page.locator('#content .kpis').first()).toBeVisible();
  expect(await page.evaluate(()=>document.documentElement.scrollWidth)).toBeLessThanOrEqual(391);
  await page.locator('.nav[data-view="data"]').click();
  await expect(page.locator('.validated-upload')).toHaveCount(3);
  await page.goto('/login');
  await expect(page.getByLabel('Workspace password')).toBeVisible();
  expect(await page.evaluate(()=>document.documentElement.scrollWidth)).toBeLessThanOrEqual(391);
  await page.screenshot({path:'.local/login-mobile.png',fullPage:true});
});
