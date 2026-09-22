const { test, expect } = require('@playwright/test');

test('every fund management view renders without browser errors', async ({page}) => {
  const errors=[];
  page.on('pageerror',error=>errors.push(error.message));
  await page.goto('/');
  await expect(page.locator('#content .kpis').first()).toBeVisible();
  for(const view of ['exposure','receivables','profit','portfolio','cash','projection','monitor','reports','statement','simcopy','data','dashboard']) {
    await page.locator(`.nav[data-view="${view}"]`).click();
    await expect(page.locator('#content')).not.toBeEmpty();
  }
  expect(errors).toEqual([]);
  await page.screenshot({path:'.local/dashboard.png',fullPage:true});
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
  await expect(page.locator('.upload')).toHaveCount(3);
  await page.goto('/login');
  await expect(page.getByLabel('Workspace password')).toBeVisible();
  expect(await page.evaluate(()=>document.documentElement.scrollWidth)).toBeLessThanOrEqual(391);
  await page.screenshot({path:'.local/login-mobile.png',fullPage:true});
});
