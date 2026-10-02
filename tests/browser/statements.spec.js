const {test,expect}=require('@playwright/test');
const fs=require('fs');
const path=require('path');

test('prepared monthly CSV previews and downloads without replacing the ledger',async({page})=>{
  const csv=path.resolve('artifacts/MISB_Account_Statement_2026-09-01_to_2026-09-30.csv');
  test.skip(!fs.existsSync(csv),'Requires the September statement validation artifact');
  const errors=[];
  page.on('pageerror',error=>errors.push(error.message));
  await page.goto('/');
  await expect(page.locator('#content .kpis').first()).toBeVisible();
  await page.locator('.nav[data-view="statement"]').click();
  await page.locator('#pdfLedgerFile').setInputFiles(csv);
  await page.getByRole('button',{name:'Prepare statement',exact:true}).click();
  await expect(page.locator('#pdfStatementPreview')).toContainText('120 transactions');
  await expect(page.locator('#pdfStatementPreview')).toContainText('664,448.11');
  for(const [label,format] of [['Download CSV','csv'],['Download account statement PDF','pdf']]){
    const downloading=page.waitForEvent('download');
    await page.getByRole('button',{name:label,exact:true}).click();
    const download=await downloading;
    expect(download.suggestedFilename()).toBe(`MISB_Account_Statement_2026-09-01_to_2026-09-30.${format}`);
  }
  expect(errors).toEqual([]);
});
