const {test,expect}=require('@playwright/test');

test('allocation email monitor filters pending notes and saves sent dates',async({page})=>{
  const row={loan_code:'IIF-1234',issuer:'Example issuer',note_name:'Example note',amount:2500,
    allocation_dates:['2026-01-01','2026-01-02'],transaction_count:2,allocation_email_date:'2026-01-02',
    disbursement_email_date:'',disbursal_date:'2026-01-03',ledger_signature:'current-ledger',
    source:'Imported simulation',pending:['Disbursement email pending'],emails_complete:false};
  await page.route('**/api/allocation-email-tracking',route=>route.fulfill({json:{rows:[row],unmatched:[],
    summary:{total:1,emails_complete:row.emails_complete?1:0,pending:row.emails_complete?0:1,missing_disbursal:0}}}));
  await page.route('**/api/allocation-email-tracking/IIF-1234',async route=>{
    expect(route.request().method()).toBe('PUT');
    expect(route.request().postDataJSON()).toEqual({allocation_email_date:'2026-01-02',disbursement_email_date:'2026-01-03',ledger_signature:'current-ledger'});
    Object.assign(row,{disbursement_email_date:'2026-01-03',pending:[],emails_complete:true,source:'Saved confirmation'});
    await route.fulfill({json:{ok:true}});
  });
  await page.goto('/');
  await page.locator('.nav[data-view="email-tracking"]').click();
  await expect(page.getByRole('heading',{name:'Allocation email register'})).toBeVisible();
  await expect(page.locator('#emailMonitorRows')).toContainText('Disbursement email pending');
  await page.getByLabel('Disbursement email sent for IIF-1234').fill('2026-01-03');
  await page.getByRole('button',{name:'Save / confirm'}).click();
  await expect(page.locator('#emailMonitorMessage')).toContainText('email dates saved');
  await expect(page.locator('#emailMonitorRows')).toContainText('No allocations match');
  await page.getByLabel('Filter allocation emails').selectOption('complete');
  await expect(page.locator('#emailMonitorRows')).toContainText('Complete');
  await expect(page.getByLabel('Disbursement email sent for IIF-1234')).toHaveValue('2026-01-03');
  await page.getByLabel('Search allocation emails').fill('not found');
  await expect(page.locator('#emailMonitorRows')).toContainText('No allocations match');
});
