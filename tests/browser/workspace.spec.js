const { test, expect } = require('@playwright/test');

test('every fund management view renders without browser errors', async ({page}) => {
  const errors=[];
  page.on('pageerror',error=>errors.push(error.message));
  await page.goto('/');
  await expect(page.locator('#content .kpis').first()).toBeVisible();
  for(const view of ['exposure','receivables','profit','portfolio','cash','projection','monitor','reports','simcopy','data','dashboard']) {
    await page.locator(`.nav[data-view="${view}"]`).click();
    await expect(page.locator('#content')).not.toBeEmpty();
  }
  expect(errors).toEqual([]);
  await page.screenshot({path:'.local/dashboard.png',fullPage:true});
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
