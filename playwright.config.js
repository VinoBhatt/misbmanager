const { defineConfig } = require('@playwright/test');
const port=process.env.MISB_TEST_PORT||'5057';
module.exports = defineConfig({
  timeout:60000,
  expect:{timeout:20000},
  testDir: './tests/browser',
  use: {baseURL:`http://127.0.0.1:${port}`,channel:'msedge'},
  reporter: 'list',
  webServer: {command:`uv run python scripts/local.py --port ${port}`, url:`http://127.0.0.1:${port}`, reuseExistingServer:true, timeout:30000}
});
