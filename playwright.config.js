const { defineConfig } = require('@playwright/test');
module.exports = defineConfig({
  testDir: './tests/browser',
  use: {baseURL:'http://127.0.0.1:5056',channel:'msedge'},
  reporter: 'list',
  webServer: {command:'python scripts/local.py', url:'http://127.0.0.1:5056', reuseExistingServer:true, timeout:30000}
});
