import { test, expect } from '@playwright/test';

test('loading status stays in the accessibility tree between announcements', async ({ page }) => {
  await page.goto('/tests/visual/loading-status-fixture.html');
  const status = page.getByRole('status');
  // getByRole excludes nodes hidden from accessibility by display:none or an
  // ancestor; the empty live region must exist before the first text mutation.
  await expect(status).toHaveCount(1);
  await expect(status).toHaveText('');
  await expect(status).toHaveAttribute('aria-live', 'polite');
  await expect(status).toHaveAttribute('aria-atomic', 'true');
  for (let episode = 0; episode < 2; episode++) {
    await page.getByRole('button', { name: 'Start loading' }).click();
    await expect(status).toHaveText('Loading page');
    await expect(status).toMatchAriaSnapshot('- status: Loading page');
    await page.getByRole('button', { name: 'Finish loading' }).click();
    await expect(status).toHaveCount(1);
    await expect(status).toHaveText('');
  }
});
