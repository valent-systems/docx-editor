import { test } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';

test('probe: natural line-height ratios for doc fonts', async ({ page }) => {
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();

  const ratios = await page.evaluate(async () => {
    // Load the families the TPX doc uses.
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href =
      'https://fonts.googleapis.com/css2?family=Onest:wght@400;600&family=Roboto:wght@300;400&display=swap';
    document.head.appendChild(link);
    await new Promise((r) => {
      link.onload = r;
      setTimeout(r, 4000);
    });
    await document.fonts.load('400 100px Onest');
    await document.fonts.load('600 100px Onest');
    await document.fonts.load('300 100px Roboto');
    await document.fonts.ready;

    const measure = (family: string, weight: number) => {
      const el = document.createElement('span');
      el.style.cssText = `font: ${weight} 100px "${family}"; line-height: normal; position: absolute; visibility: hidden;`;
      el.textContent = 'Hg';
      document.body.appendChild(el);
      const h = el.offsetHeight / 100;
      el.remove();
      return h;
    };
    return {
      onest400: measure('Onest', 400),
      onest600: measure('Onest', 600),
      roboto300: measure('Roboto', 300),
      roboto400: measure('Roboto', 400),
      loaded: {
        onest: document.fonts.check('400 16px Onest'),
        roboto: document.fonts.check('300 16px Roboto'),
      },
    };
  });
  console.log('ratios:', JSON.stringify(ratios));
});
