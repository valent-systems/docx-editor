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
      'https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400&family=Source+Sans+Pro:wght@400&family=Bebas+Neue&display=swap';
    document.head.appendChild(link);
    await new Promise((r) => {
      link.onload = r;
      setTimeout(r, 4000);
    });
    await document.fonts.load('400 100px "Source Sans 3"');
    await document.fonts.load('400 100px "Source Sans Pro"');
    await document.fonts.load('400 100px "Bebas Neue"');
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
      sourceSans3: measure('Source Sans 3', 400),
      sourceSansPro: measure('Source Sans Pro', 400),
      bebasNeue: measure('Bebas Neue', 400),
      loaded: {
        ss3: document.fonts.check('400 16px "Source Sans 3"'),
        ssp: document.fonts.check('400 16px "Source Sans Pro"'),
        bebas: document.fonts.check('400 16px "Bebas Neue"'),
      },
    };
  });
  console.log('ratios:', JSON.stringify(ratios));
});
