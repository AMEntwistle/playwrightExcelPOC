const { test, expect } = require('@playwright/test');
const path = require("node:path");

// Fill these in
const OFFICE_ACCOUNT = ''
const OFFICE_PASSWORD = ''
const XLSX_FILE = ''

async function createLocator(page, selector)  {
    return page.locator(selector)
}

test('can login to microsoft office 365', async ({ page, context }) => {
    test.setTimeout(120000)
    await page.goto('https://www.office.com/login?ru=%2f')
    const email = await createLocator(page, 'input[type="email"]')
    await email.fill(OFFICE_ACCOUNT)
    let submit = await createLocator(page,'input[type="submit"]')
    await submit.click()
    const password = await createLocator(page, 'input[type="password"]')
    await password.click()
    await password.fill(OFFICE_PASSWORD)
    submit = await createLocator(page,'button[type="submit"]')
    await submit.click()
    const decline = await createLocator(page, '#declineButton')
    await decline.click()
    const upload = await createLocator(page, '#upload-button-QuickAccess')
        const fileChooserPromise = page.waitForEvent('filechooser')
    const pagePromise = context.waitForEvent('page')
    await upload.click()
    const fileChooser = await fileChooserPromise
    await fileChooser.setFiles(path.join(__dirname, XLSX_FILE))
    const excelPage = await pagePromise
    console.log(`url is ${excelPage.url()}`)
    const excelSheet = await excelPage.locator('iframe[name="WacFrame_Excel_0"]').contentFrame().locator('.ewr-sheettable')
    await expect(excelSheet.first()).toBeVisible()
    await expect(await excelSheet.count()).toBeGreaterThan(1)
    const alert = await excelPage.locator('iframe[name="WacFrame_Excel_0"]').contentFrame().locator('#BusinessBarRegion [role="alert"]')
    await expect(alert).toHaveCount(0)
})